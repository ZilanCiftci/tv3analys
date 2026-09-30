# -*- coding: utf-8 -*-
"""
Markprofil langs bedomda ledningar - underlag for "ledningsprofil mot markprofil".

Verktyget "Markprofil" i tv3_verktyg.pyt (eller markprofil() harifran) gar langs varje
objekt i ledningslagret fran "Skapa ledningslager", tar ut markhojden med jamna
mellanrum ur ett punktlager med markhojder och hamtar vattengangsnivaerna i bada
andarna ur det ursprungliga ledningslagret (GIS-nivaer i RH2000). Resultatet skrivs
som markprofil.json, som tv3_analys.py laser ('markprofil: FIL' i listfilen eller
--markprofil) for att:

  * hanga upp filmens inklinometerprofil pa GIS-nivaerna (filmerna har ofta ett lokalt
    nollplan - bara hojdskillnaden mellan brunnarna ar palitlig),
  * rita markprofilen i samma diagram som ledningsprofilen i PDF-protokollet,
  * berakna tackningen (mark minus hjassa) per stracka.

Markhojden i en punkt = avstandsviktat medel (IDW) av de narmaste matpunkterna inom
sokradien; ligger en matpunkt inom 5 cm anvands den rakt av. Utanfor sokradien
saknas varde (null).

Positionerna i JSON-filen ar meter langs kartlinjen fran startbrunnen (uppstroms);
tv3_analys.py skalar om dem till filmens positioner.
"""
from __future__ import unicode_literals

import os
import io
import json
import datetime
import arcpy

from skapa_ledningslager import (txt, logg, hitta_lager, hitta_falt, kalla, normalisera,
                                 _avst, _punkt_segment)


def _tal(v):
    try:
        f = float(v)
    except (TypeError, ValueError):
        return None
    if f != f or f < -9000:            # NaN eller SVOA:s -9999 = saknas
        return None
    return f


class Markpunkter(object):
    """Rutnatsindex over matpunkter (x, y, z) for snabb IDW-sokning."""

    def __init__(self, cell):
        self.cell = float(cell)
        self.rutnat = {}
        self.n = 0

    def lagg_till(self, x, y, z):
        self.rutnat.setdefault((int(x // self.cell), int(y // self.cell)), []).append((x, y, z))
        self.n += 1

    def hojd(self, x, y, radie, max_punkter=8):
        """(hojd, avstand till narmaste punkt) eller (None, None) om ingen punkt inom radien."""
        c = self.cell
        cx, cy = int(x // c), int(y // c)
        r2 = radie * radie
        nara = []
        for dx in (-1, 0, 1):
            for dy in (-1, 0, 1):
                for px, py, pz in self.rutnat.get((cx + dx, cy + dy), ()):
                    d2 = (px - x) ** 2 + (py - y) ** 2
                    if d2 <= r2:
                        nara.append((d2, pz))
        if not nara:
            return None, None
        nara.sort()
        nara = nara[:max_punkter]
        d_min = nara[0][0] ** 0.5
        if d_min < 0.05:
            return nara[0][1], d_min
        vikt = [(1.0 / d2, z) for d2, z in nara]
        s = sum(w for w, z in vikt)
        return sum(w * z for w, z in vikt) / s, d_min


def _punkter(geom):
    """Alla vertex i en polyline som [(x, y)], delar sammanslagna."""
    pts = []
    for del_ in geom:
        for p in del_:
            if p is not None:
                pts.append((p.X, p.Y))
    return pts


def _langs(pts, intervall):
    """Punkter med jamna mellanrum langs en bruten linje: [(m, x, y)], inkl. slutet."""
    if len(pts) < 2:
        return []
    ut = []
    m_tot = 0.0
    nasta = 0.0
    for i in range(len(pts) - 1):
        (ax, ay), (bx, by) = pts[i], pts[i + 1]
        seg = _avst(pts[i], pts[i + 1])
        while nasta <= m_tot + seg and seg > 0:
            t = (nasta - m_tot) / seg
            ut.append((nasta, ax + t * (bx - ax), ay + t * (by - ay)))
            nasta += intervall
        m_tot += seg
    if not ut or ut[-1][0] < m_tot - 1e-6:
        ut.append((m_tot, pts[-1][0], pts[-1][1]))
    return ut


def _langd(pts):
    return sum(_avst(pts[i], pts[i + 1]) for i in range(len(pts) - 1))


def _avst_till_linje(p, pts):
    """Minsta avstand fran punkt p till den brutna linjen pts."""
    bast = None
    for i in range(len(pts) - 1):
        d = _punkt_segment(p[0], p[1], pts[i][0], pts[i][1], pts[i + 1][0], pts[i + 1][1])[0]
        if bast is None or d < bast:
            bast = d
    return bast if bast is not None else 1e30


class Ledningsandar(object):
    """Index over ledningsobjektens andpunkter med vattengangsniva: for att hitta
    GIS-nivan i en strackans start- och slutbrunn."""

    def __init__(self, cell):
        self.cell = float(cell)
        self.rutnat = {}      # cell -> [(x, y, niva, granne, lager, oid)]
        self.n = 0

    def lagg_till(self, pts, vg_fran, vg_till, lager, oid):
        if len(pts) < 2:
            return
        for (x, y), niva, granne in ((pts[0], vg_fran, pts[1]), (pts[-1], vg_till, pts[-2])):
            self.rutnat.setdefault((int(x // self.cell), int(y // self.cell)), []).append(
                (x, y, niva, granne, lager, oid))
        self.n += 1

    def niva(self, x, y, tol, linje):
        """Vattengangsnivan i ledningsanden narmast (x, y) inom tol, dar ledningen dessutom
        foljer linjen (grannvertexet ligger inom tol fran den). (niva, avstand) eller (None, None)."""
        c = self.cell
        cx, cy = int(x // c), int(y // c)
        bast = None
        for dx in (-1, 0, 1):
            for dy in (-1, 0, 1):
                for ex, ey, niva, granne, lager, oid in self.rutnat.get((cx + dx, cy + dy), ()):
                    d = _avst((x, y), (ex, ey))
                    if d > tol or niva is None:
                        continue
                    if _avst_till_linje(granne, linje) > tol:
                        continue
                    if bast is None or d < bast[1]:
                        bast = (niva, d)
        return bast if bast else (None, None)


def markprofil(bedomda_lager, ledningslager, vg_fran_falt, vg_till_falt, marklager, json_ut,
               z_falt=None, intervall=1.0, sokradie=5.0, tolerans=2.0, hojdsystem='RH2000'):
    """Skriver markprofil.json. Returnerar sokvagen."""
    intervall, sokradie, tolerans = float(intervall), float(sokradie), float(tolerans)
    if not intervall > 0 or not sokradie > 0 or not tolerans > 0:
        raise RuntimeError('Intervall, sokradie och tolerans maste vara storre an 0')
    if isinstance(ledningslager, (type(''), bytes)):
        ledningslager = [ledningslager]

    logg('Letar upp lager')
    bed = hitta_lager(bedomda_lager)
    led = [hitta_lager(n) for n in ledningslager]
    mark = hitta_lager(marklager)

    # ---------------------------------------------------- 1. Bedomda strackor
    bed_src, bed_dq = kalla(bed)
    bed_falt = [f.name.upper() for f in arcpy.ListFields(bed_src)]
    for f in ('FRAN_BRUNN', 'TILL_BRUNN', 'TV3_FIL'):
        if f not in bed_falt:
            raise RuntimeError('Lagret saknar faltet %s - valj lagret fran "Skapa ledningslager"' % f)
    har_nr = 'NR' in bed_falt
    if not har_nr:
        logg('  OBS: lagret saknar faltet NR (aldre version) - strackorna kopplas pa brunnspar')
    arcpy.MakeFeatureLayer_management(bed_src, 'lyr_bed', bed_dq)
    strackor = []
    falt = ['SHAPE@', 'FRAN_BRUNN', 'TILL_BRUNN', 'TV3_FIL'] + (['NR'] if har_nr else [])
    with arcpy.da.SearchCursor('lyr_bed', falt) as mark_:
        for rad in mark_:
            geom = rad[0]
            if geom is None:
                continue
            pts = _punkter(geom)
            if len(pts) < 2:
                continue
            strackor.append({'pts': pts, 'fran': txt(rad[1]), 'till': txt(rad[2]),
                             'fil': txt(rad[3]), 'nr': rad[4] if har_nr else None})
    logg('  %d strackor i %s' % (len(strackor), txt(getattr(bed, 'name', bed))))
    if not strackor:
        raise RuntimeError('Inga strackor i lagret')

    # ---------------------------------------------------- 2. Vattengang ur ledningslagren
    andar = Ledningsandar(max(tolerans, 5.0))
    for lyr in led:
        src, dq = kalla(lyr)
        f_fran = hitta_falt(lyr, vg_fran_falt)
        f_till = hitta_falt(lyr, vg_till_falt)
        arcpy.MakeFeatureLayer_management(src, 'lyr_led', dq)
        arcpy.SelectLayerByLocation_management('lyr_led', 'INTERSECT', 'lyr_bed',
                                               '%s Meters' % tolerans, 'NEW_SELECTION')
        n = 0
        with arcpy.da.SearchCursor('lyr_led', ['OID@', 'SHAPE@', f_fran, f_till]) as mark_:
            for oid, geom, vf, vt in mark_:
                if geom is None:
                    continue
                andar.lagg_till(_punkter(geom), _tal(vf), _tal(vt),
                                txt(getattr(lyr, 'name', lyr)), oid)
                n += 1
        arcpy.Delete_management('lyr_led')
        logg('  %s: %d ledningar nara strackorna (vattengang: %s / %s)'
             % (txt(getattr(lyr, 'name', lyr)), n, f_fran, f_till))

    # ---------------------------------------------------- 3. Markhojder
    msrc, mdq = kalla(mark)
    arcpy.MakeFeatureLayer_management(msrc, 'lyr_mark', mdq)
    arcpy.SelectLayerByLocation_management('lyr_mark', 'INTERSECT', 'lyr_bed',
                                           '%s Meters' % sokradie, 'NEW_SELECTION')
    punkter = Markpunkter(sokradie)
    if z_falt:
        zf = hitta_falt(mark, z_falt)
        falt = ['SHAPE@XY', zf]
        logg('  markhojd ur faltet %s' % zf)
    else:
        falt = ['SHAPE@XY', 'SHAPE@Z']
        logg('  markhojd ur punkternas Z')
    n_utan = 0
    with arcpy.da.SearchCursor('lyr_mark', falt) as mark_:
        for xy, z in mark_:
            z = _tal(z)
            if not xy or xy[0] is None or z is None:
                n_utan += 1
                continue
            punkter.lagg_till(xy[0], xy[1], z)
    arcpy.Delete_management('lyr_mark')
    arcpy.Delete_management('lyr_bed')
    logg('  %d markhojdspunkter inom %.0f m fran strackorna%s'
         % (punkter.n, sokradie, (', %d utan hojd' % n_utan) if n_utan else ''))
    if not punkter.n:
        raise RuntimeError('Inga markhojdspunkter nara strackorna - fel lager, fel Z-falt eller '
                           'annat koordinatsystem?')

    # ---------------------------------------------------- 4. Ga langs varje stracka
    poster = []
    n_vg = n_mark = 0
    for s in strackor:
        pts = s['pts']
        prov = []
        n_traff = 0
        for m, x, y in _langs(pts, intervall):
            z, d = punkter.hojd(x, y, sokradie)
            prov.append([round(m, 2), round(z, 3) if z is not None else None,
                         round(d, 2) if d is not None else None])
            if z is not None:
                n_traff += 1
        vg_s, d_s = andar.niva(pts[0][0], pts[0][1], tolerans, pts)
        vg_e, d_e = andar.niva(pts[-1][0], pts[-1][1], tolerans, pts)
        if vg_s is not None and vg_e is not None:
            n_vg += 1
        if n_traff:
            n_mark += 1
        poster.append({
            'fil': s['fil'], 'nr': s['nr'],
            'startbrunn': s['fran'], 'slutbrunn': s['till'],
            'langd_karta_m': round(_langd(pts), 2),
            'vg_start': round(vg_s, 3) if vg_s is not None else None,
            'vg_slut': round(vg_e, 3) if vg_e is not None else None,
            'mark': prov,
            'mark_traffar': n_traff,
        })

    logg('  %d av %d strackor har vattengang i bada andarna, %d har markhojder'
         % (n_vg, len(poster), n_mark))
    utan_vg = [p for p in poster if p['vg_start'] is None or p['vg_slut'] is None]
    if utan_vg:
        logg('  utan vattengang (forsta 10): %s'
             % ', '.join('%s-%s' % (p['startbrunn'], p['slutbrunn']) for p in utan_vg[:10]))

    data = {
        'version': 1,
        'kalla': 'markprofil.py',
        'genererad': datetime.datetime.now().strftime('%Y-%m-%d %H:%M'),
        'hojdsystem': txt(hojdsystem),
        'intervall_m': intervall,
        'sokradie_m': sokradie,
        'marklager': txt(getattr(mark, 'name', mark)),
        'strackor': poster,
    }
    with io.open(json_ut, 'w', encoding='utf-8') as f:
        f.write(txt(json.dumps(data, ensure_ascii=False, indent=1)))
    logg('  %d strackor skrivna till %s' % (len(poster), json_ut))
    logg('KLART')
    return json_ut
