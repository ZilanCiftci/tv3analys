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


class Ledningar(object):
    """Ledningsobjekt med vattengang i bada andarna. Ger GIS-nivan i en punkt pa en ledning,
    linjart interpolerad mellan andarna - sa att aven brunnar mitt pa en ledning (strackor
    som Natverk klippt ur ett langre ledningsobjekt) far en niva."""

    def __init__(self):
        self.ledningar = []    # (pts, langd, vg_fran, vg_till, lager, oid, bbox)
        self.n = 0
        self.n_stigande = 0    # ledningar dar vg_fran < vg_till (ritad mot flodet?)
        self.n_bada = 0

    def lagg_till(self, pts, vg_fran, vg_till, lager, oid):
        if len(pts) < 2:
            return
        xs = [p[0] for p in pts]
        ys = [p[1] for p in pts]
        self.ledningar.append((pts, _langd(pts), vg_fran, vg_till, lager, oid,
                               (min(xs), min(ys), max(xs), max(ys))))
        self.n += 1
        if vg_fran is not None and vg_till is not None:
            self.n_bada += 1
            if vg_fran < vg_till - 0.005:
                self.n_stigande += 1

    def niva(self, x, y, tol, linje):
        """Vattengangsnivan vid (x, y): narmaste ledning inom tol som dessutom foljer
        strackans linje (en punkt 1 m bort langs ledningen ligger inom tol fran linjen).
        Nivan interpoleras efter laget langs ledningen. (niva, avstand) eller (None, None)."""
        bast = None
        for pts, L, vf, vt, lager, oid, (x0, y0, x1, y1) in self.ledningar:
            if x < x0 - tol or x > x1 + tol or y < y0 - tol or y > y1 + tol:
                continue
            if vf is None and vt is None:
                continue
            st, d = _station(pts, (x, y))
            if d > tol or (bast is not None and d >= bast[1]):
                continue
            # foljer ledningen strackans linje har? kolla en punkt 1 m at vardera hallet
            steg = min(1.0, L / 2.0)
            if steg > 0 and not any(_avst_till_linje(_punkt_vid(pts, st + r), linje) <= tol
                                    for r in (-steg, steg) if 0 <= st + r <= L):
                continue
            if vf is None or vt is None or L <= 0:
                niva = vf if vf is not None else vt
            else:
                niva = vf + (vt - vf) * st / L
            bast = (niva, d)
        return bast if bast else (None, None)


def _station(pts, p):
    """(matt langs linjen for narmaste punkt, avstand dit)."""
    bast, st, m = None, 0.0, 0.0
    for i in range(len(pts) - 1):
        d, tt, q = _punkt_segment(p[0], p[1], pts[i][0], pts[i][1], pts[i + 1][0], pts[i + 1][1])
        seg = _avst(pts[i], pts[i + 1])
        if bast is None or d < bast:
            bast, st = d, m + tt * seg
        m += seg
    return st, (bast if bast is not None else 1e30)


def _punkt_vid(pts, st):
    """Punkten pa linjen vid mattet st (klamms till linjens andar)."""
    m = 0.0
    for i in range(len(pts) - 1):
        seg = _avst(pts[i], pts[i + 1])
        if seg > 0 and st <= m + seg:
            tt = max(0.0, (st - m) / seg)
            return (pts[i][0] + tt * (pts[i + 1][0] - pts[i][0]), pts[i][1] + tt * (pts[i + 1][1] - pts[i][1]))
        m += seg
    return pts[-1]


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
    andar = Ledningar()
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
    # 3. Rimlighet: vattengangen ska normalt falla i ritad riktning. Stiger den pa de flesta
    # ledningar betyder faltet nagot annat (uppstroms/nedstroms efter flode?) eller ledningarna
    # ar ritade mot flodet.
    if andar.n_bada and andar.n_stigande > 0.5 * andar.n_bada:
        logg('  OBS: vattengangen stiger i ritad riktning pa %d av %d ledningar - kontrollera att'
             ' faltet "vid startpunkt" verkligen galler ledningens forsta vertex'
             % (andar.n_stigande, andar.n_bada))

    # ---------------------------------------------------- 3. Markhojder
    msrc, mdq = kalla(mark)
    md = arcpy.Describe(msrc)
    multipunkt = txt(getattr(md, 'shapeType', '')) == 'Multipoint'
    if not z_falt and not getattr(md, 'hasZ', True):
        raise RuntimeError('Markhojdslagret saknar Z i geometrin - ange faltet med hojden')
    arcpy.MakeFeatureLayer_management(msrc, 'lyr_mark', mdq)
    arcpy.SelectLayerByLocation_management('lyr_mark', 'INTERSECT', 'lyr_bed',
                                           '%s Meters' % sokradie, 'NEW_SELECTION')
    punkter = Markpunkter(sokradie)
    n_utan = 0
    if z_falt:
        zf = hitta_falt(mark, z_falt)
        logg('  markhojd ur faltet %s' % zf)
        with arcpy.da.SearchCursor('lyr_mark', ['SHAPE@XY', zf]) as mark_:
            for xy, z in mark_:
                z = _tal(z)
                if not xy or xy[0] is None or z is None:
                    n_utan += 1
                    continue
                punkter.lagg_till(xy[0], xy[1], z)
    elif multipunkt:
        logg('  markhojd ur Z i multipunktlagret (t.ex. laserdata)')
        with arcpy.da.SearchCursor('lyr_mark', ['SHAPE@']) as mark_:
            for (geom,) in mark_:
                if geom is None:
                    continue
                for del_ in geom:
                    for pt in (del_ if hasattr(del_, '__iter__') else [del_]):
                        if pt is None:
                            continue
                        z = _tal(getattr(pt, 'Z', None))
                        if z is None:
                            n_utan += 1
                            continue
                        punkter.lagg_till(pt.X, pt.Y, z)
    else:
        logg('  markhojd ur punkternas Z')
        with arcpy.da.SearchCursor('lyr_mark', ['SHAPE@XY', 'SHAPE@Z']) as mark_:
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
