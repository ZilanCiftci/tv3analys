# -*- coding: utf-8 -*-
"""
Projektering av nya VA-strak i ArcMap: ritlager och profilritning.

Tva verktyg i tv3_verktyg.pyt (eller funktionerna harifran):

  skapa_projekteringslager(gdb, ...)
      Lagger upp tva tomma featureklasser i en filgeodatabas, med fardiga falt och
      vardelistor, sa att straket kan ritas med ArcMaps vanliga redigering:
        <prefix>_Ledning  (linje)  LEDN_ID, TYP (S/D/V/K), DIM (mm), MATERIAL,
                                   VG_UPP, VG_NED (vattengang i ledningens start- resp.
                                   slutpunkt, i ritad riktning = flodesriktning), KOMMENTAR
        <prefix>_Brunn    (punkt)  BRUNN_ID, TYP, LOCKNIVA, BOTTENNIVA, DIAM (mm), KOMMENTAR
      Flera ledningar i samma schakt (spill, dag, vatten) ritas som separata linjer.

  profil(ledningslager, brunnslager, marklager, ut_mapp, ...)
      Ritar profilen langs straket: markyta, vattengang och hjassa per ledning, brunnar
      fran botten till lock, fall i promille, sektioner. Alla ledningstyper i samma
      diagram. Ordningen langs straket raknas fram ur hur ledningarna hanger ihop
      (andpunkt mot brunn eller mot annan ledning inom toleransen), med den langsta
      ledningskedjan som referensaxel; ovriga typer projiceras pa den. Markhojden tas
      ur ett punktlager (Z eller falt) eller ett raster. Utdata: PDF + PNG + CSV per strak.
      Inga minimikrav kontrolleras - bakfall och ledning ovan mark noteras bara.

Ritningen gors med matplotlib i ArcMaps Python (finns i ArcGIS 10.1+). Saknas
matplotlib skrivs bara CSV och en JSON med profilen.
"""
from __future__ import unicode_literals, division

import os
import io
import csv
import json
import math
import datetime
import arcpy

from skapa_ledningslager import txt, logg, hitta_lager, hitta_falt, kalla, normalisera, _avst, _punkt_segment

LEDNINGSTYPER = [('S', 'Spillvatten'), ('D', 'Dagvatten'), ('V', 'Vatten'),
                 ('K', 'Kombinerat'), ('T', 'Tryckspill'), ('O', '\u00d6vrigt')]
TYPFARG = {'S': '#c0392b', 'D': '#1e8449', 'V': '#1f5fbf', 'K': '#7d3c98', 'T': '#e67e22', 'O': '#555555'}
MARKFARG = '#8c6d46'

LEDNING_FALT = [
    ('LEDN_ID',   'TEXT',   20,  'Ledningsnr'),
    ('TYP',       'TEXT',   2,   'Typ (S/D/V/K)'),
    ('DIM',       'SHORT',  None, 'Dimension (mm)'),
    ('MATERIAL',  'TEXT',   20,  'Material'),
    ('VG_UPP',    'DOUBLE', None, 'Vattengang uppstroms (start)'),
    ('VG_NED',    'DOUBLE', None, 'Vattengang nedstroms (slut)'),
    ('KOMMENTAR', 'TEXT',   100, 'Kommentar'),
]
BRUNN_FALT = [
    ('BRUNN_ID',   'TEXT',   30,  'Brunnsnr'),
    ('TYP',        'TEXT',   10,  'Typ'),
    ('LOCKNIVA',   'DOUBLE', None, 'Lockniva'),
    ('BOTTENNIVA', 'DOUBLE', None, 'Bottenniva'),
    ('DIAM',       'SHORT',  None, 'Diameter (mm)'),
    ('KOMMENTAR',  'TEXT',   100, 'Kommentar'),
]

# Fasta skalor for profilen (samma serier som protokollen i tv3_analys)
LANGDSKALOR = [100, 200, 250, 400, 500, 750, 1000, 1500, 2000, 2500, 4000, 5000]
HOJDSKALOR = [10, 20, 25, 50, 100, 200, 250, 500]
# Ritbladets storlek: A3 liggande, axelns andel av bladet
BLAD_MM = (420.0, 297.0)
AXEL = (0.07, 0.30, 0.90, 0.60)     # vanster, botten, bredd, hojd (andel)


def _tal(v):
    try:
        f = float(v)
    except (TypeError, ValueError):
        return None
    if f != f or f < -9000:
        return None
    return f


# =====================================================================
# 1. Ritlager
# =====================================================================

def skapa_projekteringslager(gdb, prefix='Proj', sr=None, sr_lager=None, lagg_till_i_kartan=True):
    """Skapar <prefix>_Ledning och <prefix>_Brunn i filgeodatabasen gdb (skapas om den
    inte finns). Koordinatsystem: sr, annars fran sr_lager, annars fran kartans dataram.
    Returnerar (ledning_fc, brunn_fc)."""
    gdb = txt(gdb)
    if not gdb.lower().endswith('.gdb'):
        gdb = gdb + '.gdb'
    if not arcpy.Exists(gdb):
        arcpy.CreateFileGDB_management(os.path.dirname(gdb), os.path.basename(gdb))
        logg('Skapade %s' % gdb)

    if sr is None and sr_lager:
        sr = arcpy.Describe(kalla(hitta_lager(sr_lager))[0]).spatialReference
    if sr is None:
        try:
            mxd = arcpy.mapping.MapDocument('CURRENT')
            sr = arcpy.mapping.ListDataFrames(mxd)[0].spatialReference
        except Exception:
            sr = None
    if sr is None:
        raise RuntimeError('Ange ett lager att ta koordinatsystemet fran')

    # Vardelistor
    domaner = [d.name for d in arcpy.da.ListDomains(gdb)]
    if 'PROJ_LEDNTYP' not in domaner:
        arcpy.CreateDomain_management(gdb, 'PROJ_LEDNTYP', 'Ledningstyp', 'TEXT', 'CODED')
        for kod, namn in LEDNINGSTYPER:
            arcpy.AddCodedValueToDomain_management(gdb, 'PROJ_LEDNTYP', kod, namn)

    ut = []
    for namn, geom, falt in ((prefix + '_Ledning', 'POLYLINE', LEDNING_FALT),
                             (prefix + '_Brunn', 'POINT', BRUNN_FALT)):
        fc = os.path.join(gdb, namn)
        if arcpy.Exists(fc):
            logg('  %s finns redan - rors inte' % namn)
        else:
            arcpy.CreateFeatureclass_management(gdb, namn, geom, '', 'DISABLED', 'DISABLED', sr)
            for f, typ, langd, alias in falt:
                if typ == 'TEXT':
                    arcpy.AddField_management(fc, f, typ, field_length=langd, field_alias=alias)
                else:
                    arcpy.AddField_management(fc, f, typ, field_alias=alias)
            if geom == 'POLYLINE':
                try:
                    arcpy.AssignDomainToField_management(fc, 'TYP', 'PROJ_LEDNTYP')
                except Exception as e:
                    logg('  kunde inte koppla vardelistan: %s' % txt(e))
            logg('  skapade %s' % namn)
        ut.append(fc)

    if lagg_till_i_kartan:
        try:
            mxd = arcpy.mapping.MapDocument('CURRENT')
            df = arcpy.mapping.ListDataFrames(mxd)[0]
            for fc in ut:
                lyr = arcpy.mapping.Layer(fc)
                arcpy.mapping.AddLayer(df, lyr, 'TOP')
            arcpy.RefreshTOC()
            arcpy.RefreshActiveView()
        except Exception as e:
            logg('  kunde inte lagga till lagren i kartan: %s' % txt(e))
    logg('Rita ledningarna i flodesriktningen (uppstroms -> nedstroms) och fyll i VG_UPP/VG_NED,'
         ' brunnarna med LOCKNIVA/BOTTENNIVA. Kor sedan "Projekteringsprofil".')
    return tuple(ut)


# =====================================================================
# 2. Markhojd
# =====================================================================

class Markhojd(object):
    """Markhojd ur ett punktlager (IDW av narmaste punkter inom sokradien) eller ett raster."""

    def __init__(self, lager, z_falt=None, sokradie=5.0, omrade_lager=None):
        self.sokradie = float(sokradie)
        self.raster = None
        self.rutnat = {}
        self.n = 0
        src = None
        try:
            src, dq = kalla(hitta_lager(lager))
        except Exception:
            src = txt(lager)
        d = arcpy.Describe(src)
        typ = txt(getattr(d, 'dataType', ''))
        if typ in ('RasterDataset', 'RasterLayer', 'MosaicDataset', 'MosaicLayer') or \
                txt(getattr(d, 'datasetType', '')) in ('RasterDataset', 'MosaicDataset'):
            self.raster = src
            self.cache = {}
            logg('  markhojd ur raster %s' % txt(getattr(d, 'name', src)))
            return
        arcpy.MakeFeatureLayer_management(src, 'lyr_mark_p', dq)
        if omrade_lager is not None:
            arcpy.SelectLayerByLocation_management('lyr_mark_p', 'INTERSECT', omrade_lager,
                                                   '%s Meters' % self.sokradie, 'NEW_SELECTION')
        if z_falt:
            zf = hitta_falt(hitta_lager(lager), z_falt)
            falt = ['SHAPE@XY', zf]
            logg('  markhojd ur faltet %s' % zf)
        else:
            falt = ['SHAPE@XY', 'SHAPE@Z']
            logg('  markhojd ur punkternas Z')
        with arcpy.da.SearchCursor('lyr_mark_p', falt) as mark:
            for xy, z in mark:
                z = _tal(z)
                if xy and xy[0] is not None and z is not None:
                    c = self.sokradie
                    self.rutnat.setdefault((int(xy[0] // c), int(xy[1] // c)), []).append((xy[0], xy[1], z))
                    self.n += 1
        arcpy.Delete_management('lyr_mark_p')
        logg('  %d markhojdspunkter nara straket' % self.n)

    def hojd(self, x, y):
        if self.raster is not None:
            nyckel = (round(x, 1), round(y, 1))
            if nyckel not in self.cache:
                try:
                    v = arcpy.GetCellValue_management(self.raster, '%f %f' % (x, y)).getOutput(0)
                    self.cache[nyckel] = _tal(txt(v).replace(',', '.'))
                except Exception:
                    self.cache[nyckel] = None
            return self.cache[nyckel]
        c = self.sokradie
        cx, cy = int(x // c), int(y // c)
        nara = []
        for dx in (-1, 0, 1):
            for dy in (-1, 0, 1):
                for px, py, pz in self.rutnat.get((cx + dx, cy + dy), ()):
                    d2 = (px - x) ** 2 + (py - y) ** 2
                    if d2 <= c * c:
                        nara.append((d2, pz))
        if not nara:
            return None
        nara.sort()
        nara = nara[:8]
        if nara[0][0] < 0.0025:
            return nara[0][1]
        s = sum(1.0 / d2 for d2, z in nara)
        return sum(z / d2 for d2, z in nara) / s


# =====================================================================
# 3. Geometri langs straket
# =====================================================================

def _punkter(geom):
    pts = []
    for del_ in geom:
        for p in del_:
            if p is not None:
                pts.append((p.X, p.Y))
    return pts


def _langd(pts):
    return sum(_avst(pts[i], pts[i + 1]) for i in range(len(pts) - 1))


def _station(pts, p):
    """(station langs pts for narmaste punkt, avstand dit)."""
    bast, st, m = None, 0.0, 0.0
    for i in range(len(pts) - 1):
        d, t, q = _punkt_segment(p[0], p[1], pts[i][0], pts[i][1], pts[i + 1][0], pts[i + 1][1])
        seg = _avst(pts[i], pts[i + 1])
        if bast is None or d < bast:
            bast, st = d, m + t * seg
        m += seg
    return st, (bast if bast is not None else 1e30)


def _langs(pts, intervall):
    ut, m_tot, nasta = [], 0.0, 0.0
    for i in range(len(pts) - 1):
        (ax, ay), (bx, by) = pts[i], pts[i + 1]
        seg = _avst(pts[i], pts[i + 1])
        while seg > 0 and nasta <= m_tot + seg + 1e-9:
            t = (nasta - m_tot) / seg
            ut.append((nasta, ax + t * (bx - ax), ay + t * (by - ay)))
            nasta += intervall
        m_tot += seg
    if not ut or ut[-1][0] < m_tot - 1e-6:
        ut.append((m_tot, pts[-1][0], pts[-1][1]))
    return ut


class Noder(object):
    """Andpunkter -> noder: brunn inom toleransen, annars fri ande (andar inom toleransen delas)."""

    def __init__(self, brunnar, tol):
        self.tol = float(tol)
        self.brunnar = brunnar          # id -> (x, y, post)
        self.fria = []                  # [(x, y)]

    def nod(self, p):
        bast = None
        for bid, (x, y, post) in self.brunnar.items():
            d = _avst(p, (x, y))
            if d <= self.tol and (bast is None or d < bast[1]):
                bast = (('B', bid), d)
        if bast:
            return bast[0]
        for i, (x, y) in enumerate(self.fria):
            if _avst(p, (x, y)) <= self.tol:
                return ('P', i)
        self.fria.append(p)
        return ('P', len(self.fria) - 1)


def _kedjor(ledningar, start_nod=None):
    """Ordnar ledningar (med 'fran'/'till'-noder i ritad riktning) i kedjor: fran en startnod
    foljs utgaende ledningar; forgreningar ger flera kedjor. Returnerar [[ledning, ...], ...]."""
    kvar = list(ledningar)
    ut = []
    while kvar:
        utg = {}
        ink = {}
        for l in kvar:
            utg.setdefault(l['fran'], []).append(l)
            ink.setdefault(l['till'], []).append(l)
        # startnod: angiven, annars nod utan inkommande (hogsta VG_UPP forst), annars forsta
        kandidater = [n for n in utg if n not in ink]
        if start_nod is not None and start_nod in utg:
            start = start_nod
        elif kandidater:
            start = max(kandidater, key=lambda n: max(_tal(l['vg_upp']) or -1e9 for l in utg[n]))
        else:
            start = kvar[0]['fran']
        kedja = []
        nod = start
        sedda = set()
        while nod in utg and nod not in sedda:
            sedda.add(nod)
            # vid forgrening: langsta grenen (flest ledningar nedstroms) foljs, ovriga blir egna kedjor
            val = max(utg[nod], key=lambda l: _djup(l, utg, set()))
            kedja.append(val)
            nod = val['till']
        for l in kedja:
            kvar.remove(l)
        ut.append(kedja)
        start_nod = None
    return ut


def _djup(l, utg, sedda):
    if l['till'] in sedda:
        return 1
    sedda = sedda | set([l['till']])
    return 1 + max([_djup(n, utg, sedda) for n in utg.get(l['till'], [])] or [0])


# =====================================================================
# 4. Profil
# =====================================================================

def profil(ledningslager, brunnslager, marklager, ut_mapp, namn='profil', z_falt=None,
           startbrunn=None, intervall=1.0, sokradie=5.0, tolerans=1.0, hojdsystem='RH2000',
           bara_valda=True):
    """Ritar profil(er) for straket i ledningslagret. Returnerar lista med skrivna PDF/PNG."""
    intervall, sokradie, tolerans = float(intervall), float(sokradie), float(tolerans)
    if not os.path.isdir(ut_mapp):
        os.makedirs(ut_mapp)
    logg('Letar upp lager')
    led = hitta_lager(ledningslager)
    brl = hitta_lager(brunnslager) if brunnslager else None

    # ---------------------------------------------------- ledningar (valda om nagra ar valda)
    def cursor_kalla(lyr):
        # Lagerobjekt i kartan respekterar urvalet; sokvag gor det inte
        if bara_valda and not isinstance(lyr, (type(''), bytes)):
            try:
                if lyr.getSelectionSet():
                    return lyr
            except Exception:
                pass
        return kalla(lyr)[0]

    falt = [f.name.upper() for f in arcpy.ListFields(kalla(led)[0])]
    for f in ('VG_UPP', 'VG_NED'):
        if f not in falt:
            raise RuntimeError('Ledningslagret saknar faltet %s - skapa lagren med "Skapa projekteringslager"' % f)
    lasfalt = ['OID@', 'SHAPE@'] + [f for f in ('LEDN_ID', 'TYP', 'DIM', 'MATERIAL', 'VG_UPP', 'VG_NED') if f in falt]
    ledningar = []
    with arcpy.da.SearchCursor(cursor_kalla(led), lasfalt) as mark:
        for rad in mark:
            post = dict(zip(lasfalt, rad))
            geom = post['SHAPE@']
            if geom is None:
                continue
            pts = _punkter(geom)
            if len(pts) < 2:
                continue
            ledningar.append({
                'oid': post['OID@'], 'pts': pts, 'langd': _langd(pts),
                'id': txt(post.get('LEDN_ID') or '') or ('#%s' % post['OID@']),
                'typ': (txt(post.get('TYP') or 'O').strip().upper()[:1] or 'O'),
                'dim': _tal(post.get('DIM')) or 0.0,
                'material': txt(post.get('MATERIAL') or ''),
                'vg_upp': _tal(post.get('VG_UPP')), 'vg_ned': _tal(post.get('VG_NED')),
            })
    logg('  %d ledningar' % len(ledningar))
    if not ledningar:
        raise RuntimeError('Inga ledningar (ar nagra valda? avmarkera eller valj straket)')

    # ---------------------------------------------------- brunnar
    brunnar = {}
    if brl is not None:
        bf = [f.name.upper() for f in arcpy.ListFields(kalla(brl)[0])]
        bfalt = ['OID@', 'SHAPE@XY'] + [f for f in ('BRUNN_ID', 'TYP', 'LOCKNIVA', 'BOTTENNIVA', 'DIAM') if f in bf]
        with arcpy.da.SearchCursor(kalla(brl)[0], bfalt) as mark:
            for rad in mark:
                post = dict(zip(bfalt, rad))
                xy = post['SHAPE@XY']
                if not xy or xy[0] is None:
                    continue
                bid = txt(post.get('BRUNN_ID') or '') or ('B%s' % post['OID@'])
                brunnar[bid] = (xy[0], xy[1], {
                    'id': bid, 'typ': txt(post.get('TYP') or ''),
                    'lock': _tal(post.get('LOCKNIVA')), 'botten': _tal(post.get('BOTTENNIVA')),
                    'diam': _tal(post.get('DIAM')) or 0.0})
        logg('  %d brunnar' % len(brunnar))

    noder = Noder(brunnar, tolerans)
    for l in ledningar:
        l['fran'] = noder.nod(l['pts'][0])
        l['till'] = noder.nod(l['pts'][-1])

    start_nod = None
    if startbrunn:
        s = normalisera(startbrunn)
        for bid in brunnar:
            if normalisera(bid) == s:
                start_nod = ('B', bid)
        if start_nod is None:
            logg('  VARNING: startbrunnen %s finns inte i brunnslagret' % txt(startbrunn))

    # ---------------------------------------------------- kedjor per typ, referensaxel
    per_typ = {}
    for l in ledningar:
        per_typ.setdefault(l['typ'], []).append(l)
    kedjor = []          # (typ, [ledningar i ordning])
    for typ, lista in sorted(per_typ.items()):
        for k in _kedjor(lista, start_nod):
            kedjor.append((typ, k))
    kedjor.sort(key=lambda tk: -sum(l['langd'] for l in tk[1]))
    if not kedjor:
        raise RuntimeError('Kunde inte ordna ledningarna')

    # Referensaxel = langsta kedjan; ovriga kedjor projiceras pa den om de foljer den
    # (bada andarna inom 3 x tolerans + 2 m), annars blir de egna strak.
    strak = []
    for typ, k in kedjor:
        axel = []
        for l in k:
            axel += l['pts'] if not axel else l['pts'][1:] if _avst(axel[-1], l['pts'][0]) < 1e-6 else l['pts']
        placerad = False
        for s in strak:
            ok = True
            for l in k:
                for p in (l['pts'][0], l['pts'][-1]):
                    if _station(s['axel'], p)[1] > 3 * tolerans + 2.0:
                        ok = False
            if ok:
                s['kedjor'].append((typ, k))
                placerad = True
                break
        if not placerad:
            strak.append({'axel': axel, 'kedjor': [(typ, k)], 'namn': namn if not strak else '%s_%d' % (namn, len(strak) + 1)})
    logg('  %d strak' % len(strak))

    # ---------------------------------------------------- markhojd
    mark = Markhojd(marklager, z_falt, sokradie, omrade_lager=led if bara_valda else None) if marklager else None

    ut_filer = []
    for s in strak:
        axel = s['axel']
        L = _langd(axel)
        # mark
        s['mark'] = []
        if mark is not None:
            for m, x, y in _langs(axel, intervall):
                z = mark.hojd(x, y)
                if z is not None:
                    s['mark'].append((m, z))
        # ledningar med stationer
        s['ledningar'] = []
        for typ, k in s['kedjor']:
            for l in k:
                st0 = _station(axel, l['pts'][0])[0]
                st1 = _station(axel, l['pts'][-1])[0]
                if abs(st1 - st0) < 0.01:
                    st1 = st0 + l['langd']
                post = dict(l)
                post.update({'st0': st0, 'st1': st1})
                fall = (post['vg_upp'] - post['vg_ned']) if post['vg_upp'] is not None and post['vg_ned'] is not None else None
                post['fall'] = fall
                post['lutning'] = (fall / l['langd'] * 1000) if fall is not None and l['langd'] > 0 else None
                post['tackning_min'], post['tackning_max'] = _tackning(post, s['mark'])
                s['ledningar'].append(post)
        # brunnar pa straket
        s['brunnar'] = []
        sedda = set()
        for post in s['ledningar']:
            for nod, st in ((post['fran'], post['st0']), (post['till'], post['st1'])):
                if nod[0] == 'B' and nod[1] not in sedda:
                    sedda.add(nod[1])
                    b = dict(brunnar[nod[1]][2])
                    b['st'] = st
                    # vattengangar i brunnen (for botten om BOTTENNIVA saknas)
                    vg = [p['vg_upp'] for p in s['ledningar'] if p['fran'] == nod and p['vg_upp'] is not None] + \
                         [p['vg_ned'] for p in s['ledningar'] if p['till'] == nod and p['vg_ned'] is not None]
                    b['vg'] = vg
                    s['brunnar'].append(b)

        # ---- CSV
        csv_fil = os.path.join(ut_mapp, s['namn'] + '.csv')
        with io.open(csv_fil, 'w', encoding='cp1252', errors='replace', newline='') as f:
            f.write('typ;ledning;fran;till;sektion_fran_m;sektion_till_m;langd_m;dim_mm;material;'
                    'vg_upp;vg_ned;fall_m;lutning_promille;tackning_min_m;tackning_max_m;anmarkning\n')
            for p in sorted(s['ledningar'], key=lambda p: (p['typ'], p['st0'])):
                anm = []
                if p['lutning'] is not None and p['lutning'] < 0:
                    anm.append('bakfall')
                if p['tackning_min'] is not None and p['tackning_min'] < 0:
                    anm.append('ledning ovan mark')
                if p['vg_upp'] is None or p['vg_ned'] is None:
                    anm.append('vattengang saknas')
                f.write(';'.join(txt(v) if v is not None else '' for v in [
                    p['typ'], p['id'], _nodnamn(p['fran']), _nodnamn(p['till']),
                    '%.1f' % p['st0'], '%.1f' % p['st1'], '%.1f' % p['langd'], '%g' % p['dim'], p['material'],
                    _fmt(p['vg_upp']), _fmt(p['vg_ned']), _fmt(p['fall']),
                    ('%.1f' % p['lutning']) if p['lutning'] is not None else '',
                    _fmt(p['tackning_min']), _fmt(p['tackning_max']), ', '.join(anm)]) + '\n')
        ut_filer.append(csv_fil)
        logg('  %s: %.0f m, %d ledningar, %d brunnar, %d markpunkter -> %s'
             % (s['namn'], L, len(s['ledningar']), len(s['brunnar']), len(s['mark']), csv_fil))

        # ---- JSON (for tv3_analys eller annan ritning)
        json_fil = os.path.join(ut_mapp, s['namn'] + '.json')
        with io.open(json_fil, 'w', encoding='utf-8') as f:
            f.write(txt(json.dumps({
                'version': 1, 'kalla': 'projektering.py', 'hojdsystem': txt(hojdsystem),
                'genererad': datetime.datetime.now().strftime('%Y-%m-%d %H:%M'),
                'namn': s['namn'], 'langd_m': round(L, 2),
                'mark': [[round(m, 2), round(z, 3)] for m, z in s['mark']],
                'ledningar': [dict((k, v) for k, v in p.items() if k not in ('pts', 'fran', 'till'))
                              for p in s['ledningar']],
                'brunnar': s['brunnar'],
            }, ensure_ascii=False, indent=1, default=txt)))

        # ---- ritning
        try:
            pdf = rita_profil(s, os.path.join(ut_mapp, s['namn']), hojdsystem)
            ut_filer += pdf
            logg('  ritning: %s' % ', '.join(pdf))
        except ImportError:
            logg('  matplotlib saknas i den har Python-installationen - bara CSV och JSON skrivna')
        except Exception as e:
            logg('  kunde inte rita profilen: %s' % txt(e))

    logg('KLART')
    return ut_filer


def _nodnamn(nod):
    return nod[1] if nod[0] == 'B' else 'fri ande %d' % (nod[1] + 1)


def _fmt(v):
    return ('%.2f' % v) if v is not None else ''


def _tackning(p, mark):
    """(min, max) tackning = mark - hjassa langs ledningen, eller (None, None)."""
    if p['vg_upp'] is None or p['vg_ned'] is None or not mark:
        return None, None
    st0, st1 = p['st0'], p['st1']
    if st1 < st0:
        st0, st1 = st1, st0
    varden = []
    for m, z in mark:
        if st0 - 0.01 <= m <= st1 + 0.01:
            t = (m - p['st0']) / (p['st1'] - p['st0']) if p['st1'] != p['st0'] else 0.0
            vg = p['vg_upp'] + (p['vg_ned'] - p['vg_upp']) * t
            varden.append(z - (vg + p['dim'] / 1000.0))
    if not varden:
        return None, None
    return min(varden), max(varden)


# =====================================================================
# 5. Ritning (matplotlib)
# =====================================================================

def rita_profil(s, bas, hojdsystem='RH2000'):
    """Ritar straket till <bas>.pdf och <bas>.png i exakt skala pa A3 liggande."""
    import matplotlib
    matplotlib.use('Agg')
    import matplotlib.pyplot as plt
    from matplotlib.patches import Rectangle

    L = max(_langd(s['axel']), 1.0)
    z_alla = [z for _, z in s['mark']]
    for p in s['ledningar']:
        for v in (p['vg_upp'], p['vg_ned']):
            if v is not None:
                z_alla += [v, v + p['dim'] / 1000.0]
    for b in s['brunnar']:
        for v in (b['lock'], b['botten']):
            if v is not None:
                z_alla.append(v)
    if not z_alla:
        raise RuntimeError('inga hojder att rita')
    zmin, zmax = min(z_alla), max(z_alla)

    ax_w_mm = BLAD_MM[0] * AXEL[2]
    ax_h_mm = BLAD_MM[1] * AXEL[3]
    xspann = L * 1.04
    zspann = max(zmax - zmin, 0.5) * 1.35
    langdskala = next((k for k in LANGDSKALOR if xspann * 1000 / k <= ax_w_mm), LANGDSKALOR[-1])
    hojdskala = next((k for k in HOJDSKALOR if zspann * 1000 / k <= ax_h_mm), HOJDSKALOR[-1])
    xvidd = ax_w_mm * langdskala / 1000.0
    zvidd = ax_h_mm * hojdskala / 1000.0
    x0 = -(xvidd - L) / 2.0
    zmitt = (zmax + zmin) / 2.0

    fig = plt.figure(figsize=(BLAD_MM[0] / 25.4, BLAD_MM[1] / 25.4))
    ax = fig.add_axes(AXEL)
    ax.set_xlim(x0, x0 + xvidd)
    ax.set_ylim(zmitt - zvidd / 2.0, zmitt + zvidd / 2.0)
    ax.grid(True, color='#e1e0d9', linewidth=0.6)
    for sp in ('top', 'right'):
        ax.spines[sp].set_visible(False)

    # markyta
    if s['mark']:
        xs = [m for m, _ in s['mark']]
        zs = [z for _, z in s['mark']]
        ax.plot(xs, zs, color=MARKFARG, linewidth=1.6, label='markyta')
        ax.fill_between(xs, zs, [zmitt - zvidd] * len(xs), color=MARKFARG, alpha=0.05, linewidth=0)

    # ledningar
    ritade_typer = set()
    for p in sorted(s['ledningar'], key=lambda p: p['st0']):
        if p['vg_upp'] is None or p['vg_ned'] is None:
            continue
        farg = TYPFARG.get(p['typ'], TYPFARG['O'])
        d = p['dim'] / 1000.0
        xs = [p['st0'], p['st1']]
        vg = [p['vg_upp'], p['vg_ned']]
        hj = [v + d for v in vg]
        ax.fill_between(xs, vg, hj, color=farg, alpha=0.25, linewidth=0)
        ax.plot(xs, vg, color=farg, linewidth=1.8,
                label=None if p['typ'] in ritade_typer else _typnamn(p['typ']))
        ax.plot(xs, hj, color=farg, linewidth=0.8, linestyle='--')
        ritade_typer.add(p['typ'])
        etik = '%s %s %g %s' % (p['typ'], p['id'], p['dim'], p['material'])
        etik += '\n%.1f m' % p['langd']
        if p['lutning'] is not None:
            etik += '  %.1f %%o' % p['lutning']
        ax.annotate(etik.replace('%o', '\u2030'), ((xs[0] + xs[1]) / 2.0, (hj[0] + hj[1]) / 2.0),
                    xytext=(0, 6), textcoords='offset points', ha='center', va='bottom',
                    fontsize=7, color=farg)

    # brunnar (etiketten lyfts nar en annan brunn ligger inom 4 m - t.ex. spill- och dagbrunn)
    forra_st = None
    lyft = 0
    for b in sorted(s['brunnar'], key=lambda b: b['st']):
        lyft = lyft + 1 if forra_st is not None and abs(b['st'] - forra_st) < 4.0 else 0
        forra_st = b['st']
        lock = b['lock']
        botten = b['botten'] if b['botten'] is not None else (min(b['vg']) - 0.1 if b['vg'] else None)
        if lock is None and botten is None:
            continue
        top = lock if lock is not None else (max(b['vg']) + 1.0 if b['vg'] else botten + 1.0)
        bot = botten if botten is not None else top - 1.0
        w = max((b['diam'] or 0) / 1000.0, 0.6)
        ax.add_patch(Rectangle((b['st'] - w / 2.0, bot), w, top - bot, facecolor='#d9d9d9',
                               edgecolor='#4d4d4d', linewidth=1.0, zorder=3))
        t = b['id']
        if lock is not None:
            t += '\nlock %.2f' % lock
        if botten is not None:
            t += '\nbotten %.2f' % botten
        ax.annotate(t, (b['st'], top), xytext=(0, 5 + 30 * lyft), textcoords='offset points',
                    ha='center', va='bottom', fontsize=7, fontweight='bold', color='#0b0b0b',
                    arrowprops=dict(arrowstyle='-', color='#7f7f7f', lw=0.5) if lyft else None)

    ax.set_xlabel('sektion (m)', fontsize=9)
    ax.set_ylabel('h\u00f6jd (m, %s)' % txt(hojdsystem), fontsize=9)
    ax.tick_params(labelsize=8)
    ax.set_title('Profil %s   \u00b7   l\u00e4ngdskala 1:%d   \u00b7   h\u00f6jdskala 1:%d   \u00b7   %s' % (
        s['namn'], langdskala, hojdskala, datetime.datetime.now().strftime('%Y-%m-%d')),
        fontsize=11, loc='left')
    try:
        ax.legend(loc='upper right', fontsize=8, frameon=False)
    except Exception:
        pass

    # nivatabell under diagrammet: sektion och vattengang per ledningstyp
    typer = sorted(set(p['typ'] for p in s['ledningar']))
    rader = ['sektion'] + ['VG %s' % t for t in typer] + ['mark']
    punkter = sorted(set([p['st0'] for p in s['ledningar']] + [p['st1'] for p in s['ledningar']]))
    grupper = []                       # sektioner inom 1 m blir en kolumn
    for st in punkter:
        if grupper and st - grupper[-1][-1] <= 1.0:
            grupper[-1].append(st)
        else:
            grupper.append([st])
    y0 = AXEL[1] - 0.05
    for i, rubrik in enumerate(rader):
        fig.text(AXEL[0] - 0.005, y0 - i * 0.028, rubrik, ha='right', va='center', fontsize=7.5)
    for g in grupper:
        st = sum(g) / len(g)
        fx = AXEL[0] + (st - x0) / xvidd * AXEL[2]
        if not (AXEL[0] <= fx <= AXEL[0] + AXEL[2]):
            continue
        fig.text(fx, y0, '%.1f' % st, ha='center', va='center', fontsize=7)
        for i, t in enumerate(typer, 1):
            v = [p['vg_upp'] for p in s['ledningar'] if p['typ'] == t and p['vg_upp'] is not None
                 and any(abs(p['st0'] - x) < 0.05 for x in g)] + \
                [p['vg_ned'] for p in s['ledningar'] if p['typ'] == t and p['vg_ned'] is not None
                 and any(abs(p['st1'] - x) < 0.05 for x in g)]
            if v:
                fig.text(fx, y0 - i * 0.028, '/'.join('%.2f' % x for x in sorted(set(round(x, 2) for x in v))),
                         ha='center', va='center', fontsize=7, color=TYPFARG.get(t, '#000'))
        if s['mark']:
            zm = min(s['mark'], key=lambda mz: abs(mz[0] - st))[1]
            fig.text(fx, y0 - len(rader[1:]) * 0.028, '%.2f' % zm, ha='center', va='center', fontsize=7, color=MARKFARG)

    ut = []
    for andelse in ('.pdf', '.png'):
        fil = bas + andelse
        fig.savefig(fil, dpi=150)
        ut.append(fil)
    plt.close(fig)
    return ut


def _typnamn(typ):
    for kod, namn in LEDNINGSTYPER:
        if kod == typ:
            return namn
    return typ
