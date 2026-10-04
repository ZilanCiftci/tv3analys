# -*- coding: utf-8 -*-
"""
Etapplager (ArcMap 10.x, Python 2.7 + arcpy): slar ihop strackorna i det bedomda ledningslagret
(fran Skapa ledningslager) per etapp till ett linjelager med en (flerdelad) linje per etapp, sa att
de sammanhangande etapperna och deras nummer syns i kartan. Uppgifterna per etapp (metod, langd,
kostnad, brunnar att schakta fram, flaggor) tas ur kartunderlag.json (toppfaltet "etapper") nar
filen anges, annars raknas de ur lagrets falt (ETAPP, METOD, KOSTNAD, BEDOMNING, INDEX_K, LANGD_M).

Valfritt skrivs ocksa ett punktlager med brunnarna som ska schaktas fram (ur JSON-filen; laget tas
ur strackornas andpunkter i det bedomda lagret).

Testas utan ArcMap med en latsas-arcpy (scratchpad/test_etapplager.py).
"""
from __future__ import unicode_literals

import io
import json
import os

import arcpy

from skapa_ledningslager import txt, logg, hitta_lager, kalla, normalisera, _skapa_fc, _utan_null, _mxd

LAGERNAMN = 'Etapper (atgardspaket)'
LAGERNAMN_BRUNNAR = 'Brunnar att schakta fram'
HAR = os.path.dirname(os.path.abspath(__file__))
LYR_ETAPPER = os.path.join(HAR, 'etapper.lyr')
LYR_FRAMSCHAKT = os.path.join(HAR, 'framschaktning.lyr')

ETAPP_FALT = [
    # namn,        typ,      langd, alias
    ('ETAPP',      'LONG',   None, 'Etapp'),
    ('METOD',      'TEXT',   10,  'Metod (strumpa/schakt)'),
    ('HOGSTA_KL',  'TEXT',   2,   'Hogsta gallande klass'),
    ('MAX_IDX',    'DOUBLE', None, 'Max konstruktionsindex (p/100 m)'),
    ('ANT_STR',    'LONG',   None, 'Antal strackor'),
    ('LANGD_M',    'DOUBLE', None, 'Langd (m)'),
    ('SCHAKT_M',   'DOUBLE', None, 'Schakt, langd (m)'),
    ('KOSTNAD',    'DOUBLE', None, 'Kostnad totalt (kr, strumpa)'),
    ('DIMENSION',  'TEXT',   40,  'Dimensioner (mm)'),
    ('BRUNNAR',    'TEXT',   254, 'Brunnar'),
    ('FRAMSCHAKT', 'TEXT',   254, 'Brunnar att schakta fram'),
    ('ANT_FRAM',   'LONG',   None, 'Antal framschaktade brunnar'),
    ('SERV_UPP',   'LONG',   None, 'Serviser uppstroms (max)'),
    ('FLAGGOR',    'TEXT',   254, 'Flaggor'),
    ('STRACKOR',   'TEXT',   254, 'Strackor (fil nr: brunnar)'),
    ('ETIKETT',    'TEXT',   80,  'Etikett'),
]
BRUNN_FALT = [
    ('BRUNN',      'TEXT',   50,  'Littera'),
    ('ETAPP',      'LONG',   None, 'Etapp'),
    ('MANUELL',    'TEXT',   3,   'Manuellt vald'),
    ('ETIKETT',    'TEXT',   60,  'Etikett'),
]


def _text(v):
    return txt(v) if v is not None else ''


def _tal(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


def _kr(v):
    """1 234 567 kr -> '1,2 Mkr' / '850 tkr'."""
    if not v:
        return ''
    if v >= 1e6:
        return ('%.1f Mkr' % (v / 1e6)).replace('.', ',')
    return '%d tkr' % round(v / 1e3)


def _delar(geom):
    """Punktlistor per del i en polyline-geometri."""
    ut = []
    for del_ in geom:
        pts = [p for p in del_ if p is not None]
        if len(pts) >= 2:
            ut.append(pts)
    return ut


def las_bedomda(bedomda):
    """Strackorna i det bedomda lagret: lista med dict (etapp, metod, kostnad, klass, index, langd,
    fran, till, nr, fil, delar [punktlistor]). Returnerar (poster, sr)."""
    lyr = hitta_lager(bedomda)
    src, dq = kalla(lyr)
    sr = arcpy.Describe(src).spatialReference
    befintliga = dict((f.name.upper(), f.name) for f in arcpy.ListFields(src))
    onskade = ['ETAPP', 'METOD', 'KOSTNAD', 'BEDOMNING', 'INDEX_K', 'LANGD_M', 'FRAN_BRUNN', 'TILL_BRUNN',
               'NR', 'TV3_FIL', 'MAN_BED', 'MASK_BED']
    falt = [befintliga[n] for n in onskade if n in befintliga]
    saknas = [n for n in onskade if n not in befintliga]
    if 'ETAPP' in saknas:
        raise RuntimeError('Lagret saknar faltet ETAPP - kor Skapa ledningslager med ett kartunderlag '
                           'fran en analys med etapper (standard) forst.')
    if saknas:
        logg('  VARNING: faltet/falten %s saknas i lagret' % ', '.join(saknas))
    arcpy.MakeFeatureLayer_management(src, 'lyr_etapp_in', dq)
    poster = []
    with arcpy.da.SearchCursor('lyr_etapp_in', ['SHAPE@'] + falt) as mark:
        for rad in mark:
            geom = rad[0]
            v = dict(zip(falt, rad[1:]))
            if geom is None:
                continue
            klass = _text(v.get('BEDOMNING')) or _text(v.get('MAN_BED')) or _text(v.get('MASK_BED'))
            poster.append({
                'etapp': v.get('ETAPP'), 'metod': _text(v.get('METOD')), 'kostnad': _tal(v.get('KOSTNAD')),
                'klass': klass, 'index': _tal(v.get('INDEX_K')),
                'langd': _tal(v.get('LANGD_M')) or getattr(geom, 'length', 0.0), 'fran': _text(v.get('FRAN_BRUNN')),
                'till': _text(v.get('TILL_BRUNN')), 'nr': v.get('NR'), 'fil': _text(v.get('TV3_FIL')),
                'delar': _delar(geom), 'geom': geom,
            })
    arcpy.Delete_management('lyr_etapp_in')
    return poster, sr


def _etapp_ur_json(json_fil):
    """{etappnr: post} ur kartunderlag.json, eller {} nar filen saknar etapper."""
    if not json_fil:
        return {}
    with io.open(json_fil, 'r', encoding='utf-8') as f:
        data = json.load(f)
    ut = {}
    for e in data.get('etapper') or []:
        try:
            ut[int(e.get('nr'))] = e
        except (TypeError, ValueError):
            pass
    if not ut:
        logg('  VARNING: %s saknar etapplistan ("etapper") - kor analysen igen med senaste tv3_analys.py; '
             'uppgifterna raknas ur lagret i stallet' % os.path.basename(json_fil))
    return ut


def skapa_etapplager(bedomda, ut_fc, json_fil=None, brunnar_ut=None, lyr_fil=None, lyr_brunnar=None,
                     lagg_till_i_kartan=True):
    """Skriver etapplagret (och valfritt punktlagret med framschaktade brunnar).
    Returnerar {'etapper': (fil, antal), 'brunnar': (fil, antal) | None}."""
    poster, sr = las_bedomda(bedomda)
    ur_json = _etapp_ur_json(json_fil)
    grupper = {}
    for p in poster:
        try:
            nr = int(p['etapp'])
        except (TypeError, ValueError):
            continue
        if nr <= 0:
            continue
        grupper.setdefault(nr, []).append(p)
    if not grupper:
        raise RuntimeError('Inga strackor med etappnummer i lagret (faltet ETAPP ar tomt).')
    logg('  %d strackor i %d etapper' % (sum(len(g) for g in grupper.values()), len(grupper)))

    ut_fil, ar_shapefil, max_text = _skapa_fc(ut_fc, 'POLYLINE', ETAPP_FALT, sr)
    namn = [f[0] for f in ETAPP_FALT]
    typer = [f[1] for f in ETAPP_FALT]
    insert = arcpy.da.InsertCursor(ut_fil, ['SHAPE@'] + namn)
    n = 0
    brunnspunkter = []          # (littera, etapp, manuell, x, y)
    for nr in sorted(grupper):
        st = grupper[nr]
        e = ur_json.get(nr, {})
        delar = arcpy.Array()
        for p in st:
            for pts in p['delar']:
                delar.add(arcpy.Array(pts))
        geom = arcpy.Polyline(delar, sr)
        metod = e.get('metod') or (st[0]['metod'] if st else '')
        klasser = [p['klass'] for p in st if p['klass']]
        hogsta = e.get('hogsta_klass') or (min(klasser) if klasser else '')
        idx = e.get('max_konstruktionsindex')
        if idx is None:
            idx = max([p['index'] for p in st if p['index'] is not None] or [None])
        langd = e.get('langd_m') if e.get('langd_m') is not None else sum(p['langd'] or 0 for p in st)
        schakt = e.get('schakt_m') if e.get('schakt_m') is not None else sum(
            (p['langd'] or 0) for p in st if p['metod'] == 'schakt')
        if 'kostnad_kr' in e:
            kostnad = e.get('kostnad_kr')
        else:
            k = [p['kostnad'] for p in st if p['kostnad']]
            kostnad = sum(k) if k and metod == 'strumpa' else None
        brunnar = e.get('brunnar') or sorted(set(b for p in st for b in (p['fran'], p['till']) if b))
        fram = e.get('framschaktade') or []
        fram_man = set(e.get('framschakt_manuell') or [])
        dims = e.get('dimensioner') or []
        strackor = e.get('strackor') or [{'fil': p['fil'], 'nr': p['nr'], 'startbrunn': p['fran'],
                                         'slutbrunn': p['till']} for p in st]
        str_text = '; '.join('%s nr %s: %s-%s' % (os.path.splitext(txt(s.get('fil') or ''))[0], s.get('nr'),
                                                  s.get('startbrunn'), s.get('slutbrunn')) for s in strackor)
        etikett = 'Etapp %d - %s, %d m' % (nr, metod or '?', round(langd or 0))
        if kostnad:
            etikett += ', ' + _kr(kostnad)
        if fram:
            etikett += ', %d brunn%s att schakta fram' % (len(fram), 'ar' if len(fram) != 1 else '')
        rad = [geom, nr, metod, hogsta, idx, len(st), langd, schakt, kostnad,
               ', '.join(txt(d) for d in dims), ', '.join(brunnar), ', '.join(fram), len(fram),
               e.get('serviser_uppstroms'), '; '.join(e.get('flaggor') or []), str_text, etikett]
        rad = [v[:max_text] if isinstance(v, (str, type(u''))) and typ == 'TEXT' else v
               for v, typ in zip(rad, ['SHAPE@'] + typer)]
        insert.insertRow(_utan_null(rad, ['SHAPE@'] + typer, ar_shapefil))
        n += 1
        # brunnar att schakta fram: laget ur strackornas andpunkter
        for b in fram:
            xy = _brunnsposition(b, st)
            if xy is not None:
                brunnspunkter.append((b, nr, 'Ja' if b in fram_man else '', xy))
            else:
                logg('  brunnen %s (etapp %d) hittades inte bland strackornas andpunkter' % (b, nr))
    del insert
    logg('  %d etapper skrivna till %s' % (n, ut_fil))
    ut = {'etapper': (ut_fil, n), 'brunnar': None}

    if brunnar_ut:
        b_fil, b_shp, b_max = _skapa_fc(brunnar_ut, 'POINT', BRUNN_FALT, sr)
        b_namn = [f[0] for f in BRUNN_FALT]
        b_typer = [f[1] for f in BRUNN_FALT]
        ins = arcpy.da.InsertCursor(b_fil, ['SHAPE@'] + b_namn)
        for b, nr, man, (x, y) in brunnspunkter:
            ins.insertRow(_utan_null([arcpy.PointGeometry(arcpy.Point(x, y), sr), b, nr, man,
                                      '%s (etapp %d)' % (b, nr)], ['SHAPE@'] + b_typer, b_shp))
        del ins
        logg('  %d brunnar att schakta fram skrivna till %s' % (len(brunnspunkter), b_fil))
        ut['brunnar'] = (b_fil, len(brunnspunkter))
    elif brunnspunkter:
        logg('  %d brunnar att schakta fram (ange ett punktlager for att fa dem i kartan)' % len(brunnspunkter))

    if lagg_till_i_kartan:
        _lagg_i_kartan(ut, lyr_fil, lyr_brunnar)
    return ut


def _brunnsposition(littera, strackor):
    """(x, y) for brunnen: start- eller slutpunkt pa en stracka i etappen med det litterat."""
    n = normalisera(littera)
    for p in strackor:
        if not p['delar']:
            continue
        if normalisera(p['fran']) == n:
            q = p['delar'][0][0]
            return (q.X, q.Y)
        if normalisera(p['till']) == n:
            q = p['delar'][-1][-1]
            return (q.X, q.Y)
    return None


def _lagg_i_kartan(ut, lyr_fil, lyr_brunnar):
    mxd = _mxd()
    if mxd is None:
        return
    try:
        df = arcpy.mapping.ListDataFrames(mxd)[0]
        if ut.get('brunnar'):
            l_b = arcpy.mapping.Layer(ut['brunnar'][0])
            l_b.name = LAGERNAMN_BRUNNAR
            lyr_b = lyr_brunnar or LYR_FRAMSCHAKT
            if lyr_b and os.path.isfile(lyr_b):
                arcpy.ApplySymbologyFromLayer_management(l_b, lyr_b)
            arcpy.mapping.AddLayer(df, l_b, 'TOP')
        l_e = arcpy.mapping.Layer(ut['etapper'][0])
        l_e.name = LAGERNAMN
        lyr_e = lyr_fil or LYR_ETAPPER
        if lyr_e and os.path.isfile(lyr_e):
            arcpy.ApplySymbologyFromLayer_management(l_e, lyr_e)
            logg('  symbologi applicerad fran %s' % lyr_e)
        else:
            logg('  Tips: satt symbologi (Unique values pa ETAPP eller en farg per METOD) och etiketter '
                 'pa ETIKETT, spara som %s sa anvands den nasta gang' % LYR_ETAPPER)
        arcpy.mapping.AddLayer(df, l_e, 'TOP')
        arcpy.RefreshTOC()
        arcpy.RefreshActiveView()
    except Exception as e:
        logg('  kunde inte lagga till lagret i kartan: %s' % txt(e))
