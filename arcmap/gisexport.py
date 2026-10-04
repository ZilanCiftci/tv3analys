# -*- coding: utf-8 -*-
"""
Export av brunnar och ledningsstrackor mellan brunnar ur kartan (ArcMap 10.x, Python 2.7 + arcpy),
oberoende av TV3-filerna. tv3_analys.py laser filen (gis: FIL i listfilen eller --gis) och
kontrollerar filmerna mot GIS: vattengang, littera, flodesriktning, material och dimension,
och raknar djup vid brunnarna (lockniva minus lagsta vattengang).

Ledningarna delas vid varje brunn som ligger inom toleransen fran linjen (samma graf som i
Uppstroms, natverk.Graf, med ALLA brunnar som noder), sa att en ledning i kartan som passerar
flera brunnar blir en stracka per brunnspar. Vattengangen i delningspunkterna interpoleras
linjart langs ledningen mellan dess egna vg-falt (som i Markprofil). Fria ledningsandar utan
brunn exporteras med tomt brunnsnamn och koordinater.

Utdata: gisdata.json (lases av analysen) och, bredvid den, brunnar.csv och ledningar.csv
for granskning i Excel (semikolon, decimalkomma, UTF-8 med BOM).

    gisdata.json
    {"version": 1, "hojdsystem": "RH2000", "tolerans_m": 1.0, "lager": {...},
     "brunnar": [{"littera", "typ", "lockniva", "x", "y", "lager"}],
     "ledningar": [{"fran", "till", "fran_xy", "till_xy", "langd_m", "vg_fran", "vg_till",
                    "dimension", "material", "ledningstyp", "anlaggningsar",
                    "lager", "oid", "del", "antal_delar"}]}

Testas utan ArcMap med en latsas-arcpy (scratchpad/test_gisexport.py).
"""
from __future__ import unicode_literals

import os
import io
import json
import datetime

import arcpy

from skapa_ledningslager import txt, logg, hitta_lager, hitta_falt, kalla, normalisera, _avst, TEXTTYP
from natverk import Graf, _langd, _sr_namn
from markprofil import _station, _punkt_vid

VERSION = 1


def _tal(v):
    try:
        if v is None or txt(v).strip() == '':
            return None
        return float(txt(v).replace(',', '.'))
    except (TypeError, ValueError):
        return None


def _heltal(v):
    t = _tal(v)
    return int(round(t)) if t is not None else None


def _text(v):
    return txt(v).strip() if v is not None else ''


def _faltnamn(lyr):
    return dict((f.name.upper(), f.name) for f in arcpy.ListFields(kalla(lyr)[0]))


def _valj_falt(falt, onskat, vad, lagernamn):
    """Verkligt faltnamn for ett onskat (skiftlagesokansligt), eller None med varning."""
    if not onskat:
        return None
    namn = falt.get(txt(onskat).upper())
    if not namn:
        logg('  VARNING: %s %s finns inte i %s - hoppas over for det lagret' % (vad, txt(onskat), lagernamn))
    return namn


def las_brunnar(brunnslager, brunn_id, lock_falt=None, typ_falt=None, sr=None, omrade=None):
    """[{littera, typ, lockniva, x, y, lager}] ur brunnslagren. Forsta forekomsten av ett littera
    vinner; dubbletter pa annan plats (> 1 m isar) varnas."""
    brunnar, index = [], {}
    for lyr in brunnslager:
        namn = txt(getattr(lyr, 'name', lyr))
        idfalt = hitta_falt(lyr, brunn_id)
        falt = _faltnamn(lyr)
        lock = _valj_falt(falt, lock_falt, 'locknivafaltet', namn)
        typ = _valj_falt(falt, typ_falt, 'brunnstypfaltet', namn)
        src, dq = kalla(lyr)
        sr_namn = txt(getattr(sr, 'name', '') or '') if sr is not None else ''
        projicera = bool(sr_namn) and _sr_namn(arcpy.Describe(src)) not in ('', sr_namn)
        if projicera:
            logg('  %s projiceras till ledningslagrets koordinatsystem' % namn)
        arcpy.MakeFeatureLayer_management(src, 'lyr_gis_br', dq)
        if omrade is not None:
            arcpy.SelectLayerByLocation_management('lyr_gis_br', 'INTERSECT', omrade, '', 'NEW_SELECTION')
        extra_falt = [f for f in (lock, typ) if f]
        if lock and typ == lock:
            extra_falt = [lock]            # samma falt for bada - inte tva ganger i SearchCursor
        lasfalt = ['SHAPE@' if projicera else 'SHAPE@XY', idfalt] + extra_falt
        n = 0
        with arcpy.da.SearchCursor('lyr_gis_br', lasfalt) as mark:
            for rad in mark:
                geom, bid = rad[0], normalisera(rad[1])
                if not bid or geom is None:
                    continue
                if projicera:
                    pnt = geom.projectAs(sr).firstPoint
                    x, y = pnt.X, pnt.Y
                else:
                    x, y = geom
                if x is None:
                    continue
                extra = dict(zip(extra_falt, rad[2:]))
                post = {'littera': bid, 'typ': _text(extra.get(typ)) if typ else '',
                        'lockniva': _tal(extra.get(lock)) if lock else None,
                        'x': round(x, 3), 'y': round(y, 3), 'lager': namn}
                n += 1
                if bid in index:
                    q = index[bid]
                    if _avst((x, y), (q['x'], q['y'])) > 1.0:
                        logg('  VARNING: brunnen %s finns pa flera stallen (%s och %s)' % (bid, q['lager'], namn))
                    elif q['lockniva'] is None and post['lockniva'] is not None:
                        q['lockniva'] = post['lockniva']
                    continue
                index[bid] = post
                brunnar.append(post)
        arcpy.Delete_management('lyr_gis_br')
        logg('  %s: %d brunnar' % (namn, n))
    return brunnar


def _vg_vid(st, L, vg_fran, vg_till):
    """Vattengang vid mattet st langs en ledning med langden L: andarnas varden i andarna,
    linjart daremellan nar bada finns."""
    if L <= 0:
        return vg_fran if vg_fran is not None else vg_till
    if st <= 0.05:
        return vg_fran
    if st >= L - 0.05:
        return vg_till
    if vg_fran is None or vg_till is None:
        return None
    return vg_fran + (vg_till - vg_fran) * st / L


def _nodnamn(nod):
    return txt(nod[1]) if nod[0] == 'B' else ''


class Omraden(object):
    """Driftomraden (polygoner) med namn: vilket omrade en punkt ligger i. Punkt-i-polygon i ren
    Python (ringar ur SHAPE@; hal = ringar efter None i en del), sa att det gar att testa utan
    arcpy-geometri. Forsta omradet som innehaller punkten vinner."""

    def __init__(self, lager, namnfalt, sr=None):
        self.poster = []     # (namn, bbox, [ringar]) dar ring = [(x, y), ...]
        namn_l = txt(getattr(lager, 'name', lager))
        falt = _faltnamn(lager)
        f_namn = _valj_falt(falt, namnfalt, 'omradesnamnfaltet', namn_l)
        src, dq = kalla(lager)
        sr_namn = txt(getattr(sr, 'name', '') or '') if sr is not None else ''
        projicera = bool(sr_namn) and _sr_namn(arcpy.Describe(src)) not in ('', sr_namn)
        if projicera:
            logg('  %s projiceras till ledningslagrets koordinatsystem' % namn_l)
        arcpy.MakeFeatureLayer_management(src, 'lyr_gis_duf', dq)
        with arcpy.da.SearchCursor('lyr_gis_duf', ['SHAPE@'] + ([f_namn] if f_namn else [])) as mark:
            for rad in mark:
                geom = rad[0]
                if geom is None:
                    continue
                if projicera:
                    geom = geom.projectAs(sr)
                namn = _text(rad[1]) if f_namn else ''
                ringar = []
                for del_ in geom:
                    ring = []
                    for p in del_:
                        if p is None:            # hal borjar
                            if len(ring) >= 3:
                                ringar.append(ring)
                            ring = []
                        else:
                            ring.append((p.X, p.Y))
                    if len(ring) >= 3:
                        ringar.append(ring)
                if not ringar:
                    continue
                xs = [x for r in ringar for x, _ in r]
                ys = [y for r in ringar for _, y in r]
                self.poster.append((namn, (min(xs), min(ys), max(xs), max(ys)), ringar))
        arcpy.Delete_management('lyr_gis_duf')
        logg('  %s: %d omraden' % (namn_l, len(self.poster)))

    @staticmethod
    def _i_ring(x, y, ring):
        inne = False
        n = len(ring)
        for i in range(n):
            x1, y1 = ring[i]
            x2, y2 = ring[(i + 1) % n]
            if (y1 > y) != (y2 > y):
                xk = x1 + (y - y1) * (x2 - x1) / (y2 - y1)
                if x < xk:
                    inne = not inne
        return inne

    def omrade(self, x, y):
        """Namnet pa omradet som innehaller (x, y), annars ''. Udda antal ringar runt punkten =
        inne (ytterring + hal)."""
        for namn, (xmin, ymin, xmax, ymax), ringar in self.poster:
            if not (xmin <= x <= xmax and ymin <= y <= ymax):
                continue
            if sum(1 for r in ringar if self._i_ring(x, y, r)) % 2 == 1:
                return namn
        return ''


def exportera(ledningslager, brunnslager, brunn_id, json_ut, lock_falt=None, typ_falt=None,
              vg_fran=None, vg_till=None, dim_falt=None, mat_falt=None, ledntyp_falt=None,
              ar_falt=None, tolerans=1.0, omradeslager=None, hojdsystem='RH2000', skriv_csv=True,
              duf_lager=None, duf_falt=None):
    """Exporterar brunnar och ledningsstrackor mellan brunnar till json_ut (+ CSV bredvid).
    duf_lager/duf_falt: polygonlager med driftomraden och namnfalt - varje ledning (mittpunkt) och
    brunn far 'omrade', som analysen anvander for inspektionsgrad per omrade.
    Returnerar en sammanfattning (dict)."""
    if not float(tolerans) > 0:
        raise RuntimeError('Toleransen maste vara storre an 0 m')
    if isinstance(ledningslager, (TEXTTYP, bytes)):
        ledningslager = [ledningslager]
    if isinstance(brunnslager, (TEXTTYP, bytes)):
        brunnslager = [brunnslager]
    led_lager = [hitta_lager(n) for n in ledningslager]
    brunn_lager = [hitta_lager(n) for n in brunnslager]
    omrade = None
    if omradeslager:
        src, dq = kalla(hitta_lager(omradeslager))
        arcpy.MakeFeatureLayer_management(src, 'lyr_gis_omr', dq)
        omrade = 'lyr_gis_omr'

    sr = arcpy.Describe(kalla(led_lager[0])[0]).spatialReference
    omraden = None
    if duf_lager:
        logg('Laser driftomraden')
        omraden = Omraden(hitta_lager(duf_lager), duf_falt, sr)
    logg('Laser brunnar')
    brunnar = las_brunnar(brunn_lager, brunn_id, lock_falt, typ_falt, sr, omrade)
    logg('  %d brunnar totalt' % len(brunnar))
    if omraden is not None:
        for b in brunnar:
            b['omrade'] = omraden.omrade(b['x'], b['y'])
    graf = Graf(dict((b['littera'], (b['x'], b['y'])) for b in brunnar), float(tolerans))

    # Ledningarna: originalpunkter och attribut per objekt(del), bitarna i grafen
    original = {}       # (lager, oid) -> [[punkter per del]]
    attr = {}           # (lager, oid) -> dict
    n_led = n_bitar = 0
    for lyr in led_lager:
        namn = txt(getattr(lyr, 'longName', None) or getattr(lyr, 'name', lyr))   # langt namn: tva lager kan heta lika
        falt = _faltnamn(lyr)
        f_vg1 = _valj_falt(falt, vg_fran, 'vattengangsfaltet (fran)', namn)
        f_vg2 = _valj_falt(falt, vg_till, 'vattengangsfaltet (till)', namn)
        f_dim = _valj_falt(falt, dim_falt, 'dimensionsfaltet', namn)
        f_mat = _valj_falt(falt, mat_falt, 'materialfaltet', namn)
        f_typ = _valj_falt(falt, ledntyp_falt, 'ledningstypfaltet', namn)
        f_ar = _valj_falt(falt, ar_falt, 'arfaltet', namn)
        extra = [f for f in (f_vg1, f_vg2, f_dim, f_mat, f_typ, f_ar) if f]
        sedd = set()
        lasfalt = ['OID@', 'SHAPE@'] + [f for f in extra if not (f in sedd or sedd.add(f))]
        src, dq = kalla(lyr)
        arcpy.MakeFeatureLayer_management(src, 'lyr_gis_led', dq)
        if omrade is not None:
            arcpy.SelectLayerByLocation_management('lyr_gis_led', 'INTERSECT', omrade, '', 'NEW_SELECTION')
        logg('  %s: %s ledningar' % (namn, arcpy.GetCount_management('lyr_gis_led').getOutput(0)))
        with arcpy.da.SearchCursor('lyr_gis_led', lasfalt) as mark:
            for rad in mark:
                oid, geom = rad[0], rad[1]
                if geom is None:
                    continue
                v = dict(zip(lasfalt[2:], rad[2:]))
                attr[(namn, oid)] = {
                    'vg_fran': _tal(v.get(f_vg1)) if f_vg1 else None,
                    'vg_till': _tal(v.get(f_vg2)) if f_vg2 else None,
                    'dimension': _heltal(v.get(f_dim)) if f_dim else None,
                    'dimension_text': _text(v.get(f_dim)) if f_dim else '',
                    'material': _text(v.get(f_mat)) if f_mat else '',
                    'ledningstyp': _text(v.get(f_typ)) if f_typ else '',
                    'anlaggningsar': _heltal(v.get(f_ar)) if f_ar else None,
                }
                delar = []
                for del_ in geom:
                    punkter = [(p.X, p.Y, None) for p in del_ if p is not None]
                    if len(punkter) < 2:
                        continue
                    delar.append(punkter)
                    n_bitar += graf.lagg_till(punkter, namn, oid)
                if delar:
                    original[(namn, oid)] = delar
                    n_led += 1
        arcpy.Delete_management('lyr_gis_led')
    if omrade is not None:
        arcpy.Delete_management('lyr_gis_omr')
    logg('  %d ledningar, %d bitar mellan brunnar/andar' % (n_led, n_bitar))

    # Bitarna -> ledningsstrackor. Varje bit finns tva ganger i grafen (bada riktningar);
    # ta den som lagrats i ritad riktning (fram_ids) och hoppa over nollbitar.
    ledningar = []
    sedda = set()
    per_objekt = {}
    n_stub = 0
    for nod, lista in graf.kanter.items():
        for annan, pts, lager, oid in lista:
            if id(pts) not in graf.fram_ids or id(pts) in graf.noll or id(pts) in sedda:
                continue
            sedda.add(id(pts))
            a = graf.kanon(nod)
            b = graf.kanon(annan)
            bit_langd = _langd(pts)
            # Stub: ledningen sticker ut hogst toleransen forbi en brunn (brunnen projiceras
            # innanfor anden) - ingen stracka mellan brunnar, hoppa over
            if bit_langd <= float(tolerans) and (a == b or a[0] != 'B' or b[0] != 'B'):
                n_stub += 1
                continue
            delar = original.get((lager, oid), [])
            # station langs den del av originalet som biten tillhor
            st1 = st2 = None
            L = 0.0
            basta = None
            for punkter in delar:
                s1, d1 = _station(punkter, pts[0])
                s2, d2 = _station(punkter, pts[-1])
                if basta is None or d1 + d2 < basta:
                    basta, st1, st2, L = d1 + d2, s1, s2, _langd(punkter)
            if st1 is not None and st2 is not None and st2 + 1e-6 < st1:
                st2 = L            # ringledning: sista biten slutar i startpunkten (station L, inte 0)
            at = attr.get((lager, oid), {})
            post = {
                'fran': _nodnamn(a), 'till': _nodnamn(b),
                'fran_xy': [round(pts[0][0], 3), round(pts[0][1], 3)],
                'till_xy': [round(pts[-1][0], 3), round(pts[-1][1], 3)],
                'langd_m': round(_langd(pts), 2),
                'vg_fran': _vg_vid(st1, L, at.get('vg_fran'), at.get('vg_till')) if st1 is not None else at.get('vg_fran'),
                'vg_till': _vg_vid(st2, L, at.get('vg_fran'), at.get('vg_till')) if st2 is not None else at.get('vg_till'),
                'dimension': at.get('dimension'),
                'dimension_text': at.get('dimension_text', ''),
                'material': at.get('material', ''),
                'ledningstyp': at.get('ledningstyp', ''),
                'anlaggningsar': at.get('anlaggningsar'),
                'lager': lager, 'oid': oid, '_st': st1 if st1 is not None else 0.0,
            }
            if omraden is not None:
                mitt = _punkt_vid(pts, _langd(pts) / 2.0)
                post['omrade'] = omraden.omrade(mitt[0], mitt[1])
            for k in ('vg_fran', 'vg_till'):
                if post[k] is not None:
                    post[k] = round(post[k], 3)
            per_objekt.setdefault((lager, oid), []).append(post)
            ledningar.append(post)
    for lista in per_objekt.values():
        lista.sort(key=lambda p: p['_st'])          # delarna i ritad ordning langs ledningen
        for i, post in enumerate(lista, 1):
            post['del'] = i
            post['antal_delar'] = len(lista)
    for post in ledningar:
        del post['_st']
    ledningar.sort(key=lambda p: (p['lager'], p['oid'], p['del']))
    if n_stub:
        logg('  %d stubbar (ledningsande hogst %.1f m forbi en brunn) hoppades over' % (n_stub, float(tolerans)))

    # Statistik
    med_brunn = set()
    for p in ledningar:
        if p['fran']:
            med_brunn.add(p['fran'])
        if p['till']:
            med_brunn.add(p['till'])
    utan_ledning = [b['littera'] for b in brunnar if b['littera'] not in med_brunn]
    fria = sum(1 for p in ledningar if not p['fran'] or not p['till'])
    utan_vg = sum(1 for p in ledningar if p['vg_fran'] is None or p['vg_till'] is None)
    utan_lock = sum(1 for b in brunnar if b['lockniva'] is None)
    if omraden is not None:
        per = {}
        for p in ledningar:
            per[p.get('omrade') or '(utanfor)'] = per.get(p.get('omrade') or '(utanfor)', 0) + 1
        logg('  ledningar per omrade: ' + ', '.join('%s %d' % (k, v) for k, v in sorted(per.items())))
    logg('  %d ledningsstrackor, %d med fri ande, %d utan vattengang i nagon ande'
         % (len(ledningar), fria, utan_vg))
    logg('  %d brunnar utan ledning inom toleransen, %d utan lockniva' % (len(utan_ledning), utan_lock))
    if utan_ledning:
        logg('    t.ex. ' + ', '.join(utan_ledning[:8]))

    data = {
        'version': VERSION,
        'kalla': 'gisexport.py',
        'genererad': datetime.datetime.now().strftime('%Y-%m-%d %H:%M'),
        'hojdsystem': hojdsystem,
        'tolerans_m': float(tolerans),
        'koordinatsystem': _sr_namn(arcpy.Describe(kalla(led_lager[0])[0])),
        'lager': {'ledningar': [txt(getattr(l, 'name', l)) for l in led_lager],
                  'brunnar': [txt(getattr(l, 'name', l)) for l in brunn_lager],
                  'brunn_id': brunn_id, 'lockniva': lock_falt, 'brunnstyp': typ_falt,
                  'vg_fran': vg_fran, 'vg_till': vg_till, 'dimension': dim_falt, 'material': mat_falt,
                  'ledningstyp': ledntyp_falt, 'anlaggningsar': ar_falt,
                  'driftomraden': txt(duf_lager) if duf_lager else None, 'omradesnamn': duf_falt},
        'brunnar': brunnar,
        'ledningar': ledningar,
    }
    mapp = os.path.dirname(os.path.abspath(json_ut))
    if mapp and not os.path.isdir(mapp):
        os.makedirs(mapp)
    with io.open(json_ut, 'w', encoding='utf-8') as f:
        f.write(json.dumps(data, ensure_ascii=False, indent=1))
    logg('Skrev %s' % json_ut)
    if skriv_csv:
        stam = os.path.splitext(json_ut)[0]
        _skriv_csv(stam + '_brunnar.csv', ['littera', 'typ', 'lockniva', 'omrade', 'x', 'y', 'lager'], brunnar)
        _skriv_csv(stam + '_ledningar.csv',
                   ['fran', 'till', 'omrade', 'langd_m', 'vg_fran', 'vg_till', 'dimension', 'material',
                    'ledningstyp', 'anlaggningsar', 'lager', 'oid', 'del', 'antal_delar'], ledningar)
        logg('Skrev %s_brunnar.csv och %s_ledningar.csv' % (stam, stam))
    return {'brunnar': len(brunnar), 'ledningar': len(ledningar), 'fria_andar': fria,
            'utan_vg': utan_vg, 'brunnar_utan_ledning': len(utan_ledning), 'utan_lockniva': utan_lock}


def _csvvarde(v):
    if v is None:
        return ''
    if isinstance(v, float):
        return ('%.3f' % v).rstrip('0').rstrip('.').replace('.', ',')
    return txt(v).replace(';', ',')


def _skriv_csv(path, kolumner, poster):
    with io.open(path, 'w', encoding='utf-8-sig', newline='') as f:
        f.write(';'.join(kolumner) + '\r\n')
        for p in poster:
            f.write(';'.join(_csvvarde(p.get(k)) for k in kolumner) + '\r\n')
