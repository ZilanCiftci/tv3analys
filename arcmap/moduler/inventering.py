# -*- coding: utf-8 -*-
"""
Ledningsinventering (ArcMap 10.x, Python 2.7 + arcpy) - steg 2 i PLAN_ledningsinventering.md.

Bygger pa gisexport.bygg_data (samma brunn/ledningsmodell som gisdata.json, sa att tv3_analys.py
kan lasa filen som gis:) och lagger till det som behovs for Excel och kalkyl per ledningsstracka
(brunn -> brunn eller ledningsande utan brunn):

  - alla originalfalt under 'attribut' (EntityID, PipeType, SewerFunction, Commentary ...)
  - lutning ur vattengangarna (promille, positiv = fall i ritad riktning) och GIS-faltet (Slope)
    med flaggorna bakfall, flack (< FLACK_PROMILLE), vg saknas, vg orimlig
  - markhojd var INTERVALL m langs linjen (projektering.Markhojd: raster/hojdmodell eller punktlager)
    och djup = mark - vattengang i samma punkt: medel, min, max, djupklass (DJUPKLASSER), tackning;
    saknas markhojd anvands locknivaerna i andbrunnarna som rat linje (djup_kalla 'lockniva');
    vattenledningar (tryck, utan vattengang) far schablondjupet vatten_djup till rorets centrum
  - serviser: varje servis vars ande ligger inom toleransen fran en huvudledning kopplas till den
    narmaste; per ledning antal, summerad langd och lista; okopplade serviser i datakvalitet
  - parallella ledningar: linjen samplas var INTERVALL m och alla andra ledningar inom
    parallell_avstand (c-c) soks; andel av langden >= parallell_andel (for nagon av de tva) =
    samschakt; sammanhangande samschakt-par bildar schaktgrupper (S + D + V i samma gata)
  - korsningar med andra ledningar (geometrisk skarning, inte vid gemensam brunn/ande) med
    vattengang/niva pa bada och hojdskillnaden
  - brunnar: lockniva, bottenniva, antal anslutna ledningar, lagsta vattengang och djup
  - anordningar (punktlager: ventiler, servisventiler, brandposter, ovriga anordningar, ledningsandar):
    kategori ur lagernamnet (och typkoden VSeV), kopplas till narmaste huvudledning av ratt slag
    (V-lager -> vattenledning, A-lager -> avlopp) inom anordning_radie; servisventiler och brandposter
    kopplas via servisen/brandpostledningen nar den ligger narmare; per ledning antal per kategori

Utdata: ledningsdata.json + ledningsdata_ledningar.csv, _brunnar.csv, _serviser.csv,
_parallella.csv, _korsningar.csv, _anordningar.csv (semikolon, decimalkomma, UTF-8 med BOM).

Testas utan ArcMap med en latsas-arcpy (scratchpad/test_inventering.py).
"""
from __future__ import unicode_literals

import os
import re
import datetime

import arcpy

from skapa_ledningslager import txt, logg, hitta_lager, kalla, _avst, _punkt_segment, TEXTTYP
from natverk import _langd, _sr_namn
from markprofil import _langs, _station
import gisexport
from gisexport import _tal, _heltal, _text, _faltnamn, _ovriga_falt, _json_varde, _vg_vid, _skriv_csv, skriv_json

VERSION = 1

FLACK_PROMILLE = 2.0          # lutning under detta (i belopp) flaggas "flack"
ORIMLIG_PROMILLE = 200.0      # lutning over detta flaggas "vg orimlig" (troligen fel niva i GIS)
DJUPKLASSER = (2.0, 3.0, 4.0)  # granser for djupklasserna 0-2, 2-3, 3-4, > 4 m (F13)
PARALLELL_MIN_ANDEL = 0.2     # par dar ingen av ledningarna har minst denna andel inom avstandet listas inte
                              # (korsningar och anslutningar snett intill ger nagra procent)
SERVIS_SYSTEM = {'S': 'S', 'D': 'D', 'K': 'K', 'V': 'V'}
# Anordningar: faltet med typkoden (forsta som finns i lagret), dimensionsfaltet och kategorierna.
# Kategorin tas ur lagernamnet (nyckelorden provas i ordning), annars 'anordning'.
ANORDNING_TYPFALT = ('ValveFunction', 'HydrantType', 'ApplianceType', 'PipeEndType', 'TYP', 'FUNKTION')
ANORDNING_DIMFALT = ('Dimension', 'DIMENSION', 'DIM')
ANORDNING_KATEGORIER = (('servisventil', ('servisventil',)),
                        ('ventil', ('ventil',)),
                        ('brandpost', ('brandpost', 'hydrant')),
                        ('ledningsande', ('ledningsande', 'ledningsände', 'pipeend')))
ANORDNING_TYPKODER = {'VSEV': 'servisventil'}   # typkod som avgor kategorin oavsett lagernamn
ANORDNING_VIA_SERVIS = ('servisventil', 'brandpost')   # sitter pa servis/brandpostledning


# ------------------------------------------------ smahjalp

def _lagernamn(lyr):
    """Samma nyckel som gisexport anvander for 'lager' (langt namn nar det finns)."""
    return txt(getattr(lyr, 'longName', None) or getattr(lyr, 'name', lyr))


def _lista(x):
    if not x:
        return []
    if isinstance(x, (TEXTTYP, bytes)):
        return [x]
    return list(x)


def _system(post, vatten):
    """Ledningssystem: V for vattenlager, annars forsta bokstaven i ledningstypen (S/D/K)."""
    if post.get('lager') in vatten:
        return 'V'
    typ = _text(post.get('ledningstyp')).upper()
    return typ[0] if typ else ''


def _djupklass(djup):
    if djup is None:
        return ''
    forra = 0.0
    for g in DJUPKLASSER:
        if djup < g:
            return '%s–%s' % (_heltal_text(forra), _heltal_text(g))
        forra = g
    return '> %s' % _heltal_text(forra)


def _heltal_text(v):
    return ('%d' % v) if float(v).is_integer() else ('%.1f' % v).replace('.', ',')


def _medel(lista):
    return sum(lista) / len(lista) if lista else None


def _runda(v, n=2):
    return round(v, n) if v is not None else None


def _skarning(a, b, c, d):
    """Skarningspunkten mellan segmenten a-b och c-d: (x, y, t, u) med t, u i [0, 1], annars None."""
    r = (b[0] - a[0], b[1] - a[1])
    s = (d[0] - c[0], d[1] - c[1])
    n = r[0] * s[1] - r[1] * s[0]
    if abs(n) < 1e-12:
        return None
    qp = (c[0] - a[0], c[1] - a[1])
    t = (qp[0] * s[1] - qp[1] * s[0]) / n
    u = (qp[0] * r[1] - qp[1] * r[0]) / n
    if t < -1e-9 or t > 1 + 1e-9 or u < -1e-9 or u > 1 + 1e-9:
        return None
    return a[0] + t * r[0], a[1] + t * r[1], t, u


class _Rutnat(object):
    """Rutnat over linjesegment: alla poster med ett segment inom radien fran en punkt."""

    def __init__(self, cell):
        self.cell = float(cell)
        self.rn = {}

    def lagg_till(self, pts, nyckel):
        c = self.cell
        for i in range(len(pts) - 1):
            a, b = pts[i], pts[i + 1]
            for cx in range(int(min(a[0], b[0]) // c), int(max(a[0], b[0]) // c) + 1):
                for cy in range(int(min(a[1], b[1]) // c), int(max(a[1], b[1]) // c) + 1):
                    self.rn.setdefault((cx, cy), []).append((a, b, nyckel))

    def inom(self, x, y, radie):
        """{nyckel: (minsta avstand, narmaste punkt)} for alla nycklar med ett segment inom radien."""
        c = self.cell
        cx, cy = int(x // c), int(y // c)
        r = int(radie // c) + 1
        ut = {}
        for dx in range(-r, r + 1):
            for dy in range(-r, r + 1):
                for a, b, k in self.rn.get((cx + dx, cy + dy), ()):
                    d, t, q = _punkt_segment(x, y, a[0], a[1], b[0], b[1])
                    if d <= radie and (k not in ut or d < ut[k][0]):
                        ut[k] = (d, q)
        return ut

    def segment_nara(self, a, b):
        """Alla (segment, nyckel) i cellerna som segmentet a-b beror."""
        c = self.cell
        ut = []
        for cx in range(int(min(a[0], b[0]) // c), int(max(a[0], b[0]) // c) + 1):
            for cy in range(int(min(a[1], b[1]) // c), int(max(a[1], b[1]) // c) + 1):
                ut.extend(self.rn.get((cx, cy), ()))
        return ut


# ------------------------------------------------ serviser

def _las_serviser(servislager, omrade, dim_falt, mat_falt, ledntyp_falt, id_falt, sr):
    # omrade begransar bara vattenlagren (de klippta vattenledningarna); avloppsserviser langs hela
    # avloppsledningar (aven utanfor omradet) ska med
    """[{lager, oid, id, pts, langd_m, dimension, material, ledningstyp, attribut}] ur servislagren."""
    ut = []
    for lyr in servislager:
        namn = _lagernamn(lyr)
        falt = _faltnamn(lyr)

        def f(onskat):
            return falt.get(txt(onskat).upper()) if onskat else None
        f_dim, f_mat, f_typ, f_id = f(dim_falt), f(mat_falt), f(ledntyp_falt), f(id_falt)
        ovriga = _ovriga_falt(lyr)
        lasfalt = ['OID@', 'SHAPE@'] + ovriga
        src, dq = kalla(lyr)
        projicera = sr is not None and _sr_namn(arcpy.Describe(src)) not in ('', txt(getattr(sr, 'name', '')))
        arcpy.MakeFeatureLayer_management(src, 'lyr_inv_serv', dq)
        if omrade is not None and _lagersystem(lyr) == 'V':
            arcpy.SelectLayerByLocation_management('lyr_inv_serv', 'INTERSECT', omrade, '', 'NEW_SELECTION')
        n = 0
        with arcpy.da.SearchCursor('lyr_inv_serv', lasfalt) as mark:
            for rad in mark:
                oid, geom = rad[0], rad[1]
                if geom is None:
                    continue
                if projicera:
                    geom = geom.projectAs(sr)
                v = dict(zip(ovriga, rad[2:]))
                pts = []
                for del_ in geom:
                    pts.extend((p.X, p.Y) for p in del_ if p is not None)
                if len(pts) < 2:
                    continue
                ut.append({
                    'lager': namn, 'oid': oid,
                    'id': _text(v.get(f_id)) if f_id else '%s/%s' % (namn, oid),
                    'pts': pts, 'langd_m': round(_langd(pts), 2),
                    'dimension': _heltal(v.get(f_dim)) if f_dim else None,
                    'material': _text(v.get(f_mat)) if f_mat else '',
                    'ledningstyp': _text(v.get(f_typ)) if f_typ else '',
                    'attribut': dict((k, _json_varde(v.get(k))) for k in ovriga),
                })
                n += 1
        arcpy.Delete_management('lyr_inv_serv')
        logg('  %s: %d serviser' % (namn, n))
    return ut


def _koppla_serviser(serviser, ledningar, rutnat, tolerans, intervall):
    """Kopplar varje servis till huvudledningen narmast nagon av dess andar (inom toleransen)."""
    n_ok = 0
    for s in serviser:
        bast = None
        for p in (s['pts'][0], s['pts'][-1]):
            for k, (d, q) in rutnat.inom(p[0], p[1], tolerans).items():
                if bast is None or d < bast[0]:
                    bast = (d, k, q, p)
        s['ledning'] = None
        s['position_m'] = None
        if bast is None:
            continue
        d, k, q, p = bast
        led = ledningar[k]
        st, _ = _station(led['_pts'], p)
        s['ledning'] = led['id']
        s['ledning_index'] = k
        s['system'] = led['system']
        s['position_m'] = round(st, 2)
        s['vg_vid_anslutning'] = _runda(_vg_vid(st, led['langd_m'], led['vg_fran'], led['vg_till']), 3)
        s['djup_vid_anslutning'] = None
        prof = led.get('profil') or []
        if prof:
            narmast = min(prof, key=lambda r: abs(r[0] - st))
            if narmast[1] is not None and narmast[2] is not None:
                s['djup_vid_anslutning'] = round(narmast[1] - narmast[2], 2)
        led['serviser'].append(s['id'])
        led['antal_serviser'] += 1
        led['serviser_m'] = round(led['serviser_m'] + s['langd_m'], 2)
        n_ok += 1
    return n_ok


# ------------------------------------------------ anordningar

def _kategori(lagernamn, typ):
    kod = _text(typ).upper()
    if kod in ANORDNING_TYPKODER:
        return ANORDNING_TYPKODER[kod]
    n = re.sub(r'\(\s*(ej|ej\.|inte|utan)\s[^)]*\)', '', txt(lagernamn).lower())   # "V Ventil (ej servisventil)"
    for kat, ord_ in ANORDNING_KATEGORIER:
        if any(o in n for o in ord_):
            return kat
    return 'anordning'


def _lagersystem(lyr):
    """'V' for lager i vattengruppen eller med namn 'V ...', 'A' for avlopp, annars ''."""
    kort = txt(getattr(lyr, 'name', '') or '').strip().lower()
    langt = txt(getattr(lyr, 'longName', '') or '').strip().lower()
    if kort.startswith('v ') or langt.startswith('vatten'):
        return 'V'
    if kort.startswith('a ') or langt.startswith('avlopp'):
        return 'A'
    return ''


def _las_anordningar(lagerlista, omrade, ar_falt, id_falt, sr):
    """[{id, lager, oid, kategori, typ, dimension, anlaggningsar, lagersystem, x, y, attribut}] ur
    punktlagren. Samma id i flera lager (t.ex. V Brandpost och V Brandpost Avstangningskarta) raknas en gang."""
    ut, sedda, n_dubbel = [], set(), 0
    for lyr in lagerlista:
        namn = _lagernamn(lyr)
        falt = _faltnamn(lyr)

        def f(onskat):
            return falt.get(txt(onskat).upper()) if onskat else None

        def forsta(kand):
            for k in kand:
                if falt.get(k.upper()):
                    return falt[k.upper()]
            return None
        f_typ, f_dim = forsta(ANORDNING_TYPFALT), forsta(ANORDNING_DIMFALT)
        f_ar, f_id = f(ar_falt), f(id_falt)
        lsys = _lagersystem(lyr)
        ovriga = _ovriga_falt(lyr)
        src, dq = kalla(lyr)
        projicera = sr is not None and _sr_namn(arcpy.Describe(src)) not in ('', txt(getattr(sr, 'name', '')))
        arcpy.MakeFeatureLayer_management(src, 'lyr_inv_anord', dq)
        if omrade is not None and _lagersystem(lyr) == 'V':      # som serviserna: bara vatten begransas
            arcpy.SelectLayerByLocation_management('lyr_inv_anord', 'INTERSECT', omrade, '', 'NEW_SELECTION')
        n = 0
        with arcpy.da.SearchCursor('lyr_inv_anord', ['OID@', 'SHAPE@'] + ovriga) as mark:
            for rad in mark:
                oid, geom = rad[0], rad[1]
                if geom is None:
                    continue
                if projicera:
                    geom = geom.projectAs(sr)
                pkt = geom.firstPoint
                if pkt is None:
                    continue
                v = dict(zip(ovriga, rad[2:]))
                ident = _text(v.get(f_id)) if f_id else ''
                nyckel = ident or '%s/%s' % (namn, oid)
                if nyckel in sedda:
                    n_dubbel += 1
                    continue
                sedda.add(nyckel)
                typ = _text(v.get(f_typ)) if f_typ else ''
                ut.append({
                    'id': nyckel, 'lager': namn, 'oid': oid, 'kategori': _kategori(namn, typ), 'typ': typ,
                    'dimension': _heltal(v.get(f_dim)) if f_dim else None,
                    'anlaggningsar': _heltal(v.get(f_ar)) if f_ar else None,
                    'lagersystem': lsys, 'x': round(pkt.X, 3), 'y': round(pkt.Y, 3),
                    'attribut': dict((k, _json_varde(v.get(k))) for k in ovriga),
                })
                n += 1
        arcpy.Delete_management('lyr_inv_anord')
        logg('  %s: %d anordningar' % (namn, n))
    if n_dubbel:
        logg('  %d anordningar fanns i flera lager och raknas en gang' % n_dubbel)
    return ut


def _koppla_anordningar(anordningar, ledningar, rutnat, serviser, radie):
    """Kopplar varje anordning till narmaste huvudledning av ratt slag inom radien; servisventiler och
    brandposter via servisen/brandpostledningen nar den ligger narmare an huvudledningen."""
    def passar(lsys, system):
        return not lsys or (system == 'V') == (lsys == 'V')
    srut = _Rutnat(10.0)
    for i, s in enumerate(serviser):
        if s.get('ledning') is not None:
            srut.lagg_till(s['pts'], i)
    n_ok = 0
    for a in anordningar:
        x, y = a['x'], a['y']
        led = None
        for k, (d, q) in rutnat.inom(x, y, radie).items():
            if passar(a['lagersystem'], ledningar[k]['system']) and (led is None or d < led[0]):
                led = (d, k)
        serv = None
        for k, (d, q) in srut.inom(x, y, radie).items():
            if passar(a['lagersystem'], serviser[k].get('system')) and (serv is None or d < serv[0]):
                serv = (d, k)
        a['ledning'] = a['system'] = a['position_m'] = a['avstand_m'] = a['via_servis'] = None
        if serv is not None and (led is None or (a['kategori'] in ANORDNING_VIA_SERVIS and serv[0] < led[0])):
            s = serviser[serv[1]]
            p = ledningar[s['ledning_index']]
            a['via_servis'] = s['id']
            a['position_m'] = s['position_m']
            a['avstand_m'] = round(serv[0], 2)
        elif led is not None:
            p = ledningar[led[1]]
            st, _ = _station(p['_pts'], (x, y))
            a['position_m'] = round(st, 2)
            a['avstand_m'] = round(led[0], 2)
        else:
            continue
        a['ledning'] = p['id']
        a['system'] = p['system']
        p['anordningar'].append(a['id'])
        p['anordningar_antal'][a['kategori']] = p['anordningar_antal'].get(a['kategori'], 0) + 1
        n_ok += 1
    return n_ok


# ------------------------------------------------ huvudfunktionen

def inventera(ledningslager, brunnslager, brunn_id, json_ut, servislager=None, vattenlager=None,
              markhojd=None, z_falt=None, lock_falt=None, botten_falt=None, typ_falt=None,
              vg_fran=None, vg_till=None, dim_falt=None, mat_falt=None, ledntyp_falt=None,
              ar_falt=None, ren_falt=None, lutning_falt=None, id_falt=None, tolerans=1.0,
              parallell_avstand=3.0, parallell_andel=0.6, intervall=1.0, vatten_djup=1.7,
              sokradie=5.0, hojdsystem='RH2000', skriv_csv=True, omradeslager=None,
              anordningslager=None, anordning_radie=3.0, klipp_vatten=True):
    """Inventerar ledningsnatet och skriver ledningsdata.json (+ CSV). Returnerar sammanfattningen."""
    import time
    t0 = time.time()

    def tid():
        return '%.0f s' % (time.time() - t0)
    if not float(tolerans) > 0:
        raise RuntimeError('Toleransen maste vara storre an 0 m')
    if not float(intervall) > 0:
        raise RuntimeError('Intervallet maste vara storre an 0 m')
    ledningslager = _lista(ledningslager)
    vattenlager = _lista(vattenlager)
    servislager = _lista(servislager)
    anordningslager = _lista(anordningslager)
    vatten = set(_lagernamn(hitta_lager(n)) for n in vattenlager)
    alla_led = ledningslager + [n for n in vattenlager if n not in ledningslager]

    logg('Ledningsinventering')
    klipp = None
    if klipp_vatten and omradeslager and vatten:
        # vattenledningar klipps vid omradesgransen (avlopp behalls hela, brunn till brunn)
        from omrade_export import _polygoner, _union
        klipp = _union([g for _, g in _polygoner(hitta_lager(omradeslager), None, True)])
        logg('  vattenledningarna klipps vid omradesgransen')
    data, hjalp = gisexport.bygg_data(
        alla_led, _lista(brunnslager), brunn_id, lock_falt=lock_falt, typ_falt=typ_falt,
        vg_fran=vg_fran, vg_till=vg_till, dim_falt=dim_falt, mat_falt=mat_falt, ledntyp_falt=ledntyp_falt,
        ar_falt=ar_falt, tolerans=tolerans, omradeslager=omradeslager, hojdsystem=hojdsystem,
        ren_falt=ren_falt, botten_falt=botten_falt, alla_falt=True,
        klipp_lager=sorted(vatten), klipp_omrade=klipp, brunnar_i_omrade=False)
    ledningar, brunnar = data['ledningar'], data['brunnar']
    sr = hjalp['sr']
    brunn_index = dict((b['littera'], b) for b in brunnar)
    omrade = None
    if omradeslager:
        src, dq = kalla(hitta_lager(omradeslager))
        arcpy.MakeFeatureLayer_management(src, 'lyr_inv_omr', dq)
        omrade = 'lyr_inv_omr'

    # --- id, system, lutning ------------------------------------------------------------
    f_id = txt(id_falt).upper() if id_falt else ''
    f_lut = txt(lutning_falt).upper() if lutning_falt else ''
    kvalitet = []

    def flagga(post, typ, text):
        post['flaggor'].append(typ)
        kvalitet.append({'typ': typ, 'id': post['id'], 'text': text})
    for i, p in enumerate(ledningar):
        at = p.get('attribut', {})
        upp = dict((k.upper(), v) for k, v in at.items())
        grund = _text(upp.get(f_id)) if f_id else ''
        if not grund:
            grund = '%s/%s' % (p['lager'].split('\\')[-1], p['oid'])
        p['id'] = grund if p['antal_delar'] == 1 else '%s/%d' % (grund, p['del'])
        p['index'] = i
        p['system'] = _system(p, vatten)
        p['flaggor'] = []
        p['lutning_gis'] = _tal(upp.get(f_lut)) if f_lut else None
        p['lutning_promille'] = None
        p['serviser'], p['antal_serviser'], p['serviser_m'] = [], 0, 0.0
        p['anordningar'], p['anordningar_antal'] = [], {}
        if p['system'] == 'V':
            continue
        if p['vg_fran'] is None or p['vg_till'] is None:
            flagga(p, 'vg saknas', 'vattengang saknas i %s' % ('bada andarna' if p['vg_fran'] is None and p['vg_till'] is None
                                                             else ('startanden' if p['vg_fran'] is None else 'slutanden')))
            continue
        if p['langd_m'] > 0:
            lut = (p['vg_fran'] - p['vg_till']) / p['langd_m'] * 1000.0
            p['lutning_promille'] = round(lut, 1)
            if abs(lut) > ORIMLIG_PROMILLE:
                flagga(p, 'vg orimlig', 'lutning %.0f promille - fel niva i GIS?' % lut)
            elif lut < 0:
                flagga(p, 'bakfall', 'vattengangen stiger %.1f promille i ritad riktning' % -lut)
            elif lut < FLACK_PROMILLE:
                flagga(p, 'flack', 'lutning %.1f promille' % lut)
    logg('  %d ledningar: %s (%s)' % (len(ledningar), ', '.join(
        '%s %d' % (k or '?', v) for k, v in sorted(_rakna(ledningar, 'system').items())), tid()))

    # --- markhojd och djup --------------------------------------------------------------
    mark = None
    if markhojd:
        logg('Laser markhojder')
        xs = [q[0] for p in ledningar for q in p['_pts']]
        ys = [q[1] for p in ledningar for q in p['_pts']]
        bbox = (min(xs), min(ys), max(xs), max(ys)) if xs else None
        import projektering
        mark = projektering.markhojd(_lista(markhojd), z_falt, sokradie, omrade, bbox)
    n_mark = n_lock = n_schablon = n_utan = 0
    for p in ledningar:
        pts = p['_pts']
        L = p['langd_m']
        b_fran, b_till = brunn_index.get(p['fran']), brunn_index.get(p['till'])
        lock_f = b_fran['lockniva'] if b_fran else None
        lock_t = b_till['lockniva'] if b_till else None
        prov = _langs([(q[0], q[1]) for q in pts], float(intervall))
        profil, djup, markh = [], [], []
        kalla_ = ''
        for m, x, y in prov:
            z = mark.hojd(x, y) if mark is not None else None
            if z is None and lock_f is not None and lock_t is not None and L > 0:
                z = lock_f + (lock_t - lock_f) * m / L
                kalla_ = kalla_ or 'lockniva'
            elif z is not None and mark is not None:
                kalla_ = 'markmodell' if kalla_ != 'lockniva' else kalla_
            if p['system'] == 'V':
                vg = None
                d = float(vatten_djup)
            else:
                vg = _vg_vid(m, L, p['vg_fran'], p['vg_till']) \
                    if (L > 0 and p['vg_fran'] is not None and p['vg_till'] is not None) else None
                d = (z - vg) if (z is not None and vg is not None) else None
            profil.append([round(m, 2), _runda(z), _runda(vg, 3)])
            if z is not None:
                markh.append(z)
            if d is not None:
                djup.append(d)
        if p['system'] == 'V':
            kalla_ = 'schablon'
            n_schablon += 1
        elif not djup:
            kalla_ = ''
            n_utan += 1
            if p['vg_fran'] is not None and p['vg_till'] is not None:
                flagga(p, 'djup saknas', 'ingen markhojd och inga locknivaer i brunnarna')
        elif kalla_ == 'lockniva':
            n_lock += 1
        else:
            n_mark += 1
        p['djup_kalla'] = kalla_
        p['djup_medel'] = _runda(_medel(djup))
        p['djup_min'] = _runda(min(djup)) if djup else None
        p['djup_max'] = _runda(max(djup)) if djup else None
        p['djupklass'] = _djupklass(p['djup_medel'])
        p['mark_medel'] = _runda(_medel(markh))
        p['mark_min'] = _runda(min(markh)) if markh else None
        p['mark_max'] = _runda(max(markh)) if markh else None
        dim_m = (p['dimension'] or 0) / 1000.0
        p['tackning_min'] = _runda(p['djup_min'] - dim_m) if p['djup_min'] is not None and p['system'] != 'V' else None
        p['djup_start'] = _runda(lock_f - p['vg_fran']) if (lock_f is not None and p['vg_fran'] is not None) else None
        p['djup_slut'] = _runda(lock_t - p['vg_till']) if (lock_t is not None and p['vg_till'] is not None) else None
        p['profil'] = profil if (mark is not None or kalla_ == 'lockniva') else []
        if p['tackning_min'] is not None and p['tackning_min'] < 0:
            flagga(p, 'ledning over mark', 'tackning %.2f m - vattengang eller markhojd fel' % p['tackning_min'])
    logg('  djup: %d ur markmodell, %d ur locknivaer, %d schablon (vatten), %d utan (%s)'
         % (n_mark, n_lock, n_schablon, n_utan, tid()))

    # --- rutnat over huvudledningarna ---------------------------------------------------
    rutnat = _Rutnat(max(10.0, 2 * float(parallell_avstand)))
    for p in ledningar:
        rutnat.lagg_till(p['_pts'], p['index'])

    # --- serviser -----------------------------------------------------------------------
    serviser = []
    if servislager:
        logg('Laser serviser')
        serviser = _las_serviser([hitta_lager(n) for n in servislager], omrade if klipp is not None else None,
                                 dim_falt, mat_falt,
                                 ledntyp_falt, id_falt, sr)
        n_ok = _koppla_serviser(serviser, ledningar, rutnat, float(tolerans), float(intervall))
        for s in serviser:
            if s['ledning'] is None:
                kvalitet.append({'typ': 'servis utan huvudledning', 'id': s['id'],
                                 'text': 'ingen huvudledning inom %.1f m fran nagon ande' % float(tolerans)})
        logg('  %d av %d serviser kopplade till en huvudledning (%s)' % (n_ok, len(serviser), tid()))

    # --- anordningar (ventiler, brandposter ...) -------------------------------------------
    anordningar = []
    if anordningslager:
        logg('Laser anordningar (ventiler, brandposter ...)')
        anordningar = _las_anordningar([hitta_lager(n) for n in anordningslager],
                                       omrade if klipp is not None else None, ar_falt, id_falt, sr)
        n_ok = _koppla_anordningar(anordningar, ledningar, rutnat, serviser, float(anordning_radie))
        for a in anordningar:
            if a['ledning'] is None:
                kvalitet.append({'typ': 'anordning utan ledning', 'id': a['id'],
                                 'text': '%s: ingen ledning inom %.1f m' % (a['kategori'], float(anordning_radie))})
        logg('  %d av %d anordningar kopplade till en ledning: %s (%s)' % (n_ok, len(anordningar), ', '.join(
            '%s %d' % (k, v) for k, v in sorted(_rakna(anordningar, 'kategori').items())), tid()))

    # --- parallella ledningar -----------------------------------------------------------
    logg('Soker parallella ledningar (c-c inom %.1f m)' % float(parallell_avstand))
    par = {}        # (i, j) i<j -> {'n_i', 'n_j', 'd': [avstand]}
    n_prov = {}
    for p in ledningar:
        i = p['index']
        pts = [(q[0], q[1]) for q in p['_pts']]
        prov = _langs(pts, float(intervall))
        n_prov[i] = len(prov)
        andar = (pts[0], pts[-1])
        for m, x, y in prov:
            for j, (d, q) in rutnat.inom(x, y, float(parallell_avstand)).items():
                if j == i:
                    continue
                o = [(r[0], r[1]) for r in (ledningar[j]['_pts'][0], ledningar[j]['_pts'][-1])]
                # gemensam brunn/ande: provpunkter nara den raknas inte (anslutning, inte parallell)
                gemensam = [e for e in andar if any(_avst(e, g) <= float(tolerans) for g in o)]
                if any(_avst((x, y), g) <= float(parallell_avstand) for g in gemensam):
                    continue
                nyckel = (i, j) if i < j else (j, i)
                rec = par.setdefault(nyckel, {'n': {nyckel[0]: 0, nyckel[1]: 0}, 'd': []})
                rec['n'][i] += 1
                rec['d'].append(d)
    parallella = []
    for (i, j), rec in sorted(par.items()):
        a, b = ledningar[i], ledningar[j]
        andel_a = rec['n'][i] / float(max(n_prov[i], 1))
        andel_b = rec['n'][j] / float(max(n_prov[j], 1))
        if max(andel_a, andel_b) < PARALLELL_MIN_ANDEL:
            continue                                # korsning eller strax forbi - inte parallell
        gemensam_m = min(max(rec['n'][i], rec['n'][j]) * float(intervall), min(a['langd_m'], b['langd_m']))
        post = {'ledning_a': a['id'], 'ledning_b': b['id'], 'system_a': a['system'], 'system_b': b['system'],
                'andel_a': round(andel_a, 2), 'andel_b': round(andel_b, 2),
                'medelavstand_m': _runda(_medel(rec['d'])), 'gemensam_m': round(gemensam_m, 1),
                'samschakt': max(andel_a, andel_b) >= float(parallell_andel), '_i': i, '_j': j}
        parallella.append(post)
    # schaktgrupper (union-find over samschakt-paren)
    foralder = {}

    def rot(k):
        while foralder.get(k, k) != k:
            k = foralder[k]
        return k
    for post in parallella:
        if post['samschakt']:
            ra, rb = rot(post['_i']), rot(post['_j'])
            if ra != rb:
                foralder[max(ra, rb)] = min(ra, rb)
    grupper = {}
    i_sam = set()
    for post in parallella:
        if post['samschakt']:
            i_sam.update((post['_i'], post['_j']))
    for i in sorted(i_sam):
        grupper.setdefault(rot(i), []).append(i)
    gruppnr = {}
    for n, r in enumerate(sorted(grupper, key=lambda r: min(grupper[r])), 1):
        for i in grupper[r]:
            gruppnr[i] = n
    for p in ledningar:
        p['parallella'] = []
        g = gruppnr.get(p['index'])
        p['schaktgrupp'] = g
        medl = [ledningar[i] for i in grupper[rot(p['index'])]] if g else []
        p['schakt_ledningar'] = [q['id'] for q in medl if q['index'] != p['index']]
        p['schakt_system'] = '+'.join(sorted(set(q['system'] for q in medl if q['system']))) if g else ''
        p['schakt_antal'] = len(medl) if g else 1
    for post in parallella:
        for egen, annan, andel, andel_annan in ((post['_i'], post['_j'], post['andel_a'], post['andel_b']),
                                               (post['_j'], post['_i'], post['andel_b'], post['andel_a'])):
            ledningar[egen]['parallella'].append({
                'ledning': ledningar[annan]['id'], 'system': ledningar[annan]['system'],
                'andel': andel, 'andel_annan': andel_annan, 'medelavstand_m': post['medelavstand_m'],
                'gemensam_m': post['gemensam_m'], 'samschakt': post['samschakt']})
    for post in parallella:
        post['schaktgrupp'] = gruppnr.get(post['_i']) if post['samschakt'] else None
        del post['_i'], post['_j']
    n_sam = sum(1 for p in parallella if p['samschakt'])
    logg('  %d parallella par, %d samschakt, %d schaktgrupper (%s)' % (len(parallella), n_sam, len(grupper), tid()))

    # --- korsningar ---------------------------------------------------------------------
    korsningar = []
    sedda = set()
    for p in ledningar:
        i = p['index']
        pts = p['_pts']
        for k in range(len(pts) - 1):
            a, b = (pts[k][0], pts[k][1]), (pts[k + 1][0], pts[k + 1][1])
            for c, d, j in rutnat.segment_nara(a, b):
                if j <= i or (p['lager'], p['oid']) == (ledningar[j]['lager'], ledningar[j]['oid']):
                    continue
                sk = _skarning(a, b, c, d)
                if sk is None:
                    continue
                x, y = sk[0], sk[1]
                q = ledningar[j]
                nara_ande = any(_avst((x, y), (e[0], e[1])) <= float(tolerans)
                                for e in (pts[0], pts[-1], q['_pts'][0], q['_pts'][-1]))
                if nara_ande:
                    continue                        # anslutning i brunn/ande, ingen korsning
                nyckel = (i, j, round(x, 1), round(y, 1))
                if nyckel in sedda:
                    continue
                sedda.add(nyckel)
                korsningar.append(_korsning(p, q, x, y, float(vatten_djup)))
    korsningar.sort(key=lambda k: (k['ledning_a'], k['station_a']))
    for p in ledningar:
        p['korsningar'] = []
    for k in korsningar:
        ledningar[k['_i']]['korsningar'].append({'ledning': k['ledning_b'], 'system': k['system_b'],
                                                 'station_m': k['station_a'], 'hojdskillnad_m': k['hojdskillnad_m']})
        ledningar[k['_j']]['korsningar'].append({'ledning': k['ledning_a'], 'system': k['system_a'],
                                                 'station_m': k['station_b'],
                                                 'hojdskillnad_m': -k['hojdskillnad_m'] if k['hojdskillnad_m'] is not None else None})
        del k['_i'], k['_j']
    for p in ledningar:
        p['korsningar'].sort(key=lambda k: k['station_m'])
    logg('  %d korsningar mellan ledningar (%s)' % (len(korsningar), tid()))

    # --- brunnar ------------------------------------------------------------------------
    for b in brunnar:
        b['antal_ledningar'] = 0
        b['system'] = set()
        b['vg_min'] = None
        b.setdefault('bottenniva', None)
    for p in ledningar:
        for namn, vg in ((p['fran'], p['vg_fran']), (p['till'], p['vg_till'])):
            b = brunn_index.get(namn)
            if b is None:
                continue
            b['antal_ledningar'] += 1
            if p['system']:
                b['system'].add(p['system'])
            if vg is not None and (b['vg_min'] is None or vg < b['vg_min']):
                b['vg_min'] = vg
    for b in brunnar:
        b['system'] = '+'.join(sorted(b['system']))
        b['djup_m'] = _runda(b['lockniva'] - b['vg_min']) if (b['lockniva'] is not None and b['vg_min'] is not None) else None
        if b['djup_m'] is None and b['bottenniva'] is not None and b['lockniva'] is not None:
            b['djup_m'] = _runda(b['lockniva'] - b['bottenniva'])
        if b['antal_ledningar'] == 0:
            kvalitet.append({'typ': 'brunn utan ledning', 'id': b['littera'], 'text': 'ingen ledning inom toleransen'})
        elif b['lockniva'] is None:
            kvalitet.append({'typ': 'brunn utan lockniva', 'id': b['littera'], 'text': ''})
    if omrade is not None:
        arcpy.Delete_management('lyr_inv_omr')

    # --- utdata -------------------------------------------------------------------------
    for p in ledningar:
        del p['_pts']
        del p['index']
    for s in serviser:
        s['fran_xy'] = [round(s['pts'][0][0], 3), round(s['pts'][0][1], 3)]
        s['till_xy'] = [round(s['pts'][-1][0], 3), round(s['pts'][-1][1], 3)]
        del s['pts']
        s.pop('ledning_index', None)
    langd = {}
    for p in ledningar:
        langd[p['system'] or '?'] = langd.get(p['system'] or '?', 0.0) + p['langd_m']
    samman = data.pop('sammanfattning')
    samman.update({
        'serviser': len(serviser), 'serviser_kopplade': sum(1 for s in serviser if s['ledning']),
        'serviser_m': round(sum(s['langd_m'] for s in serviser), 1),
        'langd_per_system_m': dict((k, round(v, 1)) for k, v in langd.items()),
        'parallella_par': len(parallella), 'samschakt_par': n_sam, 'schaktgrupper': len(grupper),
        'korsningar': len(korsningar),
        'anordningar': len(anordningar), 'anordningar_kopplade': sum(1 for a in anordningar if a['ledning']),
        'anordningar_per_kategori': _rakna(anordningar, 'kategori'), 'djup_ur_markmodell': n_mark, 'djup_ur_lockniva': n_lock,
        'djup_schablon': n_schablon, 'djup_saknas': n_utan, 'datakvalitet': len(kvalitet),
    })
    data.update({
        'version': VERSION,
        'kalla': 'inventering.py',
        'genererad': datetime.datetime.now().strftime('%Y-%m-%d %H:%M'),
        'parametrar': {'tolerans_m': float(tolerans), 'parallell_avstand_m': float(parallell_avstand),
                       'parallell_andel': float(parallell_andel), 'intervall_m': float(intervall),
                       'vatten_djup_m': float(vatten_djup), 'sokradie_m': float(sokradie),
                       'anordning_radie_m': float(anordning_radie),
                       'vatten_klippt_vid_omrade': klipp is not None,
                       'markhojd': '; '.join(txt(m) for m in _lista(markhojd)) if markhojd else None, 'z_falt': z_falt,
                       'djupklasser': list(DJUPKLASSER), 'flack_promille': FLACK_PROMILLE},
        'serviser': serviser,
        'anordningar': anordningar,
        'parallella': parallella,
        'korsningar': korsningar,
        'datakvalitet': kvalitet,
        'sammanfattning': samman,
    })
    data['lager'].update({'vatten': sorted(vatten), 'serviser': [txt(n) for n in servislager],
                          'anordningar': [txt(n) for n in anordningslager],
                          'bottenniva': botten_falt, 'lutning': lutning_falt, 'ledning_id': id_falt})
    skriv_json(data, json_ut)
    if skriv_csv:
        _skriv_csvfiler(json_ut, ledningar, brunnar, serviser, parallella, korsningar, anordningar)
    logg('Klart (%s): %d ledningar, %d brunnar, %d serviser, %d anordningar, %d samschakt-par i %d grupper, '
         '%d korsningar, %d anmarkningar' % (tid(), len(ledningar), len(brunnar), len(serviser), len(anordningar),
                                             n_sam, len(grupper), len(korsningar), len(kvalitet)))
    return samman


def _korsning(p, q, x, y, vatten_djup):
    """Korsningspost mellan ledningarna p och q i (x, y): station och niva pa bada."""
    def niva(led, st):
        if led['system'] == 'V':
            prof = led.get('profil') or []
            if prof:
                n = min(prof, key=lambda r: abs(r[0] - st))
                if n[1] is not None:
                    return round(n[1] - vatten_djup, 3), 'schablon'
            return None, 'saknas'
        v = _vg_vid(st, led['langd_m'], led['vg_fran'], led['vg_till'])
        return (_runda(v, 3), 'vattengang') if v is not None else (None, 'saknas')
    st_a, _ = _station(p['_pts'], (x, y))
    st_b, _ = _station(q['_pts'], (x, y))
    niva_a, kalla_a = niva(p, st_a)
    niva_b, kalla_b = niva(q, st_b)
    return {'ledning_a': p['id'], 'ledning_b': q['id'], 'system_a': p['system'], 'system_b': q['system'],
            'x': round(x, 3), 'y': round(y, 3), 'station_a': round(st_a, 2), 'station_b': round(st_b, 2),
            'niva_a': niva_a, 'niva_b': niva_b, 'niva_kalla_a': kalla_a, 'niva_kalla_b': kalla_b,
            'hojdskillnad_m': _runda(niva_a - niva_b, 3) if (niva_a is not None and niva_b is not None) else None,
            '_i': p['index'], '_j': q['index']}


def _rakna(poster, nyckel):
    ut = {}
    for p in poster:
        ut[p.get(nyckel)] = ut.get(p.get(nyckel), 0) + 1
    return ut


def _skriv_csvfiler(json_ut, ledningar, brunnar, serviser, parallella, korsningar, anordningar=()):
    stam = os.path.splitext(json_ut)[0]
    rader = []
    for p in ledningar:
        r = dict(p)
        r['parallella'] = ', '.join('%s (%s, %.0f %%, %s m)' % (q['ledning'], q['system'], 100 * q['andel'], _csvtal(q['medelavstand_m']))
                                    for q in p['parallella'])
        r['korsningar'] = ', '.join('%s vid %s m' % (q['ledning'], _csvtal(q['station_m'])) for q in p['korsningar'])
        r['serviser'] = ', '.join(p['serviser'])
        r['flaggor'] = ', '.join(p['flaggor'])
        r['schakt_ledningar'] = ', '.join(p['schakt_ledningar'])
        r['anordningar'] = ', '.join('%s %d' % (k, v) for k, v in sorted(p.get('anordningar_antal', {}).items()))
        rader.append(r)
    _skriv_csv(stam + '_ledningar.csv', [
        'id', 'system', 'fran', 'till', 'langd_m', 'dimension', 'material', 'ledningstyp', 'anlaggningsar',
        'renoveringsar', 'vg_fran', 'vg_till', 'lutning_promille', 'lutning_gis', 'djup_medel', 'djup_min',
        'djup_max', 'djupklass', 'djup_kalla', 'djup_start', 'djup_slut', 'tackning_min', 'mark_medel',
        'antal_serviser', 'serviser_m', 'serviser', 'anordningar', 'schaktgrupp', 'schakt_system', 'schakt_antal',
        'schakt_ledningar', 'parallella', 'korsningar', 'flaggor', 'omrade', 'lager', 'oid', 'del', 'antal_delar',
        'klippt_m'], rader)
    _skriv_csv(stam + '_brunnar.csv', ['littera', 'typ', 'system', 'lockniva', 'bottenniva', 'vg_min', 'djup_m',
                                       'antal_ledningar', 'omrade', 'x', 'y', 'lager'], brunnar)
    _skriv_csv(stam + '_serviser.csv', ['id', 'ledning', 'system', 'position_m', 'langd_m', 'dimension', 'material',
                                        'ledningstyp', 'vg_vid_anslutning', 'djup_vid_anslutning', 'lager', 'oid'], serviser)
    _skriv_csv(stam + '_parallella.csv', ['ledning_a', 'ledning_b', 'system_a', 'system_b', 'andel_a', 'andel_b',
                                          'medelavstand_m', 'gemensam_m', 'samschakt', 'schaktgrupp'], parallella)
    _skriv_csv(stam + '_korsningar.csv', ['ledning_a', 'ledning_b', 'system_a', 'system_b', 'station_a', 'station_b',
                                          'niva_a', 'niva_b', 'niva_kalla_a', 'niva_kalla_b', 'hojdskillnad_m', 'x', 'y'],
               korsningar)
    _skriv_csv(stam + '_anordningar.csv', ['id', 'kategori', 'typ', 'ledning', 'system', 'position_m', 'avstand_m',
                                           'via_servis', 'dimension', 'anlaggningsar', 'x', 'y', 'lager', 'oid'],
               anordningar)
    logg('Skrev %s_ledningar/_brunnar/_serviser/_parallella/_korsningar/_anordningar.csv' % stam)


def _csvtal(v):
    return '' if v is None else ('%.1f' % v).replace('.', ',')
