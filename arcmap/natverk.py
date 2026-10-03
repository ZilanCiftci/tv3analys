# -*- coding: utf-8 -*-
"""
Uppstromsanalys i ledningsnatet (ArcMap 10.x, Python 2.7 + arcpy).

Bygger en graf av ledningslagret dar ALLA brunnar ar noder (till skillnad fran
skapa_ledningslager, dar bara TV3-filens brunnar ar noder) och ger varje kant en
flodesriktning. Fran en startpunkt (brunn eller vald ledning) samlas allt som ligger
uppstroms: ledningar, brunnar, serviser, total langd och langsta gren.

Flodesriktning per ledning, i denna ordning:
  1. riktningsattributet (falt + vilka varden som betyder "med ritad riktning"
     respektive "mot ritad riktning"),
  2. vattengangsfalten (hogre vattengang = uppstroms),
  3. ritad riktning (startvertex = uppstroms) - loggas som antagen.

Tva lagen:
  uppstroms(...)        en startpunkt -> lager uppstroms_<littera> + sammanfattning
  uppstroms_batch(...)  bedomda-lagret fran "Skapa ledningslager" -> per stracka
                        ANT_SERV_U (serviser uppstroms) och L_UPPSTR (langd uppstroms)
                        skrivs till lagret och till en CSV som tv3_analys.py laser
                        (uppstroms: FIL i listfilen).

Serviser: ett eget servislager, eller ledningar i ledningslagret med ett visst varde i
ett typfalt. Serviser ingar inte i grafen; en servis raknas nar nagon av dess andar
ligger inom toleransen fran en uppstromsledning. Saknas bada anvands i batchlaget
fältet ANT_ANSL (anslutningar enligt TV-inspektionen) for de inspekterade strackorna
uppstroms som skattning.

Stopp: ledningar med ett visst varde i ett falt (t.ex. tryckledningar) passeras inte,
och brunnar i en stopplista (t.ex. pumpstationer) stoppar sokningen.

Testas utan ArcMap med en latsas-arcpy (scratchpad/test_natverk.py).
"""
from __future__ import unicode_literals

import os
import io
from collections import deque

import arcpy

from skapa_ledningslager import (txt, logg, hitta_lager, hitta_falt, kalla, normalisera,
                                 _avst, _punkt_segment, _tal, _mxd, Natverk, _skapa_fc,
                                 _utan_null, TEXTTYP)

UPPSTROMS_FALT = [
    ('START',      'TEXT',   50,  'Startpunkt'),
    ('FRAN_NOD',   'TEXT',   50,  'Uppströms nod'),
    ('TILL_NOD',   'TEXT',   50,  'Nedströms nod'),
    ('AVSTAND_M',  'DOUBLE', None, 'Avstånd till startpunkten (m)'),
    ('NIVA',       'LONG',   None, 'Nivå (noder från start)'),
    ('LANGD_M',    'DOUBLE', None, 'Längd (m)'),
    ('SERVIS',     'TEXT',   3,   'Servis'),
    ('RIKTN_UR',   'TEXT',   12,  'Riktning ur (attribut/vattengång/ritad)'),
    ('SRC_LAGER',  'TEXT',   100, 'Källager'),
    ('SRC_OID',    'LONG',   None, 'Käll-OID'),
]
LAGERNAMN_UPPSTROMS = 'Uppstroms'


def _langd(pts):
    return sum(_avst(pts[i], pts[i + 1]) for i in range(len(pts) - 1))


def _nodnamn(nod):
    return txt(nod[1]) if nod[0] == 'B' else 'ande %d' % nod[1]


def _lista(v):
    """Parametervarden 'A;B' / ['A','B'] / None -> set med versala, trimmade texter."""
    if v is None:
        return set()
    if isinstance(v, (TEXTTYP, bytes)):
        v = txt(v).split(';')
    return set(txt(x).strip().strip("'").upper() for x in v if txt(x).strip().strip("'"))


class Graf(Natverk):
    """Natverk med riktning per kant och attribut per ledning.

    Natverk ger varje ledningsande bade en fri-ande-nod (P) och, nar en brunn ligger dar, en
    brunnsnod (B) pa samma stalle med en bit utan langd emellan. Sadana noder behandlas har
    som EN nod (en grupp): sokningen passerar nollbitarna fritt, stopp galler hela gruppen och
    ledningar raknas inte dubbelt. Bitarna identifieras pa sin punktlista (id), sa att tva bitar
    av samma ledning mellan samma noder (ringledning) halls isar."""

    def __init__(self, brunnar, tol, med_varden=None, mot_varden=None, sokradie=25.0):
        Natverk.__init__(self, brunnar, tol, sokradie)
        self.fram_ids = set()    # id(pts) for punktlistor lagrade i ritad riktning
        self.noll = set()        # id(pts) for bitar utan langd (P- och B-nod pa samma stalle)
        self.oid_noder = {}      # (lager, oid) -> noder som objektets bitar ror
        self.attr = {}           # (lager, oid) -> {'riktning', 'vg_upp', 'vg_ned', 'stopp'}
        self.med = _lista(med_varden) or set(['MED', 'JA', '1', 'TRUE', 'WITH', 'F'])
        self.mot = _lista(mot_varden) or set(['MOT', '-1', 'AGAINST', 'B'])
        self.stat = {'attribut': 0, 'vattengang': 0, 'ritad': 0, 'okant_attribut': 0}
        self._riktn = {}
        self._grupp = None       # nod -> kanonisk nod (lazy)
        self._medlemmar = None   # kanonisk nod -> [noder]

    def _kant(self, n1, n2, pts, lager, oid):
        bak = pts[::-1]
        self.kanter.setdefault(n1, []).append((n2, pts, lager, oid))
        self.kanter.setdefault(n2, []).append((n1, bak, lager, oid))
        self.fram_ids.add(id(pts))
        if _langd(pts) < 1e-6:
            self.noll.add(id(pts))
            self.noll.add(id(bak))
            self._grupp = None
        self.oid_noder.setdefault((lager, oid), set()).update((n1, n2))

    def inom_bbox(self, xmin, ymin, xmax, ymax):
        # Alla ledningar ska in i grafen (Natverk hoppar annars over ledningar langt fran
        # de sokta brunnarna - har ar alla brunnar sokta, men ledningar utan brunn ska med)
        return True

    # --- nodgrupper (noder forbundna med nollbitar) ---------------------------------
    def _bygg_grupper(self):
        foralder = {}

        def hitta(n):
            while foralder.get(n, n) != n:
                n = foralder[n]
            return n
        for nod, lista in self.kanter.items():
            for annan, pts, lager, oid in lista:
                if id(pts) in self.noll:
                    a, b = hitta(nod), hitta(annan)
                    if a != b:
                        foralder[b] = a
        grupp, medlemmar = {}, {}
        for nod in self.kanter:
            medlemmar.setdefault(hitta(nod), []).append(nod)
        for rot, noder in medlemmar.items():
            # Brunnsnoden ar kanonisk (kortast och forst i sortering: ('B', ...) < ('P', ...))
            kanon = sorted(noder, key=lambda n: (n[0] != 'B', txt(n[1])))[0]
            for n in noder:
                grupp[n] = kanon
        self._grupp, self._medlemmar = grupp, dict((grupp[r], n) for r, n in medlemmar.items())

    def kanon(self, nod):
        if self._grupp is None:
            self._bygg_grupper()
        return self._grupp.get(nod, nod)

    def medlemmar(self, nod):
        if self._grupp is None:
            self._bygg_grupper()
        return self._medlemmar.get(self.kanon(nod), [nod])

    def grannar(self, nod):
        """Riktiga kanter (inte nollbitar) fran nagon nod i gruppen:
        (annan nod, pts fran gruppen till annan, lager, oid)."""
        for m in self.medlemmar(nod):
            for annan, pts, lager, oid in self.kanter.get(m, ()):
                if id(pts) not in self.noll:
                    yield annan, pts, lager, oid

    # --- riktning -------------------------------------------------------------------
    def riktning(self, lager, oid):
        """('MED'|'MOT', kalla) - flodet i forhallande till ritad riktning."""
        n = (lager, oid)
        if n in self._riktn:
            return self._riktn[n]
        post = self.attr.get(n, {})
        ut = None
        v = post.get('riktning')
        if isinstance(v, float) and v == int(v):
            v = int(v)                          # 1.0 i ett Double-falt ska matcha '1'
        if v is not None and txt(v).strip():
            t = txt(v).strip().upper()
            if t in self.med:
                ut = ('MED', 'attribut')
            elif t in self.mot:
                ut = ('MOT', 'attribut')
            else:
                self.stat['okant_attribut'] += 1
        if ut is None:
            vu, vn = _tal(post.get('vg_upp')), _tal(post.get('vg_ned'))
            if vu is not None and vn is not None and vu != vn:
                ut = ('MED' if vu > vn else 'MOT', 'vattengang')
        if ut is None:
            ut = ('MED', 'ritad')
        self.stat[ut[1]] += 1
        self._riktn[n] = ut
        return ut

    def flodar_ut(self, pts, lager, oid):
        """True om flodet gar i punktlistans riktning (fran dess forsta till dess sista punkt)."""
        r, kalla_ = self.riktning(lager, oid)
        return (id(pts) in self.fram_ids) == (r == 'MED')

    def nedstroms_nod(self, lager, oid):
        """Nedstromsanden (kanonisk nod) for ett ledningsobjekt: den nod ingen av objektets
        bitar lamnar i flodesriktningen. Flera kandidater -> den forsta i sortering."""
        noder_fran, noder_till = set(), set()
        for nod in self.oid_noder.get((lager, oid), ()):
            for annan, pts, l, o in self.kanter.get(nod, ()):
                if (l, o) != (lager, oid) or id(pts) in self.noll:
                    continue
                if self.flodar_ut(pts, l, o):
                    noder_fran.add(self.kanon(nod)); noder_till.add(self.kanon(annan))
        slut = noder_till - noder_fran
        kand = slut or noder_till
        return sorted(kand, key=lambda n: (n[0] != 'B', txt(n[1])))[0] if kand else None

    def uppstroms(self, start, stopp_brunnar=None, max_kanter=None):
        """Allt uppstroms om noden start. Returnerar (kanter, noder, stoppade) dar
        kanter = [(nod_upp, nod_ned, pts upp->ned, lager, oid, avstand vid nod_ned, niva)],
        noder = {kanonisk nod: (avstand till start langs natet, niva)}, stoppade = antal
        kanter som inte passerats (stoppfalt) eller stoppbrunnar dar sokningen stannat.
        Stoppbrunnar laggs inte i noder och ledningen fran dem tas inte med."""
        stopp_brunnar = set(normalisera(b) for b in (stopp_brunnar or []))
        start = self.kanon(start)
        noder = {start: (0.0, 0)}
        ko = deque([start])
        kanter, stoppade, stoppade_noder = [], 0, set()

        def ar_stopp(nod):
            return any(m[0] == 'B' and m[1] in stopp_brunnar for m in self.medlemmar(nod))

        while ko:
            nod = ko.popleft()
            avst, niva = noder[nod]
            for annan, pts, lager, oid in self.grannar(nod):
                # pts gar fran gruppen (nod) till annan; kanten ar uppstroms om flodet gar mot pts
                if self.flodar_ut(pts, lager, oid):
                    continue
                if self.attr.get((lager, oid), {}).get('stopp'):
                    stoppade += 1
                    continue
                k_annan = self.kanon(annan)
                if k_annan != start and ar_stopp(k_annan):
                    if k_annan not in stoppade_noder:
                        stoppade_noder.add(k_annan)
                        stoppade += 1
                    continue
                kanter.append((k_annan, nod, pts[::-1], lager, oid, avst, niva))
                if k_annan not in noder:
                    noder[k_annan] = (avst + _langd(pts), niva + 1)
                    ko.append(k_annan)
                if max_kanter and len(kanter) >= max_kanter:
                    return kanter, noder, stoppade
        return kanter, noder, stoppade


class Segmentindex(object):
    """Rutnat over linjesegment for narhetstest (servisandar mot uppstromsledningar)."""

    def __init__(self, cell=25.0):
        self.cell = float(cell)
        self.rn = {}

    def lagg_till(self, pts, nyckel):
        c = self.cell
        for i in range(len(pts) - 1):
            a, b = pts[i], pts[i + 1]
            for cx in range(int(min(a[0], b[0]) // c), int(max(a[0], b[0]) // c) + 1):
                for cy in range(int(min(a[1], b[1]) // c), int(max(a[1], b[1]) // c) + 1):
                    self.rn.setdefault((cx, cy), []).append((a, b, nyckel))

    def narmast(self, x, y, tol):
        """(avstand, nyckel) for narmaste segment inom tol, annars (None, None)."""
        c = self.cell
        cx, cy = int(x // c), int(y // c)
        r = int(tol // c) + 1
        bast, nyckel = None, None
        for dx in range(-r, r + 1):
            for dy in range(-r, r + 1):
                for a, b, k in self.rn.get((cx + dx, cy + dy), ()):
                    d, t, q = _punkt_segment(x, y, a[0], a[1], b[0], b[1])
                    if d <= tol and (bast is None or d < bast):
                        bast, nyckel = d, k
        return bast, nyckel


# ------------------------------------------------ inlasning

def _sr_namn(d):
    try:
        return txt(d.spatialReference.name)
    except Exception:
        return ''


def _las_brunnar(brunnslager, brunn_id, omrade=None, bbox=None, sr=None):
    """{littera: (x, y)} ur brunnslagren (forsta forekomsten vinner). Brunnar i ett annat
    koordinatsystem an sr (ledningslagrets) projiceras."""
    brunnar = {}
    for lyr in brunnslager:
        idfalt = hitta_falt(lyr, brunn_id)
        src, dq = kalla(lyr)
        sr_namn = txt(getattr(sr, 'name', '') or '') if sr is not None else ''
        projicera = bool(sr_namn) and _sr_namn(arcpy.Describe(src)) not in ('', sr_namn)
        if projicera:
            logg('  %s projiceras till ledningslagrets koordinatsystem' % txt(getattr(lyr, 'name', lyr)))
        arcpy.MakeFeatureLayer_management(src, 'lyr_nat_br', dq)
        if omrade is not None:
            arcpy.SelectLayerByLocation_management('lyr_nat_br', 'INTERSECT', omrade, '', 'NEW_SELECTION')
        falt = ['SHAPE@' if projicera else 'SHAPE@XY', idfalt]
        with arcpy.da.SearchCursor('lyr_nat_br', falt) as mark:
            for geom, bid in mark:
                bid = normalisera(bid)
                if not bid or geom is None:
                    continue
                if projicera:
                    pnt = geom.projectAs(sr).firstPoint
                    x, y = pnt.X, pnt.Y
                else:
                    x, y = geom
                if x is None:
                    continue
                if bbox and not (bbox[0] <= x <= bbox[2] and bbox[1] <= y <= bbox[3]):
                    continue
                if bid not in brunnar:
                    brunnar[bid] = (x, y)
        arcpy.Delete_management('lyr_nat_br')
    return brunnar


def bygg_graf(ledningslager, brunnslager, brunn_id, tolerans=1.0, riktningsfalt=None,
              med_varden=None, mot_varden=None, vg_fran=None, vg_till=None,
              servis_falt=None, servis_varden=None, stopp_falt=None, stopp_varden=None,
              omradeslager=None, bbox=None):
    """Laser lagren och bygger grafen. Returnerar (graf, serviser) dar serviser =
    [(lager, oid, [punkter])] for ledningar som ar serviser (ingar inte i grafen)."""
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
        arcpy.MakeFeatureLayer_management(src, 'lyr_nat_omr', dq)
        omrade = 'lyr_nat_omr'

    sr = arcpy.Describe(kalla(led_lager[0])[0]).spatialReference
    logg('Laser brunnar')
    brunnar = _las_brunnar(brunn_lager, brunn_id, omrade, bbox, sr)
    logg('  %d brunnar' % len(brunnar))
    graf = Graf(brunnar, tolerans, med_varden, mot_varden)
    servis_v = _lista(servis_varden)
    stopp_v = _lista(stopp_varden)
    serviser = []
    n_led = n_bitar = n_servis = n_stopp = 0
    for lyr in led_lager:
        namn = txt(getattr(lyr, 'name', lyr))
        falt = dict((f.name.upper(), f.name) for f in arcpy.ListFields(kalla(lyr)[0]))
        extra = []
        for f, vad in ((riktningsfalt, 'riktningsfaltet'), (vg_fran, 'vattengangsfaltet (fran)'),
                       (vg_till, 'vattengangsfaltet (till)'), (servis_falt, 'servisfaltet'),
                       (stopp_falt, 'stoppfaltet')):
            namn_f = falt.get(txt(f).upper()) if f else None
            if f and not namn_f:
                logg('  VARNING: %s %s finns inte i %s - ignoreras for det lagret' % (vad, txt(f), namn))
            extra.append(namn_f)
        sedd = set()
        lasfalt = ['OID@', 'SHAPE@'] + [f for f in extra if f and not (f in sedd or sedd.add(f))]
        src, dq = kalla(lyr)
        arcpy.MakeFeatureLayer_management(src, 'lyr_nat_led', dq)
        if omrade is not None:
            arcpy.SelectLayerByLocation_management('lyr_nat_led', 'INTERSECT', omrade, '', 'NEW_SELECTION')
        logg('  %s: %s ledningar' % (namn, arcpy.GetCount_management('lyr_nat_led').getOutput(0)))
        with arcpy.da.SearchCursor('lyr_nat_led', lasfalt) as mark:
            for rad in mark:
                oid, geom = rad[0], rad[1]
                if geom is None:
                    continue
                if bbox:
                    try:
                        e = geom.extent
                        if e.XMax < bbox[0] or e.XMin > bbox[2] or e.YMax < bbox[1] or e.YMin > bbox[3]:
                            continue
                    except Exception:
                        pass
                varden = dict(zip([f for f in extra if f], rad[2:]))
                post = {'riktning': varden.get(extra[0]) if extra[0] else None,
                        'vg_upp': varden.get(extra[1]) if extra[1] else None,
                        'vg_ned': varden.get(extra[2]) if extra[2] else None,
                        'stopp': bool(extra[4] and txt(varden.get(extra[4])).strip().upper() in stopp_v)}
                ar_servis = bool(extra[3] and txt(varden.get(extra[3])).strip().upper() in servis_v)
                n_led += 1
                if post['stopp']:
                    n_stopp += 1
                for del_ in geom:
                    punkter = [(p.X, p.Y, None) for p in del_ if p is not None]
                    if len(punkter) < 2:
                        continue
                    if ar_servis:
                        serviser.append((namn, oid, punkter))
                        n_servis += 1
                        continue
                    graf.attr[(namn, oid)] = post
                    n_bitar += graf.lagg_till(punkter, namn, oid)
        arcpy.Delete_management('lyr_nat_led')
    if omrade is not None:
        arcpy.Delete_management('lyr_nat_omr')
    logg('  %d ledningar, %d bitar i grafen, %d serviser i ledningslagret, %d stoppledningar'
         % (n_led, n_bitar, n_servis, n_stopp))
    return graf, serviser


def _las_serviser(servislager, omradeslager=None, bbox=None):
    """[(lager, oid, [punkter])] ur ett eget servislager."""
    ut = []
    if not servislager:
        return ut
    if isinstance(servislager, (TEXTTYP, bytes)):
        servislager = [servislager]
    for n in servislager:
        lyr = hitta_lager(n)
        namn = txt(getattr(lyr, 'name', lyr))
        src, dq = kalla(lyr)
        arcpy.MakeFeatureLayer_management(src, 'lyr_nat_serv', dq)
        with arcpy.da.SearchCursor('lyr_nat_serv', ['OID@', 'SHAPE@']) as mark:
            for oid, geom in mark:
                if geom is None:
                    continue
                if bbox:
                    try:
                        e = geom.extent
                        if e.XMax < bbox[0] or e.XMin > bbox[2] or e.YMax < bbox[1] or e.YMin > bbox[3]:
                            continue
                    except Exception:
                        pass
                for del_ in geom:
                    punkter = [(p.X, p.Y, None) for p in del_ if p is not None]
                    if len(punkter) >= 2:
                        ut.append((namn, oid, punkter))
        arcpy.Delete_management('lyr_nat_serv')
    logg('  %d serviser i servislagret' % len(ut))
    return ut


def serviser_vid(kanter, serviser, tolerans):
    """Serviser vars ena ande ligger inom toleransen fran nagon av kanterna.
    Returnerar [(lager, oid, punkter, (lager, oid) for ledningen den ansluter till)]."""
    if not serviser or not kanter:
        return []
    idx = Segmentindex(max(25.0, tolerans * 5))
    for n_upp, n_ned, pts, lager, oid, avst, niva in kanter:
        idx.lagg_till(pts, (lager, oid))
    ut = []
    for lager, oid, punkter in serviser:
        for p in (punkter[0], punkter[-1]):
            d, nyckel = idx.narmast(p[0], p[1], tolerans)
            if nyckel is not None:
                ut.append((lager, oid, punkter, nyckel))
                break
    return ut


# ------------------------------------------------ enstaka startpunkt

def _startnod(graf, startbrunn, start_oid, ledningslager):
    if startbrunn:
        nod = ('B', normalisera(startbrunn))
        if nod not in graf.kanter:
            if nod not in graf.koord:
                raise RuntimeError('Brunnen %s finns inte i brunnslagren' % txt(startbrunn))
            raise RuntimeError('Brunnen %s ligger inte pa nagon ledning (tolerans?)' % txt(startbrunn))
        return nod, txt(startbrunn)
    if start_oid is not None:
        lager, oid = start_oid
        nod = graf.nedstroms_nod(lager, oid)
        if nod is None:
            raise RuntimeError('Den valda ledningen (OID %s) finns inte i grafen' % txt(oid))
        return nod, 'ledning %s' % txt(oid)
    raise RuntimeError('Ange en startbrunn eller markera en ledning i kartan')


def vald_ledning(ledningslager):
    """(lagernamn, OID) for den enda valda ledningen i nagot av lagren, annars None."""
    if isinstance(ledningslager, (TEXTTYP, bytes)):
        ledningslager = [ledningslager]
    for n in ledningslager:
        lyr = hitta_lager(n)
        try:
            valda = lyr.getSelectionSet()
        except Exception:
            valda = None
        if valda:
            if len(valda) > 1:
                raise RuntimeError('Markera bara en ledning (nu %d valda i %s)' % (len(valda), txt(lyr.name)))
            return txt(lyr.name), list(valda)[0]
    return None


def uppstroms(ledningslager, brunnslager, brunn_id, startbrunn=None, start_oid=None, ut_fc=None,
              csv_ut=None, riktningsfalt=None, med_varden=None, mot_varden=None,
              vg_fran=None, vg_till=None, servislager=None, servis_falt=None, servis_varden=None,
              stopp_falt=None, stopp_varden=None, stopp_brunnar=None, omradeslager=None,
              sokradie=None, tolerans=1.0, lagg_till_i_kartan=True):
    """Uppstromsanalys fran en startpunkt. Returnerar sammanfattningen (dict)."""
    bbox = None
    if sokradie and float(sokradie) > 0:
        # Begransa inlasningen till en ruta kring startbrunnen (hela natet ar stort)
        if startbrunn:
            if isinstance(brunnslager, (TEXTTYP, bytes)):
                brunnslager = [brunnslager]
            alla = _las_brunnar([hitta_lager(n) for n in brunnslager], brunn_id)
            p = alla.get(normalisera(startbrunn))
            if p is None:
                raise RuntimeError('Brunnen %s finns inte i brunnslagren' % txt(startbrunn))
            r = float(sokradie)
            bbox = (p[0] - r, p[1] - r, p[0] + r, p[1] + r)
        else:
            logg('  sokradien anvands bara med startbrunn - hela lagret lases')
    graf, serviser = bygg_graf(ledningslager, brunnslager, brunn_id, tolerans, riktningsfalt,
                               med_varden, mot_varden, vg_fran, vg_till, servis_falt, servis_varden,
                               stopp_falt, stopp_varden, omradeslager, bbox)
    serviser += _las_serviser(servislager, omradeslager, bbox)
    if start_oid is None and not startbrunn:
        start_oid = vald_ledning(ledningslager)
    start, startnamn = _startnod(graf, startbrunn, start_oid, ledningslager)

    logg('Soker uppstroms fran %s' % startnamn)
    kanter, noder, stoppade = graf.uppstroms(start, stopp_brunnar)
    serv = serviser_vid(kanter, serviser, tolerans)
    objekt = set((lager, oid) for n_upp, n_ned, pts, lager, oid, avst, niva in kanter)
    langd = sum(_langd(pts) for n_upp, n_ned, pts, lager, oid, avst, niva in kanter)
    brunnar_u = [n for n in noder if n[0] == 'B' and n != start]
    langst = max([a + _langd(pts) for n_upp, n_ned, pts, lager, oid, a, niva in kanter] or [0.0])
    sammanf = {
        'start': startnamn, 'antal_ledningar': len(objekt), 'antal_kanter': len(kanter),
        'langd_m': round(langd, 1), 'antal_brunnar': len(brunnar_u), 'antal_serviser': len(serv),
        'langsta_gren_m': round(langst, 1), 'stoppade': stoppade,
        'riktning_attribut': graf.stat['attribut'], 'riktning_vattengang': graf.stat['vattengang'],
        'riktning_ritad': graf.stat['ritad'],
    }
    logg('  %d ledningsobjekt (%d bitar), %.0f m, %d brunnar, %d serviser, langsta gren %.0f m'
         % (len(objekt), len(kanter), langd, len(brunnar_u), len(serv), langst))
    if stoppade:
        logg('  %d kanter/brunnar stoppade sokningen (stoppfalt/stopplista)' % stoppade)
    _logga_riktning(graf)

    if ut_fc:
        sr = arcpy.Describe(kalla(hitta_lager(ledningslager[0] if not isinstance(ledningslager, (TEXTTYP, bytes))
                                              else ledningslager))[0]).spatialReference
        fil = skriv_uppstroms_lager(ut_fc, sr, startnamn, kanter, serv, graf)
        sammanf['lager'] = fil
        if lagg_till_i_kartan:
            mxd = _mxd()
            if mxd is not None:
                try:
                    df = arcpy.mapping.ListDataFrames(mxd)[0]
                    l = arcpy.mapping.Layer(fil)
                    l.name = '%s %s' % (LAGERNAMN_UPPSTROMS, startnamn)
                    arcpy.mapping.AddLayer(df, l, 'TOP')
                    arcpy.RefreshTOC()
                    arcpy.RefreshActiveView()
                except Exception as e:
                    logg('  kunde inte lagga till lagret: %s' % txt(e))
    if csv_ut:
        with io.open(csv_ut, 'w', encoding='cp1252', errors='replace') as f:
            f.write('start;antal_ledningar;langd_m;antal_brunnar;antal_serviser;langsta_gren_m;stoppade\n')
            f.write(';'.join(txt(sammanf[k]).replace('.', ',') for k in
                             ('start', 'antal_ledningar', 'langd_m', 'antal_brunnar', 'antal_serviser',
                              'langsta_gren_m', 'stoppade')) + '\n')
            f.write('\nuppstroms_nod;nedstroms_nod;avstand_m;niva;langd_m;lager;oid\n')
            for n_upp, n_ned, pts, lager, oid, avst, niva in kanter:
                f.write(';'.join([_nodnamn(n_upp), _nodnamn(n_ned), ('%.1f' % avst).replace('.', ','),
                                  txt(niva), ('%.1f' % _langd(pts)).replace('.', ','), lager, txt(oid)]) + '\n')
        logg('  sammanfattning -> %s' % csv_ut)
    logg('KLART')
    return sammanf


def _logga_riktning(graf):
    s = graf.stat
    n = s['attribut'] + s['vattengang'] + s['ritad']
    if not n:
        return
    logg('  flodesriktning for de %d provade ledningarna: ur attribut %d, vattengang %d, '
         'ritad riktning (antagen) %d' % (n, s['attribut'], s['vattengang'], s['ritad']))
    if s['okant_attribut']:
        logg('  VARNING: %d ledningar har ett riktningsvarde som varken ar "med" eller "mot" - '
             'kontrollera vardena i dialogen' % s['okant_attribut'])
    if s['ritad'] > 0.5 * n:
        logg('  OBS: over halften av de provade ledningarna saknar riktningsuppgift - '
             'resultatet beror pa ritriktningen')


def skriv_uppstroms_lager(ut_fc, sr, startnamn, kanter, serv, graf):
    fil, ar_shp, max_text = _skapa_fc(ut_fc, 'POLYLINE', UPPSTROMS_FALT, sr)
    falt = ['SHAPE@'] + [n for n, t, l, a in UPPSTROMS_FALT]
    typer = ['SHAPE@'] + [t for n, t, l, a in UPPSTROMS_FALT]
    insert = arcpy.da.InsertCursor(fil, falt)
    try:
        for n_upp, n_ned, pts, lager, oid, avst, niva in kanter:
            arr = arcpy.Array()
            for x, y, z in pts:
                arr.add(arcpy.Point(x, y))
            insert.insertRow(_utan_null([
                arcpy.Polyline(arr, sr, False, False), startnamn[:50],
                _nodnamn(n_upp)[:50], _nodnamn(n_ned)[:50], round(avst, 1), niva,
                round(_langd(pts), 1), 'Nej', graf.riktning(lager, oid)[1], lager[:100], oid],
                typer, ar_shp))
        for lager, oid, punkter, ansl in serv:
            arr = arcpy.Array()
            for x, y, z in punkter:
                arr.add(arcpy.Point(x, y))
            insert.insertRow(_utan_null([
                arcpy.Polyline(arr, sr, False, False), startnamn[:50], '', '', None, None,
                round(_langd(punkter), 1), 'Ja', '', lager[:100], oid], typer, ar_shp))
    finally:
        del insert
    logg('  %d ledningar och %d serviser skrivna till %s' % (len(kanter), len(serv), fil))
    return fil


# ------------------------------------------------ batch per bedomd stracka

def uppstroms_batch(bedomda, ledningslager, brunnslager, brunn_id, csv_ut, riktningsfalt=None,
                    med_varden=None, mot_varden=None, vg_fran=None, vg_till=None, servislager=None,
                    servis_falt=None, servis_varden=None, stopp_falt=None, stopp_varden=None,
                    stopp_brunnar=None, omradeslager=None, tolerans=1.0, skriv_falt=True):
    """Uppstromsanalys for varje stracka i bedomda-lagret (fran "Skapa ledningslager"):
    startpunkt = strackans slutbrunn (nedstroms), sa att strackan sjalv ingar. Skriver
    ANT_SERV_U och L_UPPSTR till lagret (skriv_falt) och en CSV for tv3_analys.py."""
    graf, serviser = bygg_graf(ledningslager, brunnslager, brunn_id, tolerans, riktningsfalt,
                               med_varden, mot_varden, vg_fran, vg_till, servis_falt, servis_varden,
                               stopp_falt, stopp_varden, omradeslager)
    serviser += _las_serviser(servislager, omradeslager)
    har_serviser = bool(serviser)
    if not har_serviser:
        logg('  inga serviser angivna - antal anslutningar ur TV-inspektionen (ANT_ANSL) summeras'
             ' for de inspekterade strackorna uppstroms som skattning')

    lyr = hitta_lager(bedomda) if isinstance(bedomda, (TEXTTYP, bytes)) else bedomda
    src, dq = kalla(lyr)
    falt = dict((f.name.upper(), f.name) for f in arcpy.ListFields(src))
    for f in ('FRAN_BRUNN', 'TILL_BRUNN'):
        if f not in falt:
            raise RuntimeError('Lagret saknar faltet %s - valj lagret fran "Skapa ledningslager"' % f)
    lasfalt = ['OID@', falt['FRAN_BRUNN'], falt['TILL_BRUNN']]
    for f in ('NR', 'TV3_FIL', 'ANT_ANSL', 'MASK_BED'):
        lasfalt.append(falt.get(f))
    rader = []
    with arcpy.da.SearchCursor(src, [f for f in lasfalt if f], dq) as mark:
        for rad in mark:
            rad = list(rad)
            post = {}
            for f in lasfalt:
                post[f] = rad.pop(0) if f else None
            rader.append(post)
    logg('  %d strackor i %s' % (len(rader), txt(getattr(lyr, 'name', lyr))))

    # Anslutningar per slutbrunn (for skattningen): strackor vars nedstromsnod ligger uppstroms
    ansl_per_nod = {}
    for post in rader:
        nod = ('B', normalisera(post[falt['TILL_BRUNN']]))
        ansl_per_nod[nod] = ansl_per_nod.get(nod, 0) + int(_tal(post.get(falt.get('ANT_ANSL'))) or 0)

    resultat = {}          # oid -> (serviser, langd, antal ledningar, antal brunnar, kalla)
    cache = {}
    saknade = 0
    for post in rader:
        oid = post['OID@']
        till = normalisera(post[falt['TILL_BRUNN']])
        start = ('B', till)
        if start not in graf.kanter:
            saknade += 1
            continue
        if start in cache:
            resultat[oid] = cache[start]
            continue
        kanter, noder, stoppade = graf.uppstroms(start, stopp_brunnar)
        langd = sum(_langd(pts) for n_upp, n_ned, pts, lager, oid_, avst, niva in kanter)
        objekt = set((lager, oid_) for n_upp, n_ned, pts, lager, oid_, avst, niva in kanter)
        brunnar_u = sum(1 for n in noder if n[0] == 'B')      # inkl. startbrunnen
        if har_serviser:
            n_serv = len(serviser_vid(kanter, serviser, tolerans))
            kalla_ = 'servislager'
        else:
            n_serv = sum(ansl_per_nod.get(n, 0) for n in noder)
            kalla_ = 'skattning (ANT_ANSL)'
        cache[start] = (n_serv, round(langd, 1), len(objekt), brunnar_u, kalla_, stoppade)
        resultat[oid] = cache[start]
    if saknade:
        logg('  %d strackor vars slutbrunn inte ligger pa nagon ledning i grafen (hoppas over)' % saknade)
    _logga_riktning(graf)

    if skriv_falt:
        for namn, typ, alias in (('ANT_SERV_U', 'LONG', 'Serviser uppströms (inkl. sträckan)'),
                                 ('L_UPPSTR', 'DOUBLE', 'Längd uppströms (m, inkl. sträckan)')):
            if namn not in falt:
                arcpy.AddField_management(src, namn, typ, field_alias=alias)
                falt[namn] = namn
        n = 0
        with arcpy.da.UpdateCursor(src, ['OID@', falt['ANT_SERV_U'], falt['L_UPPSTR']], dq) as mark:
            for rad in mark:
                r = resultat.get(rad[0])
                if r is None:
                    continue
                rad[1], rad[2] = r[0], r[1]
                mark.updateRow(rad)
                n += 1
        logg('  ANT_SERV_U och L_UPPSTR skrivna for %d strackor' % n)

    with io.open(csv_ut, 'w', encoding='cp1252', errors='replace') as f:
        f.write('fil;nr;startbrunn;slutbrunn;serviser_uppstroms;langd_uppstroms_m;'
                'antal_ledningar;antal_brunnar;serviser_kalla;stoppade\n')
        for post in rader:
            r = resultat.get(post['OID@'])
            if r is None:
                continue
            f.write(';'.join([txt(post.get(falt.get('TV3_FIL')) or ''), txt(post.get(falt.get('NR')) or ''),
                              txt(post[falt['FRAN_BRUNN']]), txt(post[falt['TILL_BRUNN']]),
                              txt(r[0]), txt(r[1]).replace('.', ','), txt(r[2]), txt(r[3]), r[4], txt(r[5])])
                    + '\n')
    logg('  %d strackor -> %s (uppstroms: %s i listfilen till tv3_analys.py)' % (len(resultat), csv_ut, csv_ut))
    logg('KLART')
    return resultat
