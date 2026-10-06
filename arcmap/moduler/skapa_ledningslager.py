# -*- coding: utf-8 -*-
"""
Skapar ett ledningslager for bedomda stracker ur tv3_analys.

Tva satt att kora:

  1. Som verktyg i ArcMap: lagg till arcmap/tv3_verktyg.pyt i ArcToolbox och kor
     "Skapa ledningslager" - da valjer du JSON-fil, lager och utdata i en dialog.

  2. I ArcMaps Python-fonster med kartan oppen, med installningarna i KONFIG nedan:
         execfile(r'H:\PY\tv3analys\arcmap\moduler\skapa_ledningslager.py')

Indata ar kartunderlag.json fran tv3_analys.py. Ledningslagret byggs om till en graf
(Natverk): varje ledningsdel delas dar en brunn ur JSON-filen ligger inom TOLERANS fran
linjen, fria ledningsandar blir noder och andar inom toleransen slas ihop. For varje
stracka (brunnspar) soks vagen med farst bitar mellan brunnarna, med hogst MAX_HOPP - 1
andra sokta brunnar emellan och hogst MAX_DELAR bitar, och den klipps ut som ett eget
objekt. Det hanterar bade ledningar som passerar flera brunnar utan att vara fysiskt
uppdelade och strackor som bestar av flera ledningsobjekt.

Varje objekt far tva bedomningsfalt:
    MASK_BED  "Maskinell bedomning" - prioritetsklass A-E fran poangmodellen
    MAN_BED   "Manuell bedomning"   - tomt, fylls i for hand i ArcMap

samt de harledda falten BEDOMNING (manuell om ifylld, annars maskinell),
BED_TYP (Maskinell/Manuell) och STIL ("A - Maskinell"), som symbologin utgar
fran: farg efter klass, streckad linje for maskinell och heldragen for manuell.

Efter att du fyllt i manuella bedomningar: kor verktyget "Uppdatera bedomning"
(eller detta skript med BARA_UPPDATERA = True), sa raknas BEDOMNING, BED_TYP
och STIL om utan att geometrin byggs om. Manuella bedomningar bevaras aven vid
en full omkorning.
"""
from __future__ import unicode_literals

import os
import re
import io
import sys
import json
import arcpy

# =====================================================================
# KONFIG - galler bara vid korning med execfile (verktyget har egen dialog)
# =====================================================================

# Del av lagernamnet i innehallsforteckningen. Flera lager tillatna.
LEDNINGSLAGER = ['A Ledning']
# Nedstigningsbrunnar (xNB/xNBL) och rens-/tillsynsbrunnar (xRB/xTB) ligger i olika lager -
# ta med alla lager dar brunnar i TV3-filerna kan finnas.
BRUNNSLAGER   = ['A Nedstign och övriga brunnar', 'A Rensbrunn/tillsynsbrunn',
                 'A Platsgjuten brunnspunkt']

BRUNN_ID = 'EntityID'        # faltet med brunnsbeteckning i brunnslagren

# Valfritt polygonlager att begransa sokningen till. None = hela lagren.
OMRADESLAGER = None          # t.ex. 'paverkansomrade_grovt'

JSON_IN = r'H:\PY\tv3analys\tv3_resultat\kartunderlag.json'
UT_FC   = r'H:\PY\tv3analys\Karta\bedomda_ledningar'       # .gdb-vag eller mapp (= shapefil)
CSV_UT  = r'H:\PY\tv3analys\Karta\omatchade_par.csv'
# Symbologi. Satts en gang i ArcMap och sparas som .lyr i arcmap-mappen (handledningen 7.2).
ARCMAP_MAPP = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))   # .lyr-filerna ligger i arcmap-mappen
LYR_FIL = os.path.join(ARCMAP_MAPP, 'bedomda_ledningar.lyr')

TOLERANS = 2.0     # meter mellan ledningens vertex och brunnen
MARGINAL = 100.0   # meter utanfor omradet dar brunnar anda las in (om OMRADESLAGER anges)
MAX_HOPP = 2       # brunnar en filmad stracka far passera, inkl. slutbrunnen (2 = en mellanbrunn)
MAX_DELAR = 8      # hogsta antal ledningsbitar en stracka far besta av
GRENROR_RADIE_EXTRA = 1.0   # m utover toleransen: sa langt fran kartunderlagets koordinat far den fria anden ligga
# Parallella ledningar (spill och dag i samma schakt, braddbrunnar med tva kammare): nar flera vagar
# ar lika korta valjs den vars ledningstyp (forsta bokstaven S/D/K) och dimension stammer med filmen.
TYP_FALT = None            # falt i ledningslagret med ledningstyp, t.ex. 'PipeType' (DSL/SSL/KSL)
DIM_FALT = None            # falt med dimension, t.ex. 'PipeDimension'
DIM_TOLERANS = 0.1         # dimensioner inom 10 % raknas som samma (225 i filmen, 230 i GIS)
POANG_TYP = 2              # avvikelsepoang per ledning med annan typ an filmen
POANG_DIM = 1              # ... med annan dimension
EXTRA_DELAR = 1            # en vag med ratt typ far ha sa manga fler bitar an den kortaste
KLIPP_AVBRUTNA = True      # avbruten/ofullstandig film: rita bara den filmade delen fran kamerans brunn
URVAL    = 'INTERSECT'     # 'WITHIN' om ledningen maste ligga helt inom omradet

# Falt fran ledningslagret som ska folja med till resultatet.
KOPIERA_FALT = []          # t.ex. ['DIMENSION', 'MATERIAL', 'ANLAGGNINGSAR']

# Var PDF-rapporterna och filmerna ligger SETT FRAN DEN HAR DATORN (t.ex. Citrix), for
# hyperlankfalten RAPPORT och VIDEO. None = rapporter i mappen "rapporter" bredvid JSON-filen,
# filmer pa sokvagen som tv3_analys hittade dem (bara ratt om det ar samma dator).
RAPPORTMAPP = None         # t.ex. r'\\server\share\tv3_resultat\rapporter'
FILMMAPP = None            # t.ex. r'\\server\share\Inspektioner\Filmer'

# Valfri GeoJSON-export (WGS84, 2D) for webb-GIS. None = ingen.
GEOJSON_UT = None          # t.ex. r'H:\PY\tv3analys\Karta\bedomda_ledningar.geojson'

# Valfria lager med svackor (punkter dar stående vatten ar djupast) och bakfall (linjer
# dar ledningen lutar mot flodesriktningen), ur inklinometerprofilerna. None = skrivs inte.
# Positionerna i JSON-filen ar kamerans; de laggs ut langs kartlinjen fran utgangsbrunnen,
# skalade med kartlangd/filmlangd (ingen skalning vid avbruten inspektion).
SVACKOR_UT = None          # t.ex. r'H:\PY\tv3analys\Karta\svackor'
BAKFALL_UT = None          # t.ex. r'H:\PY\tv3analys\Karta\bakfall'
SVACKA_MIN_CM = 2          # svackor grundare an sa tas inte med i svacklagret

# True = bygg inte om geometrin, rakna bara om BEDOMNING/BED_TYP/STIL i UT_FC
# efter att manuella bedomningar fyllts i.
BARA_UPPDATERA = False

# =====================================================================

arcpy.env.overwriteOutput = True

KLASSER = ['A', 'B', 'C', 'D', 'E']
KLASSORDNING = dict((k, i) for i, k in enumerate(KLASSER))   # A = varst
LAGERNAMN = 'Bedomda ledningar'

SYMBOLOGI_TIPS = [
    'Ingen .lyr-fil an. Satt symbologi och hyperlankar en gang:',
    '  Egenskaper > Symbology > Categories > Unique values, Value Field: STIL',
    '  Add All Values ger tio kategorier:',
    '    A/B/C/D/E - Maskinell  ->  STRECKAD linje i klassens farg',
    '    A/B/C/D/E - Manuell    ->  HELDRAGEN linje i klassens farg',
    '  Fargar: A rott, B orange, C gult, D gront, E gratt.',
    '  Egenskaper > Display > Support Hyperlinks using field: RAPPORT (Document)',
    '  -> Hyperlank-verktyget (blixten) oppnar PDF-rapporten nar du klickar pa ledningen.',
    '  Spara sedan lagret som .lyr - nasta korning applicerar allt automatiskt.',
]


# ------------------------------------------------ unicode-hjalp (Python 2)

try:
    TEXTTYP = unicode          # Python 2
except NameError:
    TEXTTYP = str              # Python 3


def txt(v):
    """Gor vad som helst till unicode utan att krascha."""
    if v is None:
        return ''
    if isinstance(v, TEXTTYP):
        return v
    if isinstance(v, bytes):
        for enc in ('utf-8', 'cp1252'):
            try:
                return v.decode(enc)
            except UnicodeDecodeError:
                continue
        return v.decode('latin-1', 'replace')
    try:
        return TEXTTYP(v)
    except UnicodeDecodeError:          # t.ex. IOError med cp1252-sokvag i Python 2
        return txt(str(v))


def normalisera(littera):
    """Gor brunnsbeteckningar jamforbara mellan TV3-filer och databasen."""
    if littera is None:
        return None
    return re.sub(r'[\s\-_]', '', txt(littera)).upper()


def logg(*args):
    """Skriver bade till geoprocessing-meddelanden (verktyg) och Python-fonstret."""
    rad = ' '.join(txt(a) for a in args)
    try:
        arcpy.AddMessage(rad)
    except Exception:
        pass
    try:
        print(rad)
        sys.stdout.flush()
    except (UnicodeEncodeError, IOError, AttributeError):
        pass                             # stdout utan teckenkodning eller flush (ArcMap-dialog)


# ------------------------------------------------ lager och falt

def _mxd():
    try:
        return arcpy.mapping.MapDocument('CURRENT')
    except Exception:
        return None


def hitta_lager(namn):
    """Hamtar ett lager. Sokvag eller lagernamn som arcpy kanner igen anvands som
    det ar; annars soks kartans innehallsforteckning - exakt namn i forsta hand,
    annars delstrang. Kastar fel vid tvetydighet istallet for att gissa."""
    if not isinstance(namn, (TEXTTYP, bytes)):
        return namn                       # redan ett lagerobjekt
    n = txt(namn)
    sokt = n.strip().lower()

    mxd = _mxd()
    if mxd is None:
        if arcpy.Exists(n):
            return n
        raise RuntimeError('Hittar inte "%s" och ingen karta ar oppen' % n)
    alla = [l for l in arcpy.mapping.ListLayers(mxd) if l.isFeatureLayer]

    def langt(l):
        return txt(getattr(l, 'longName', '') or '').strip().lower()

    # Exakt pa namn eller langt namn (Grupp\Lager) - lagerobjektet returneras, inte namnet,
    # eftersom namn med '/' inte gar att skicka som text till geoprocessing-verktyg
    exakta = [l for l in alla if txt(l.name).strip().lower() == sokt or langt(l) == sokt]
    if not exakta and (os.path.sep in n or n.lower().endswith('.shp')) and arcpy.Exists(n):
        return n                          # sokvag till en featureklass
    if len(exakta) == 1:
        logg('  "%s" -> %s' % (n, txt(exakta[0].name)))
        return exakta[0]
    if len(exakta) > 1:
        raise RuntimeError('Flera lager heter exakt "%s"' % n)

    delvis = [l for l in alla if sokt in txt(l.name).strip().lower()]
    if len(delvis) == 1:
        logg('  "%s" -> %s' % (n, txt(delvis[0].name)))
        return delvis[0]
    if len(delvis) > 1:
        raise RuntimeError('Flera lager matchar "%s": %s\nAnge exakt namn.'
                           % (n, ', '.join(txt(l.name) for l in delvis)))
    raise RuntimeError('Hittade inget lager som matchar "%s"' % n)


def kalla(lyr):
    """(datakalla, definitionsfraga) att ge MakeFeatureLayer. Ett lagerobjekt oversatts
    till sin datakalla (sokvag till featureklassen) sa att lagernamn med '/' eller
    grupplager inte tolkas som sokvagar av geoprocessing-verktygen."""
    ds = getattr(lyr, 'dataSource', None)
    if ds:
        try:
            dq = lyr.definitionQuery
        except Exception:
            dq = ''
        return ds, (dq or None)
    return lyr, None


def hitta_falt(lyr, faltnamn):
    """Returnerar faltets riktiga namn, oberoende av versaler."""
    falt = arcpy.ListFields(kalla(lyr)[0])
    for f in falt:
        if f.name.upper() == txt(faltnamn).upper():
            return f.name
    textfalt = [f.name for f in falt if f.type == 'String']
    raise RuntimeError(
        'Faltet "%s" finns inte i lagret "%s".\nTextfalt som finns: %s'
        % (txt(faltnamn), txt(getattr(lyr, 'name', lyr)), ', '.join(textfalt)))


# ------------------------------------------------ geometri

def _avst(p, q):
    return ((p[0] - q[0]) ** 2 + (p[1] - q[1]) ** 2) ** 0.5


def _punkt_segment(px, py, ax, ay, bx, by):
    """Avstand fran punkt till segment a-b. Returnerar (avstand, t, (qx, qy)) dar
    t i [0, 1] ar laget langs segmentet och q ar narmaste punkt pa det."""
    dx, dy = bx - ax, by - ay
    l2 = dx * dx + dy * dy
    if l2 == 0:
        t = 0.0
    else:
        t = ((px - ax) * dx + (py - ay) * dy) / l2
        t = 0.0 if t < 0 else (1.0 if t > 1 else t)
    qx, qy = ax + t * dx, ay + t * dy
    return ((px - qx) ** 2 + (py - qy) ** 2) ** 0.5, t, (qx, qy)


class Natverk(object):
    """Graf av ledningsbitar mellan brunnar och fria ledningsandar.

    Bara de brunnar som forekommer i JSON-filen ar noder. Varje ledningsdel delas
    vid de sokta brunnar som ligger inom toleransen fran linjen - i en vertex eller
    mitt pa ett segment - och bitarna blir kanter. Ledningsandar utan brunn blir
    ocksa noder, och tva andar inom toleransen blir samma nod. Darmed hittas aven
    strackor som ar uppdelade i flera ledningsobjekt (t.ex. vid materialbyte) och
    strackor dar brunnen ligger pa linjen utan egen vertex."""

    def __init__(self, sokta, tol, sokradie=25.0):
        self.tol = float(tol)
        if not self.tol > 0:
            raise RuntimeError('Toleransen maste vara storre an 0 m')
        self.sokradie = max(float(sokradie), self.tol)
        self.cell = self.sokradie
        self.rutnat = {}
        for bid, (x, y) in sokta.items():
            self.rutnat.setdefault((int(x // self.cell), int(y // self.cell)), []).append((bid, x, y))
        self.narmast = {}      # brunn -> minsta avstand till nagon ledning (diagnostik)
        self.kanter = {}       # nod -> [(annan nod, punkter fran nod till annan, lager, oid)]
        self.andar = {}        # rutnat over fria andar: cell -> [(nyckel, x, y)]
        self.n_andar = 0
        self.koord = dict((('B', bid), (x, y)) for bid, (x, y) in sokta.items())
        if sokta:
            xs = [p[0] for p in sokta.values()]
            ys = [p[1] for p in sokta.values()]
            m = self.sokradie
            self.bbox = (min(xs) - m, min(ys) - m, max(xs) + m, max(ys) + m)
        else:
            self.bbox = None

    def inom_bbox(self, xmin, ymin, xmax, ymax):
        """Snabbtest: kan ledningen alls beror nagon sokt brunn?"""
        if self.bbox is None:
            return False
        return not (xmax < self.bbox[0] or xmin > self.bbox[2]
                    or ymax < self.bbox[1] or ymin > self.bbox[3])

    def _brunnar_nara(self, ax, ay, bx, by):
        r, c = self.sokradie, self.cell
        x0, x1 = min(ax, bx) - r, max(ax, bx) + r
        y0, y1 = min(ay, by) - r, max(ay, by) + r
        for cx in range(int(x0 // c), int(x1 // c) + 1):
            for cy in range(int(y0 // c), int(y1 // c) + 1):
                for b in self.rutnat.get((cx, cy), ()):
                    yield b

    def _andnod(self, x, y):
        """Nod for en fri ledningsande; andar inom toleransen delar nod."""
        c = self.tol
        cx, cy = int(x // c), int(y // c)
        for dx in (-1, 0, 1):
            for dy in (-1, 0, 1):
                for nyckel, ex, ey in self.andar.get((cx + dx, cy + dy), ()):
                    if (ex - x) ** 2 + (ey - y) ** 2 <= c * c:
                        return nyckel
        self.n_andar += 1
        nyckel = ('P', self.n_andar)
        self.andar.setdefault((cx, cy), []).append((nyckel, x, y))
        self.koord[nyckel] = (x, y)
        return nyckel

    def _kant(self, n1, n2, pts, lager, oid):
        self.kanter.setdefault(n1, []).append((n2, pts, lager, oid))
        self.kanter.setdefault(n2, []).append((n1, pts[::-1], lager, oid))

    def lagg_till(self, punkter, lager, oid):
        """Lagger in en ledningsdel [(x, y, z)] och delar den vid sokta brunnar.
        Returnerar antal bitar."""
        if len(punkter) < 2:
            return 0
        matt = [0.0]
        for i in range(1, len(punkter)):
            matt.append(matt[-1] + _avst(punkter[i - 1], punkter[i]))

        traffar = {}   # brunn -> (matt, avstand, punkt)
        for i in range(len(punkter) - 1):
            a, b = punkter[i], punkter[i + 1]
            seglen = matt[i + 1] - matt[i]
            for bid, bx, by in self._brunnar_nara(a[0], a[1], b[0], b[1]):
                d, t, q = _punkt_segment(bx, by, a[0], a[1], b[0], b[1])
                if d < self.narmast.get(bid, 1e30):
                    self.narmast[bid] = d
                if d <= self.tol and (bid not in traffar or d < traffar[bid][1]):
                    z = None
                    if a[2] is not None and b[2] is not None:
                        z = a[2] + t * (b[2] - a[2])
                    traffar[bid] = (matt[i] + t * seglen, d, (q[0], q[1], z))

        noder = [(0.0, self._andnod(punkter[0][0], punkter[0][1]), punkter[0])]
        for bid, (m, d, q) in traffar.items():
            noder.append((m, ('B', bid), q))
        noder.append((matt[-1], self._andnod(punkter[-1][0], punkter[-1][1]), punkter[-1]))
        noder.sort(key=lambda n: n[0])

        n_bitar = 0
        for k in range(len(noder) - 1):
            m1, n1, p1 = noder[k]
            m2, n2, p2 = noder[k + 1]
            if n1 == n2:
                continue
            pts = [p1] + [punkter[j] for j in range(len(punkter)) if m1 < matt[j] < m2] + [p2]
            self._kant(n1, n2, pts, lager, oid)
            n_bitar += 1
        return n_bitar

    def vag(self, a, b, max_hopp, max_delar=8, poang=None):
        """Vagen med farst bitar fran brunn a till brunn b, eller None. Hogst
        max_hopp - 1 andra sokta brunnar far passeras, och hogst max_delar bitar.
        poang: se _vag - val mellan parallella ledningar efter typ och dimension."""
        return self._vag(('B', a), ('B', b), max_hopp, max_delar, poang)

    def fri_ande_nara(self, x, y, radie):
        """Noden for den fria ledningsande som ligger narmast (x, y) inom radie, annars None."""
        c = self.tol
        n = int(radie // c) + 1
        cx, cy = int(x // c), int(y // c)
        basta = None
        for dx in range(-n, n + 1):
            for dy in range(-n, n + 1):
                for nyckel, ex, ey in self.andar.get((cx + dx, cy + dy), ()):
                    d2 = (ex - x) ** 2 + (ey - y) ** 2
                    if d2 <= radie * radie and (basta is None or d2 < basta[0]):
                        basta = (d2, nyckel)
        return basta[1] if basta else None

    def vag_till_punkt(self, a, x, y, max_hopp, max_delar=8, radie=None, poang=None):
        """Vagen fran brunn a till den fria ledningsande som ligger vid (x, y) - en ledning som
        slutar i ett grenror/pastick pa en annan ledning i stallet for i en brunn. Returnerar
        None om ingen fri ande finns inom radie (standard tolerans + GRENROR_RADIE_EXTRA) eller ingen vag."""
        mal = self.fri_ande_nara(x, y, radie if radie is not None else self.tol + GRENROR_RADIE_EXTRA)
        if mal is None:
            return None
        return self._vag(('B', a), mal, max_hopp, max_delar, poang)

    def _vag(self, start, mal, max_hopp, max_delar=8, poang=None):
        """Vagen med farst bitar. Med poang (funktion (lager, oid) -> avvikelsepoang mot filmen,
        t.ex. annan ledningstyp eller dimension) jamfors alla vagar med hogst EXTRA_DELAR fler
        bitar an den kortaste, och den med lagst poang tas - sa att en spillvattenfilm hamnar
        pa spilledningen och inte pa dagvattenledningen bredvid i samma schakt."""
        forsta = self._kortaste(start, mal, max_hopp, max_delar)
        if forsta is None or poang is None:
            return forsta
        if self._poang(forsta, poang) == 0:
            return forsta
        return self._basta(start, mal, max_hopp, min(max_delar, len(forsta) + EXTRA_DELAR), poang, forsta)

    @staticmethod
    def _poang(vagen, poang):
        return sum(poang(lager, oid) for lager, oid in set((l, o) for n, p, l, o in vagen))

    def _basta(self, start, mal, max_hopp, max_delar, poang, forsta, tak=2000):
        """Alla enkla vagar start -> mal med hogst max_delar bitar och max_hopp - 1 mellanbrunnar
        (djupet forst, hogst tak vagar), rankade pa (poang, antal bitar, langd)."""
        basta = (self._poang(forsta, poang), len(forsta), _langd(sla_ihop(forsta)), forsta)
        n = [0]

        def gren(nod, vagen, hopp, sedda):
            if n[0] >= tak or len(vagen) >= max_delar:
                return
            for annan, pts, lager, oid in self.kanter.get(nod, ()):
                if annan == mal:
                    n[0] += 1
                    v = vagen + [(annan, pts, lager, oid)]
                    nyckel = (self._poang(v, poang), len(v), _langd(sla_ihop(v)), v)
                    if nyckel[:3] < tuple(basta[:3]):
                        basta[:] = nyckel
                    continue
                if annan in sedda:
                    continue
                h = hopp + (1 if annan[0] == 'B' else 0)
                if h > max_hopp - 1:
                    continue
                sedda.add(annan)
                gren(annan, vagen + [(annan, pts, lager, oid)], h, sedda)
                sedda.discard(annan)

        basta = list(basta)
        gren(start, [], 0, set([start]))
        return basta[3]

    def _kortaste(self, start, mal, max_hopp, max_delar=8):
        from collections import deque
        if start not in self.kanter or mal not in self.kanter:
            return None
        ko = deque([(start, [], 0)])
        # Farsta antal mellanbrunnar som noden natts med. En nod far besokas igen om den
        # nas med farre mellanbrunnar - annars kan en kortare vag via en brunn blockera
        # den enda tillatna vagen runt en slinga i natet.
        basta_hopp = {start: 0}
        while ko:
            nod, vagen, hopp = ko.popleft()
            if len(vagen) >= max_delar:
                continue
            for annan, pts, lager, oid in self.kanter.get(nod, ()):
                if annan == mal:
                    return vagen + [(annan, pts, lager, oid)]
                h = hopp
                if annan[0] == 'B':
                    h += 1
                    if h > max_hopp - 1:
                        continue
                if h >= basta_hopp.get(annan, 999):
                    continue
                basta_hopp[annan] = h
                ko.append((annan, vagen + [(annan, pts, lager, oid)], h))
        return None


    def komponent(self, start):
        """Alla noder som gar att na fran start."""
        sedda = set([start])
        ko = [start]
        while ko:
            nod = ko.pop()
            for annan, pts, lager, oid in self.kanter.get(nod, ()):
                if annan not in sedda:
                    sedda.add(annan)
                    ko.append(annan)
        return sedda

    def diagnos(self, a, b, max_hopp, max_delar=8):
        """Varfor hittades ingen vag mellan a och b? Returnerar en forklaring."""
        start, mal = ('B', a), ('B', b)
        if start not in self.kanter or mal not in self.kanter:
            return 'brunnen ligger inte pa nagon ledning'
        v = self.vag(a, b, 999, 200)
        if v:
            n_br = sum(1 for nod, pts, lager, oid in v[:-1] if nod[0] == 'B')
            if n_br + 1 > max_hopp:
                return ('vag finns via %d bitar och %d andra brunnar - hoj max hopp till %d'
                        % (len(v), n_br, n_br + 1))
            return ('vag finns men via %d bitar (max %d) - hoj max delar'
                    % (len(v), max_delar))
        ka = self.komponent(start)
        kb = self.komponent(mal)
        if ka == kb:
            return 'samma natverk men vagen ar orimligt lang'
        # Narmaste avstand mellan de tva natverksdelarna (rutnat over den storre delen,
        # sokning fran den mindre)
        if len(kb) < len(ka):
            ka, kb = kb, ka
        c = self.cell
        rn = {}
        for nod in kb:
            x, y = self.koord[nod]
            rn.setdefault((int(x // c), int(y // c)), []).append((x, y))
        bast, var = None, None
        for nod in ka:
            x, y = self.koord[nod]
            cx, cy = int(x // c), int(y // c)
            for dx in (-1, 0, 1):
                for dy in (-1, 0, 1):
                    for bx, by in rn.get((cx + dx, cy + dy), ()):
                        d = ((bx - x) ** 2 + (by - y) ** 2) ** 0.5
                        if bast is None or d < bast:
                            bast, var = d, ((x + bx) / 2.0, (y + by) / 2.0)
        if bast is None:
            return ('ledningen ar bruten med mer an %.0f m glapp, eller ligger i ett annat lager'
                    % self.sokradie)
        return ('glapp %.1f m i ledningen vid (%.0f, %.0f) - hoj toleransen till %.0f m'
                ' eller kontrollera ledningen dar' % (bast, var[0], var[1], bast + 0.5))


def sla_ihop(vagen):
    """Punktlista for en vag av kanter, utan dubbla skarvpunkter."""
    pts = []
    for annan, p, lager, oid in vagen:
        for q in p:
            if pts and abs(q[0] - pts[-1][0]) < 1e-6 and abs(q[1] - pts[-1][1]) < 1e-6:
                continue
            pts.append(q)
    return pts


def _langd(pts):
    return sum(_avst(pts[i], pts[i + 1]) for i in range(len(pts) - 1))


def _typbokstav(v):
    """Forsta bokstaven i en ledningstyp: 'Spillvatten'/'SSL' -> 'S', 'DSL' -> 'D', 'Kombinerat' -> 'K'."""
    t = (txt(v) or '').strip().upper()
    return t[:1] if t and t[:1].isalpha() else ''


def _dim_mm(v):
    """Forsta talet i en dimensionstext: '600/900' -> 600, 'Ø 225' -> 225, 'ODEF' -> None."""
    m = re.search(r'\d+', txt(v) or '')
    return int(m.group(0)) if m else None


def _avvikelse(post, typ, dim):
    """(poang, text) for en GIS-ledning med ledningstyp typ och dimension dim mot filmens uppgifter.
    Okanda varden pa nagon sida ger ingen avvikelse."""
    p, delar = 0, []
    ft, gt = _typbokstav(post.get('ledningstyp')), _typbokstav(typ)
    if ft and gt and ft != gt:
        p += POANG_TYP
        delar.append('typ %s i filmen, %s i GIS' % (ft, txt(typ).strip()))
    fd, gd = _dim_mm(post.get('dimension')), _dim_mm(dim)
    if fd and gd and abs(fd - gd) > DIM_TOLERANS * max(fd, gd):
        p += POANG_DIM
        delar.append('dimension %d i filmen, %d i GIS' % (fd, gd))
    return p, '; '.join(delar)


def _klipp_pts(pts, m0, m1):
    """Delen av punktlistan mellan matten m0 och m1 (m fran forsta punkten), med interpolerade
    andpunkter. Utanfor linjen kapas matten till [0, langd]."""
    L = _langd(pts)
    m0, m1 = max(0.0, m0), min(L, m1)
    if m1 - m0 <= 1e-9:
        return pts[:1]

    def punkt_vid(m):
        matt = 0.0
        for i in range(len(pts) - 1):
            a, b = pts[i], pts[i + 1]
            seg = _avst(a, b)
            if seg > 0 and matt + seg >= m - 1e-9:
                t = min(1.0, max(0.0, (m - matt) / seg))
                z = (a[2] + t * (b[2] - a[2])) if a[2] is not None and b[2] is not None else None
                return (a[0] + t * (b[0] - a[0]), a[1] + t * (b[1] - a[1]), z)
            matt += seg
        return pts[-1]

    ut, matt = [punkt_vid(m0)], 0.0
    for i in range(len(pts) - 1):
        matt += _avst(pts[i], pts[i + 1])
        if m0 < matt < m1:
            ut.append(pts[i + 1])
    ut.append(punkt_vid(m1))
    return ut


def _punkt_vid(pts, d):
    """Punkt (x, y) pa avstandet d langs punktlistan (klamms till andarna)."""
    if d <= 0:
        return pts[0][0], pts[0][1]
    g = 0.0
    for i in range(len(pts) - 1):
        seg = _avst(pts[i], pts[i + 1])
        if seg > 0 and g + seg >= d:
            t = (d - g) / seg
            return (pts[i][0] + t * (pts[i + 1][0] - pts[i][0]),
                    pts[i][1] + t * (pts[i + 1][1] - pts[i][1]))
        g += seg
    return pts[-1][0], pts[-1][1]


def _delstracka(pts, d0, d1):
    """Punktlista for delen mellan avstanden d0 och d1 langs linjen (d0 < d1)."""
    ut = [_punkt_vid(pts, d0)]
    g = 0.0
    for i in range(len(pts) - 1):
        seg = _avst(pts[i], pts[i + 1])
        if d0 < g + seg < d1 and g + seg > d0:
            ut.append((pts[i + 1][0], pts[i + 1][1]))
        g += seg
    ut.append(_punkt_vid(pts, d1))
    return ut


def _kartposition(post, pos, l_karta, a, b):
    """Kamerans position pos (m fran utgangsbrunnen) -> avstand fran a langs kartlinjen a-b.
    Skalas med kartlangd/filmlangd, utom vid avbruten inspektion (da ar filmpositionen
    kartmeter fran kamerans brunn, samma regel som markprofilen)."""
    langd_film = _tal(post.get('langd_m'))
    skala = 1.0
    if not (post.get('avbruten') or post.get('ofullstandig')) and langd_film and langd_film > 0:
        skala = l_karta / langd_film
    d = pos * skala
    utg = normalisera(post.get('utgangsbrunn'))
    if utg == b or (utg != a and normalisera(post.get('startbrunn')) == b):
        d = l_karta - d
    return max(0.0, min(l_karta, d))


def _tal(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


# ------------------------------------------------ bedomning

def galler(mask_bed, man_bed):
    """(BEDOMNING, BED_TYP, STIL) - manuell bedomning gar fore maskinell."""
    m = (txt(man_bed) or '').strip().upper()[:1]
    if m in KLASSORDNING:
        return m, 'Manuell', '%s - Manuell' % m
    k = (txt(mask_bed) or '').strip().upper()[:1]
    if k not in KLASSORDNING:
        k = 'E'
    return k, 'Maskinell', '%s - Maskinell' % k


EGNA_FALT = [
    # namn,        typ,      langd, alias
    ('MASK_BED',   'TEXT',   2,   'Maskinell bedömning'),
    ('MAN_BED',    'TEXT',   2,   'Manuell bedömning'),
    ('BEDOMNING',  'TEXT',   2,   'Gällande bedömning'),
    ('BED_TYP',    'TEXT',   10,  'Bedömningstyp'),
    ('STIL',       'TEXT',   20,  'Symbologi (klass + typ)'),
    ('FRAN_BRUNN', 'TEXT',   50,  'Startbrunn (uppströms)'),
    ('TILL_BRUNN', 'TEXT',   50,  'Slutbrunn (nedströms)'),
    ('KLASSTEXT',  'TEXT',   40,  'Maskinell bedömning, klartext'),
    ('INDEX_TOT',  'DOUBLE', None, 'Totalindex (p/100 m)'),
    ('INDEX_K',    'DOUBLE', None, 'Konstruktionsindex (p/100 m)'),
    ('MAXGRAD_K',  'SHORT',  None, 'Konstruktion, maxgrad'),
    ('ANT_SKADOR', 'LONG',   None, 'Antal skador'),
    ('ANT_ANSL',   'LONG',   None, 'Antal anslutningar'),
    ('SKADOR',     'TEXT',   200, 'Skador (kod+grad)'),
    ('DRIFTATG',   'TEXT',   100, 'Driftåtgärd'),
    ('LANGD_M',    'DOUBLE', None, 'Längd (m)'),
    ('MATERIAL',   'TEXT',   50,  'Material'),
    ('DIMENSION',  'TEXT',   20,  'Dimension (mm)'),
    ('LEDN_TYP',   'TEXT',   30,  'Ledningstyp'),
    ('RELINAD',    'TEXT',   3,   'Relinad'),
    ('AVBRUTEN',   'TEXT',   3,   'Avbruten inspektion'),
    ('SVACKA_CM',  'LONG',   None, 'Svackdjup (cm)'),
    ('LUTNING',    'DOUBLE', None, 'Lutning (promille)'),
    ('ETAPP',      'LONG',   None, 'Etapp (åtgärdspaket)'),
    ('METOD',      'TEXT',   10,  'Metod (strumpa/schakt)'),
    ('KOSTNAD',    'DOUBLE', None, 'Kostnad (kr, strumpa)'),
    ('OMRADE',     'TEXT',   60,  'Område'),
    ('DATUM',      'TEXT',   10,  'Inspektionsdatum'),
    ('TV3_FIL',    'TEXT',   100, 'TV3-fil'),
    ('NR',         'LONG',   None, 'Sträcknummer i TV3-filen'),
    ('RAPPORT',    'TEXT',   254, 'Inspektionsprotokoll (PDF)'),
    ('VIDEO',      'TEXT',   254, 'Videofil'),
    ('VIDEO2',     'TEXT',   254, 'Videofil 2 (andra filmen i en sammanslagen stracka)'),
    ('ANT_FILM',   'LONG',   None, 'Antal inspektioner av sträckan'),
    ('ANT_DELAR',  'LONG',   None, 'Antal ledningsobjekt i kartan'),
    ('SRC_LAGER',  'TEXT',   100, 'Källager'),
    ('SRC_OID',    'LONG',   None, 'Käll-OID'),
    ('GIS_AVVIK',  'TEXT',   100, 'Avvikelse mot GIS-ledningen (typ/dimension)'),
    ('KLIPPT',     'TEXT',   100, 'Avbruten film: ritad del av kartlinjen'),
]


# Lager med svackor (punkt) och bakfall (linje) ur inklinometerprofilerna
GEMENSAMMA_FALT = [
    ('FRAN_BRUNN', 'TEXT',   50,  'Startbrunn (uppströms)'),
    ('TILL_BRUNN', 'TEXT',   50,  'Slutbrunn (nedströms)'),
]
SVACK_FALT = GEMENSAMMA_FALT + [
    ('SVACKA_CM',  'LONG',   None, 'Svackdjup (cm)'),
    ('SVACKLANGD', 'DOUBLE', None, 'Svacklängd (m)'),
    ('ANDEL_DIAM', 'DOUBLE', None, 'Svackdjup/diameter'),
    ('POS_M',      'DOUBLE', None, 'Position från utgångsbrunnen (m)'),
]
BAKFALL_FALT = GEMENSAMMA_FALT + [
    ('LANGD_M',    'DOUBLE', None, 'Längd (m)'),
    ('LUTNING',    'DOUBLE', None, 'Lutning mot flödet (‰)'),
    ('FRAN_M',     'DOUBLE', None, 'Från (m från utgångsbrunnen)'),
    ('TILL_M',     'DOUBLE', None, 'Till (m från utgångsbrunnen)'),
]
SLUTFALT = [
    ('UTG_BRUNN',  'TEXT',   50,  'Utgångsbrunn (kamerans start)'),
    ('MASK_BED',   'TEXT',   2,   'Maskinell bedömning'),
    ('OSAKER',     'TEXT',   3,   'Profil osäker'),
    ('MATERIAL',   'TEXT',   50,  'Material'),
    ('DIMENSION',  'TEXT',   20,  'Dimension'),
    ('NR',         'LONG',   None, 'Sträcknr'),
    ('KALLFIL',    'TEXT',   100, 'TV3-fil'),
    ('RAPPORT',    'TEXT',   254, 'Rapport (PDF)'),
]
SVACK_FALT = SVACK_FALT + SLUTFALT
BAKFALL_FALT = BAKFALL_FALT + SLUTFALT
LAGERNAMN_SVACKOR = 'Svackor'
LAGERNAMN_BAKFALL = 'Bakfall'
# Symbologi for de tva lagren, sparad fran ArcMap en gang (som bedomda_ledningar.lyr)
LYR_SVACKOR = os.path.join(ARCMAP_MAPP, 'svackor.lyr')
LYR_BAKFALL = os.path.join(ARCMAP_MAPP, 'bakfall.lyr')
SYMBOLOGI_TIPS_SVACKOR = [
    'Symbologi for svackor (spara som arcmap/svackor.lyr): Quantities > Graduated symbols pa',
    '  SVACKA_CM i tre steg (2-5, 5-10, >10 cm); rod farg nar ANDEL_DIAM > 0,5.',
    'Symbologi for bakfall (spara som arcmap/bakfall.lyr): bred linje graderad pa LUTNING.',
]


def _utdata(ut_fc):
    """Tolkar en utdatasokvag: (ut_fil, ut_ws, ut_namn, ar_shapefil, max_text, doman_ws).
    Vanlig mapp = shapefil. Geodatabas (.gdb/.mdb/.sde) och feature dataset i en sadan
    = featureklass med alias, NULL och vardelista."""
    ut_ws = os.path.dirname(ut_fc)
    ut_namn = os.path.basename(ut_fc)
    if not os.path.isdir(ut_ws) and not arcpy.Exists(ut_ws):
        raise RuntimeError('Utdatamappen finns inte: %s' % ut_ws)
    try:
        ar_shapefil = txt(arcpy.Describe(ut_ws).dataType) == 'Folder'
    except Exception:
        ar_shapefil = os.path.isdir(ut_ws) and not re.search(r'\.(gdb|mdb|sde)(\\|/|$)', ut_ws.lower())
    if ar_shapefil and not ut_namn.lower().endswith('.shp'):
        ut_namn = ut_namn + '.shp'
    ut_fil = os.path.join(ut_ws, ut_namn)
    max_text = 254 if ar_shapefil else 400
    # Domanen ligger i geodatabasen, aven nar utdata skrivs i ett feature dataset
    doman_ws = ut_ws
    if not ar_shapefil:
        try:
            if txt(arcpy.Describe(ut_ws).dataType) == 'FeatureDataset':
                doman_ws = os.path.dirname(ut_ws)
        except Exception:
            pass
    return ut_fil, ut_ws, ut_namn, ar_shapefil, max_text, doman_ws


def _radera_utdata(ut_fil):
    """Finns utdata redan: ta bort lagret ur kartan (ArcMap haller annars schemalas pa
    featureklassen) och radera den innan den skapas pa nytt."""
    if not arcpy.Exists(ut_fil):
        return
    mxd = _mxd()
    if mxd is not None:
        try:
            for l in arcpy.mapping.ListLayers(mxd):
                if getattr(l, 'supports', lambda x: False)('DATASOURCE') and \
                        os.path.normcase(txt(l.dataSource)) == os.path.normcase(txt(ut_fil)):
                    arcpy.mapping.RemoveLayer(arcpy.mapping.ListDataFrames(mxd)[0], l)
                    logg('  lagret "%s" togs bort ur kartan infor omkorningen' % txt(l.name))
        except Exception as e:
            logg('  kunde inte ta bort det gamla lagret ur kartan: %s' % txt(e))
    try:
        arcpy.Delete_management(ut_fil)
    except Exception as e:
        raise RuntimeError('Kan inte skriva over %s - ta bort lagret ur kartan och kor igen (%s)'
                           % (ut_fil, txt(e)))


def _skapa_fc(ut_fc, geometri, falt, sr):
    """Skapar en tom featureklass/shapefil med falten. Returnerar (ut_fil, ar_shapefil, max_text)."""
    ut_fil, ut_ws, ut_namn, ar_shapefil, max_text, doman_ws = _utdata(ut_fc)
    _radera_utdata(ut_fil)
    arcpy.CreateFeatureclass_management(ut_ws, ut_namn, geometri, '', 'DISABLED', 'DISABLED', sr)
    for namn, typ, langd, alias in falt:
        if typ == 'TEXT':
            arcpy.AddField_management(ut_fil, namn, typ,
                                      field_length=min(langd, max_text), field_alias=alias)
        else:
            arcpy.AddField_management(ut_fil, namn, typ, field_alias=alias)
    return ut_fil, ar_shapefil, max_text


def _utan_null(rad, typer, ar_shapefil):
    """Shapefiler kan inte lagra NULL i tal- och textfalt: tomma varden blir 0 resp. ''."""
    if not ar_shapefil:
        return rad
    ut = []
    for v, typ in zip(rad, typer):
        if v is None and typ not in ('DATE', 'SHAPE@'):
            v = '' if typ == 'TEXT' else 0
        ut.append(v)
    return ut


def skriv_svackor_bakfall(data, vagar, sr, svackor_ut, bakfall_ut, rapport_sokvag,
                          svacka_min_cm=2, geojson_ut=None):
    """Skriver svacklagret (punkt) och/eller bakfallslagret (linje) ur JSON-filens profilmatt.
    vagar = {frozenset(brunnspar): (punkter a->b, a, b)} for de par som matchats i kartan.
    Alla strackor i filen med ett matchat par tas med (aven syskon till avbrutna). Returnerar
    {'svackor': (fil, antal), 'bakfall': (fil, antal)} for det som skrevs."""
    ut = {}
    poster = data.get('strackor', [])
    rapporter = {}

    def _rapp(post):
        n = (post.get('fil'), post.get('nr'))
        if n not in rapporter:
            rapporter[n] = rapport_sokvag(post)
        return rapporter[n]

    def _slutfalt(post, max_text):
        def k(v, langd):
            return txt(v)[:min(langd, max_text)] if v is not None else ''
        return [k(post.get('utgangsbrunn'), 50), k(post.get('maskinell_bedomning') or 'E', 2),
                'Ja' if post.get('profil_osaker') else 'Nej',
                k(post.get('material'), 50), k(post.get('dimension'), 20), post.get('nr'),
                k(os.path.basename(txt(post.get('tv3_fil') or '')), 100), k(_rapp(post), 254)]

    def _geojson_namn(suffix):
        if not geojson_ut:
            return None
        stam, andelse = os.path.splitext(geojson_ut)
        return '%s_%s%s' % (stam, suffix, andelse or '.geojson')

    if svackor_ut:
        fil, ar_shp, max_text = _skapa_fc(svackor_ut, 'POINT', SVACK_FALT, sr)
        falt = ['SHAPE@'] + [n for n, t, l, a in SVACK_FALT]
        typer = ['SHAPE@'] + [t for n, t, l, a in SVACK_FALT]
        n = 0
        insert = arcpy.da.InsertCursor(fil, falt)
        try:
            for post in poster:
                djup = _tal(post.get('svackdjup_cm'))
                pos = _tal(post.get('svackpos_m'))
                if djup is None or pos is None or djup < svacka_min_cm:
                    continue
                par = par_for(post)[2]
                if par not in vagar:
                    continue
                pts, a, b = vagar[par]
                l_karta = _langd(pts)
                x, y = _punkt_vid(pts, _kartposition(post, pos, l_karta, a, b))
                insert.insertRow(_utan_null(
                    [arcpy.PointGeometry(arcpy.Point(x, y), sr),
                     txt(post.get('startbrunn') or '')[:50], txt(post.get('slutbrunn') or '')[:50],
                     int(round(djup)), post.get('svacklangd_m'), post.get('svacka_andel'), pos]
                    + _slutfalt(post, max_text), typer, ar_shp))
                n += 1
        finally:
            del insert
        logg('  %d svackor (>= %s cm) skrivna till %s' % (n, svacka_min_cm, fil))
        ut['svackor'] = (fil, n)
        g = _geojson_namn('svackor')
        if g:
            skriv_geojson(fil, falt[1:], g)
            logg('  svackor aven som GeoJSON: %s' % g)

    if bakfall_ut:
        fil, ar_shp, max_text = _skapa_fc(bakfall_ut, 'POLYLINE', BAKFALL_FALT, sr)
        falt = ['SHAPE@'] + [n for n, t, l, a in BAKFALL_FALT]
        typer = ['SHAPE@'] + [t for n, t, l, a in BAKFALL_FALT]
        n = 0
        insert = arcpy.da.InsertCursor(fil, falt)
        try:
            for post in poster:
                segment = post.get('bakfall_segment') or []
                if not segment:
                    continue
                par = par_for(post)[2]
                if par not in vagar:
                    continue
                pts, a, b = vagar[par]
                l_karta = _langd(pts)
                for seg in segment:
                    f0, t0 = _tal(seg.get('fran_m')), _tal(seg.get('till_m'))
                    if f0 is None or t0 is None:
                        continue
                    d0 = _kartposition(post, f0, l_karta, a, b)
                    d1 = _kartposition(post, t0, l_karta, a, b)
                    d0, d1 = min(d0, d1), max(d0, d1)
                    if d1 - d0 < 0.05:
                        continue
                    arr = arcpy.Array()
                    for x, y in _delstracka(pts, d0, d1):
                        arr.add(arcpy.Point(x, y))
                    insert.insertRow(_utan_null(
                        [arcpy.Polyline(arr, sr, False, False),
                         txt(post.get('startbrunn') or '')[:50], txt(post.get('slutbrunn') or '')[:50],
                         seg.get('langd_m'), seg.get('lutning_promille'), f0, t0]
                        + _slutfalt(post, max_text), typer, ar_shp))
                    n += 1
        finally:
            del insert
        logg('  %d bakfallssegment skrivna till %s' % (n, fil))
        ut['bakfall'] = (fil, n)
        g = _geojson_namn('bakfall')
        if g:
            skriv_geojson(fil, falt[1:], g)
            logg('  bakfall aven som GeoJSON: %s' % g)
    if ut and not (os.path.isfile(LYR_SVACKOR) and os.path.isfile(LYR_BAKFALL)):
        for rad in SYMBOLOGI_TIPS_SVACKOR:
            logg('  ' + rad)
    return ut


def rakna_om_bedomning(fc):
    """Uppdaterar BEDOMNING, BED_TYP och STIL utifran MASK_BED och MAN_BED.
    Returnerar (antal andrade, antal med manuell bedomning)."""
    n, manuella = 0, 0
    with arcpy.da.UpdateCursor(fc, ['MASK_BED', 'MAN_BED', 'BEDOMNING', 'BED_TYP', 'STIL']) as mark:
        for rad in mark:
            bed, typ, stil = galler(rad[0], rad[1])
            if typ == 'Manuell':
                manuella += 1
            if [rad[2], rad[3], rad[4]] != [bed, typ, stil]:
                rad[2], rad[3], rad[4] = bed, typ, stil
                mark.updateRow(rad)
                n += 1
    return n, manuella


def par_for(post):
    """Nyckeln for en post: brunnsparet, dar en grenrorsande (ingen brunn) gors unik med den fria
    andens koordinat - samma platshallarlittera (AG) kan sta for flera pastick fran samma brunn."""
    a, b = normalisera(post.get('startbrunn')), normalisera(post.get('slutbrunn'))
    xy = post.get('grenror_xy')
    if post.get('grenror') in ('start', 'slut') and xy:
        try:
            tag = '@%d,%d' % (round(float(xy[0])), round(float(xy[1])))
        except (TypeError, ValueError, IndexError):
            tag = ''
        if post['grenror'] == 'start':
            a = (a or '') + tag
        else:
            b = (b or '') + tag
    return a, b, frozenset((a, b))


def las_kartunderlag(json_in):
    """Laser JSON-filen. Returnerar (data, bedomda, antal_per_par) dar bedomda ar
    {frozenset(brunnspar): post} med den varsta bedomningen per par."""
    if not os.path.isfile(json_in):
        raise RuntimeError('JSON-filen finns inte: %s\nKor tv3_analys.py forst.' % json_in)
    with io.open(json_in, 'r', encoding='utf-8') as f:
        data = json.loads(f.read())

    bedomda, antal_per_par = {}, {}
    for post in data.get('strackor', []):
        a, b, par = par_for(post)
        if not a or not b or a == b:
            continue
        antal_per_par[par] = antal_per_par.get(par, 0) + 1
        tidigare = bedomda.get(par)
        if tidigare is None or (KLASSORDNING.get(txt(post.get('maskinell_bedomning')), 9)
                                < KLASSORDNING.get(txt(tidigare.get('maskinell_bedomning')), 9)):
            bedomda[par] = post
    return data, bedomda, antal_per_par


def skriv_geojson(fc, falt, geojson_ut, decimaler=7):
    """Skriver featureklassen som GeoJSON enligt RFC 7946: WGS84 (EPSG:4326), 2D, utan
    crs-medlem. Det ar vad webb-GIS forvantar sig; ArcMaps egen export behaller kartans
    koordinatsystem och Z/M, vilket manga webbkartor inte laser."""
    wgs84 = arcpy.SpatialReference(4326)
    poster = []
    with arcpy.da.SearchCursor(fc, ['SHAPE@'] + list(falt)) as mark:
        for rad in mark:
            geom = rad[0]
            if geom is None:
                continue
            try:
                geom = geom.projectAs(wgs84)
            except Exception as e:
                raise RuntimeError('Kunde inte projicera till WGS84: %s' % txt(e))
            if txt(getattr(geom, 'type', '')).lower() == 'point':
                p = geom.firstPoint
                geometri = {'type': 'Point', 'coordinates': [round(p.X, decimaler), round(p.Y, decimaler)]}
            else:
                delar = []
                for del_ in geom:
                    pts = [[round(p.X, decimaler), round(p.Y, decimaler)] for p in del_ if p is not None]
                    if len(pts) >= 2:
                        delar.append(pts)
                if not delar:
                    continue
                if len(delar) == 1:
                    geometri = {'type': 'LineString', 'coordinates': delar[0]}
                else:
                    geometri = {'type': 'MultiLineString', 'coordinates': delar}
            egenskaper = {}
            for namn, v in zip(falt, rad[1:]):
                if isinstance(v, bytes):
                    v = txt(v)
                if isinstance(v, TEXTTYP):
                    v = v.strip()
                elif isinstance(v, float) and v != v:      # NaN -> null
                    v = None
                egenskaper[namn] = v
            poster.append({'type': 'Feature', 'geometry': geometri, 'properties': egenskaper})
    text = json.dumps({'type': 'FeatureCollection', 'features': poster}, ensure_ascii=False,
                      default=txt)                  # datum m.m. som text
    with io.open(geojson_ut, 'w', encoding='utf-8') as f:
        f.write(txt(text))
    return len(poster)


# =====================================================================
# Huvudfunktioner
# =====================================================================

def uppdatera(ut_fc):
    """Raknar om de harledda bedomningsfalten i ett befintligt lager.
    ut_fc: lagernamn i kartan, lagerobjekt eller sokvag till featureklassen."""
    mal = None
    try:
        mal = kalla(hitta_lager(ut_fc))[0]
    except Exception:
        pass
    if not mal or not arcpy.Exists(mal):
        mal = ut_fc if arcpy.Exists(ut_fc) else txt(ut_fc) + '.shp'
    if not arcpy.Exists(mal):
        raise RuntimeError('Hittar inte %s - skapa lagret forst' % ut_fc)
    n, manuella = rakna_om_bedomning(mal)
    logg('%d objekt uppdaterade i %s' % (n, mal))
    logg('  %d objekt har manuell bedomning' % manuella)
    try:
        arcpy.RefreshActiveView()
    except Exception:
        pass
    return mal


def skapa(json_in, ledningslager, brunnslager, brunn_id, ut_fc,
          omradeslager=None, csv_ut=None, lyr_fil=None,
          tolerans=2.0, marginal=100.0, max_hopp=2, urval='INTERSECT',
          kopiera_falt=None, lagg_till_i_kartan=True,
          rapportmapp=None, filmmapp=None, geojson_ut=None, max_delar=8,
          svackor_ut=None, bakfall_ut=None, svacka_min_cm=SVACKA_MIN_CM,
          typ_falt=None, dim_falt=None, klipp_avbrutna=KLIPP_AVBRUTNA):
    """Bygger ledningslagret. Returnerar sokvagen till den skrivna featureklassen.
    svackor_ut/bakfall_ut: valfria lager med svackor (punkt) och bakfall (linje).
    typ_falt/dim_falt: falt i ledningslagret med ledningstyp och dimension - vid parallella
    ledningar (spill/dag i samma schakt) tas den som stammer med filmen (GIS_AVVIK annars).
    klipp_avbrutna: avbruten/ofullstandig film ritas bara sa langt kameran kom (KLIPPT)."""
    kopiera_falt = kopiera_falt or []
    if not float(tolerans) > 0:
        raise RuntimeError('Toleransen maste vara storre an 0 m')
    if isinstance(ledningslager, (TEXTTYP, bytes)):
        ledningslager = [ledningslager]
    if isinstance(brunnslager, (TEXTTYP, bytes)):
        brunnslager = [brunnslager]

    # ---------------------------------------------------- 1. Lager
    logg('Letar upp lager')
    led_lager = [hitta_lager(n) for n in ledningslager]
    brunn_lager = [hitta_lager(n) for n in brunnslager]
    omrade = None
    if omradeslager:
        src, dq = kalla(hitta_lager(omradeslager))
        arcpy.MakeFeatureLayer_management(src, 'lyr_omr', dq)
        omrade = 'lyr_omr'

    # ---------------------------------------------------- 2. JSON
    logg('Laser %s' % json_in)
    data, bedomda, antal_per_par = las_kartunderlag(json_in)
    logg('  %d stracker i filen, %d unika brunnspar'
         % (len(data.get('strackor', [])), len(bedomda)))
    if not bedomda:
        raise RuntimeError('Inga brunnspar kunde lasas ur JSON-filen')

    # ---------------------------------------------------- 3. Brunnar
    # Allt raknas i ledningslagrets koordinatsystem; brunnar i ett annat system projiceras.
    d0 = arcpy.Describe(kalla(led_lager[0])[0])
    sr = d0.spatialReference

    def _sr_namn(d):
        try:
            return txt(d.spatialReference.name)
        except Exception:
            return ''

    brunnar = []
    for lyr in brunn_lager:
        idfalt = hitta_falt(lyr, brunn_id)
        src, dq = kalla(lyr)
        projicera = _sr_namn(arcpy.Describe(src)) not in ('', _sr_namn(d0))
        if projicera:
            logg('  %s har annat koordinatsystem an ledningslagret - projiceras'
                 % txt(getattr(lyr, 'name', lyr)))
        arcpy.MakeFeatureLayer_management(src, 'lyr_br', dq)
        if omrade is not None:
            arcpy.SelectLayerByLocation_management(
                'lyr_br', 'INTERSECT', omrade, '%s Meters' % marginal, 'NEW_SELECTION')
        n = 0
        with arcpy.da.SearchCursor('lyr_br', [idfalt, 'SHAPE@' if projicera else 'SHAPE@XY']) as mark:
            for bid, geom in mark:
                nid = normalisera(bid)
                if not nid or geom is None:
                    continue
                if projicera:
                    pt = geom.projectAs(sr).firstPoint
                    xy = (pt.X, pt.Y) if pt is not None else None
                else:
                    xy = geom
                if xy and xy[0] is not None:
                    brunnar.append((nid, xy[0], xy[1]))
                    n += 1
        logg('  %s: %d brunnar' % (txt(getattr(lyr, 'name', lyr)), n))
        arcpy.Delete_management('lyr_br')

    logg('  %d brunnar totalt' % len(brunnar))
    if not brunnar:
        raise RuntimeError('Inga brunnar hittades - kontrollera brunnsfaltet och koordinatsystem')

    brunns_id = set(b[0] for b in brunnar)

    # Tackning: hur manga av JSON-filens brunnar finns i brunnslagren? Saknas hela
    # brunnstyper (t.ex. alla KRB/KTB) ligger de troligen i ett annat lager.
    json_brunnar = set()
    grenror_namn = set()           # littera som ar pastick pa en annan ledning (grenror), inte brunnar
    for par, post in bedomda.items():
        json_brunnar.update(n for n in par if '@' not in n)    # '@' = grenrorsande med koordinat
        if post.get('grenror'):
            grenror_namn.add(normalisera(post.get('startbrunn') if post['grenror'] == 'start'
                                         else post.get('slutbrunn')))
    json_brunnar -= grenror_namn
    saknade = sorted(json_brunnar - brunns_id)
    if grenror_namn:
        logg('  %d grenrorsanslutningar (ingen brunn) soks som ledning med fri ande fran den kanda brunnen'
             % len(grenror_namn))
    logg('  %d av %d brunnar i JSON-filen finns i brunnslagren'
         % (len(json_brunnar) - len(saknade), len(json_brunnar)))
    if saknade:
        prefix = {}
        for b in saknade:
            pfx = re.match(r'[A-Z\u00c5\u00c4\u00d6]+', b)
            pfx = pfx.group(0) if pfx else b
            prefix[pfx] = prefix.get(pfx, 0) + 1
        topp = sorted(prefix.items(), key=lambda kv: -kv[1])[:8]
        logg('  saknade brunnar per typ: %s' % ', '.join('%s %d' % kv for kv in topp))
        # Manga saknade av samma typ tyder pa att ett helt lager fattas; enstaka
        # saknade ar normalt (nytt littera, brunn utanfor kartan, anslutning "AG").
        if topp[0][1] >= 10:
            logg('  (saknas en hel brunnstyp - lagg till lagret den ligger i under Brunnslager)')
        elif len(saknade) <= 20:
            logg('  saknade: %s' % ', '.join(saknade))

    # Bara JSON-filens brunnar behovs for matchningen (forsta forekomsten vinner).
    # Samma id pa flera stallen (t.ex. i tva brunnslager) langre isar an toleransen loggas.
    sokta = {}
    dubbla = []
    for bid, x, y in brunnar:
        if bid not in json_brunnar:
            continue
        if bid not in sokta:
            sokta[bid] = (x, y)
        elif _avst(sokta[bid], (x, y)) > tolerans:
            dubbla.append('%s (%.0f m)' % (bid, _avst(sokta[bid], (x, y))))
    if dubbla:
        logg('  VARNING: %d brunnar finns pa flera stallen i brunnslagren, forsta anvands: %s'
             % (len(dubbla), ', '.join(dubbla[:10]) + (' ...' if len(dubbla) > 10 else '')))

    # ---------------------------------------------------- 4. Utdata
    # Utdata skrivs alltid i 2D: Z i ledningsnatverket ar odefinierat (-9999) och
    # 3D-shapefiler/GeoJSON med fyra koordinater stoppar de flesta webb-GIS.
    har_z = False

    ut_fil, ut_ws, ut_namn, ar_shapefil, max_text, doman_ws = _utdata(ut_fc)
    for extra in (svackor_ut, bakfall_ut):
        if extra:
            _utdata(extra)          # kontrollerar att mappen finns innan vi borjar

    # Manuella bedomningar per brunnspar fran en tidigare korning bevaras
    tidigare_manuella = {}
    if arcpy.Exists(ut_fil):
        try:
            with arcpy.da.SearchCursor(ut_fil, ['FRAN_BRUNN', 'TILL_BRUNN', 'MAN_BED']) as mark:
                for fran, till, man in mark:
                    if man and txt(man).strip():
                        tidigare_manuella[frozenset((normalisera(fran), normalisera(till)))] = txt(man).strip()
            if tidigare_manuella:
                logg('  %d manuella bedomningar fran forra korningen bevaras' % len(tidigare_manuella))
        except Exception as e:
            logg('  kunde inte lasa tidigare manuella bedomningar: %s' % txt(e))

    if ar_shapefil:
        logg('  utdata blir en shapefil - textfalt kapas till %d tecken' % max_text)
        logg('  (skapa en filgeodatabas om du vill ha faltalias som "Maskinell bedomning")')

    _radera_utdata(ut_fil)

    arcpy.CreateFeatureclass_management(
        ut_ws, ut_namn, 'POLYLINE', '', 'DISABLED',
        'ENABLED' if har_z else 'DISABLED', sr)

    for namn, typ, langd, alias in EGNA_FALT:
        if typ == 'TEXT':
            arcpy.AddField_management(ut_fil, namn, typ,
                                      field_length=min(langd, max_text), field_alias=alias)
        else:
            arcpy.AddField_management(ut_fil, namn, typ, field_alias=alias)

    if not ar_shapefil:
        try:
            if 'TV3_BEDOMNING' not in [d.name for d in arcpy.da.ListDomains(doman_ws)]:
                arcpy.CreateDomain_management(doman_ws, 'TV3_BEDOMNING', 'Prioritetsklass A-E',
                                              'TEXT', 'CODED')
                for k in KLASSER:
                    arcpy.AddCodedValueToDomain_management(
                        doman_ws, 'TV3_BEDOMNING', k, data.get('klasser', {}).get(k, k))
            arcpy.AssignDomainToField_management(ut_fil, 'MAN_BED', 'TV3_BEDOMNING')
            logg('  vardelista TV3_BEDOMNING kopplad till MAN_BED')
        except Exception as e:
            logg('  kunde inte skapa vardelista: %s' % txt(e))

    typkarta = {'String': 'TEXT', 'Integer': 'LONG', 'SmallInteger': 'SHORT',
                'Double': 'DOUBLE', 'Single': 'FLOAT', 'Date': 'DATE',
                'GUID': 'GUID', 'Blob': 'BLOB'}
    # Falt att kopiera fran ledningslagren: typen tas fran forsta lagret som har faltet.
    # Per lager laser vi bara de falt som finns dar; ovriga blir NULL.
    kallfalt_per_lager = {}
    for lyr in led_lager:
        kallfalt_per_lager[txt(getattr(lyr, 'name', lyr))] = dict(
            (f.name.upper(), f) for f in arcpy.ListFields(kalla(lyr)[0]))
    kopiera = []          # (kallnamn versalt, utdatanamn, typ)
    for namn in kopiera_falt:
        f = None
        for kf in kallfalt_per_lager.values():
            f = kf.get(txt(namn).upper())
            if f:
                break
        if not f:
            logg('  VARNING: faltet %s finns inte i nagot ledningslager - hoppas over' % txt(namn))
            continue
        typ = typkarta.get(f.type)
        if not typ:
            logg('  VARNING: falttypen %s stods inte (%s) - hoppas over' % (f.type, f.name))
            continue
        fore = set(x.name.upper() for x in arcpy.ListFields(ut_fil))
        if typ == 'TEXT':
            arcpy.AddField_management(ut_fil, f.name, typ,
                                      field_length=min(f.length or 255, max_text))
        else:
            arcpy.AddField_management(ut_fil, f.name, typ)
        # Shapefiler kortar namn over 10 tecken - lasa vad faltet faktiskt heter i utdata
        nya = [x.name for x in arcpy.ListFields(ut_fil) if x.name.upper() not in fore]
        utnamn = nya[0] if len(nya) == 1 else f.name
        kopiera.append((f.name.upper(), utnamn, typ))

    ut_falt = ['SHAPE@'] + [n for n, t, l, a in EGNA_FALT] + [u for k, u, typ in kopiera]
    ut_typer = ['SHAPE@'] + [t for n, t, l, a in EGNA_FALT] + [typ for k, u, typ in kopiera]

    def klipp(v, langd):
        return txt(v)[:min(langd, max_text)] if v is not None else ''

    def utan_null(rad):
        """Shapefiler kan inte lagra NULL i tal- och textfalt: tomma varden blir
        0 respektive '' (samma sak ArcGIS gor vid export till shapefil)."""
        if not ar_shapefil:
            return rad
        ut = []
        for v, typ in zip(rad, ut_typer):
            if v is None and typ not in ('DATE', 'SHAPE@'):
                v = '' if typ == 'TEXT' else 0
            ut.append(v)
        return ut

    if ar_shapefil:
        logg('  tomma tal (t.ex. lutning utan profil) skrivs som 0 i shapefil')

    # ---------------------------------------------------- 5. Bygg natverk av ledningsbitar
    nat = Natverk(sokta, tolerans)
    extra_per_oid = {}
    attr_per_oid = {}          # (lager, oid) -> (ledningstyp, dimension) nar typ_falt/dim_falt angetts
    n_lednkoll = n_nara = n_bitar = 0
    for lyr in led_lager:
        namn = txt(getattr(lyr, 'name', lyr))
        src, dq = kalla(lyr)
        arcpy.MakeFeatureLayer_management(src, 'lyr_led', dq)
        if omrade is not None:
            arcpy.SelectLayerByLocation_management(
                'lyr_led', urval, omrade, '', 'NEW_SELECTION')
        antal = int(arcpy.GetCount_management('lyr_led').getOutput(0))
        logg('  %s: %d ledningar att ga igenom' % (namn, antal))

        # Bara de kopierade falt som finns i just detta lager lases; ovriga blir NULL
        har = kallfalt_per_lager.get(namn, {})
        lasbara = [(i, har[k].name) for i, (k, u, typ) in enumerate(kopiera) if k in har]
        attrfalt = []                     # ledningstyp och dimension for val mellan parallella ledningar
        for f in (typ_falt, dim_falt):
            fn = har.get(txt(f).upper()) if f else None
            if f and not fn:
                logg('  VARNING: faltet %s finns inte i %s - typ/dimension jamfors inte dar' % (txt(f), namn))
            attrfalt.append(fn.name if fn else None)
        attrlas = [fn for fn in attrfalt if fn]
        with arcpy.da.SearchCursor('lyr_led', ['OID@', 'SHAPE@'] + [n for i, n in lasbara] + attrlas) as mark:
            for rad in mark:
                oid, geom = rad[0], rad[1]
                if attrlas:
                    v = dict(zip(attrlas, rad[2 + len(lasbara):]))
                    attr_per_oid[(namn, oid)] = (v.get(attrfalt[0]), v.get(attrfalt[1]))
                n_lednkoll += 1
                if geom is None:
                    continue
                # Hoppa snabbt over ledningar langt fran alla sokta brunnar
                try:
                    e = geom.extent
                    if not nat.inom_bbox(e.XMin, e.YMin, e.XMax, e.YMax):
                        continue
                except Exception:
                    pass
                n_nara += 1
                extra = [None] * len(kopiera)
                for (i, n), v in zip(lasbara, rad[2:]):
                    extra[i] = v
                extra_per_oid[(namn, oid)] = extra
                for del_ in geom:
                    punkter = [(p.X, p.Y, p.Z if har_z else None) for p in del_ if p is not None]
                    n_bitar += nat.lagg_till(punkter, namn, oid)

        arcpy.Delete_management('lyr_led')

    if omrade is not None:
        arcpy.Delete_management('lyr_omr')

    logg('  %d ledningar genomgangna, %d nara brunnarna, %d bitar i natverket'
         % (n_lednkoll, n_nara, n_bitar))

    # ---------------------------------------------------- 6. Sok vag per brunnspar och skriv
    utdata_mapp = os.path.dirname(os.path.abspath(json_in))
    if rapportmapp:
        logg('  rapporter hamtas fran %s' % rapportmapp)
    if filmmapp:
        logg('  filmer hamtas fran %s' % filmmapp)

    def rapport_sokvag(post):
        """Sokvag till strackans PDF: i rapportmapp om angiven, annars relativt JSON-filen."""
        r = post.get('rapport')
        if not r:
            return ''
        r = txt(r).replace('/', os.sep)
        if rapportmapp:
            return os.path.join(txt(rapportmapp), os.path.basename(r))
        return r if os.path.isabs(r) else os.path.join(utdata_mapp, r)

    def video_sokvag(post, suffix=''):
        """Sokvag till filmen: filnamnet i filmmapp om angiven, annars som tv3_analys fann den.
        suffix '_b' ger den andra filmen i en sammanslagen stracka (videofil_b/video_sokvag_b)."""
        namn = txt(post.get('videofil' + suffix) or '') or os.path.basename(txt(post.get('video_sokvag' + suffix) or ''))
        if filmmapp:
            return os.path.join(txt(filmmapp), namn) if namn else ''
        return txt(post.get('video_sokvag' + suffix) or '')

    traffade = set()
    vagar = {}                 # par -> (punkter a->b, a, b) for svack- och bakfallslagren
    n_skrivna = n_flerdelade = n_grenror = n_avvik = n_klippta = n_via = 0
    avvik_lista = []
    if attr_per_oid:
        logg('  ledningstyp/dimension ur %s jamfors med filmen vid val mellan parallella ledningar'
             % '/'.join(txt(f) for f in (typ_falt, dim_falt) if f))

    insert = arcpy.da.InsertCursor(ut_fil, ut_falt)
    try:
        for par, s in bedomda.items():
            a, b = normalisera(s.get('startbrunn')), normalisera(s.get('slutbrunn'))
            poang = None
            if attr_per_oid:
                def poang(lager, oid, s=s):
                    t = attr_per_oid.get((lager, oid))
                    return _avvikelse(s, t[0], t[1])[0] if t else 0
            # Passerar filmen flera brunnar enligt tv3_analys GIS-koppling (ingen direkt ledning, t.ex.
            # en avbruten film som gick forbi brunnar utan att stanna) far vagen passera just sa manga
            via = [v for v in (s.get('gis_via') or []) if v]
            hopp_s = max(max_hopp, len(via) + 1)
            delar_s = max(max_delar, 4 * (len(via) + 1))
            vagen = nat.vag(a, b, hopp_s, delar_s, poang)
            if vagen and hopp_s > max_hopp:
                n_via += 1
            grenror = s.get('grenror') if s.get('grenror') in ('start', 'slut') else None
            vand = False
            if not vagen and grenror and s.get('grenror_xy'):
                # anslutning till grenror: vag fran den kanda brunnen till den fria anden som
                # tv3_analys hittade i GIS-exporten (koordinat i kartunderlaget)
                try:
                    gx, gy = float(s['grenror_xy'][0]), float(s['grenror_xy'][1])
                    vagen = nat.vag_till_punkt(b if grenror == 'start' else a, gx, gy, hopp_s, delar_s,
                                               poang=poang)
                except (TypeError, ValueError, IndexError):
                    vagen = None
                if vagen:
                    n_grenror += 1
                    vand = grenror == 'start'      # vagen soktes fran slutbrunnen
            if not vagen:
                continue
            pts = sla_ihop(vagen)
            if vand:
                pts = pts[::-1]                    # orientera start -> slut
                vagen = vagen[::-1]                # sa att SRC_LAGER/SRC_OID och extrafalt tas vid startbrunnen
            if len(pts) < 2:
                continue
            n_objekt = len(set((lager, oid) for nod, p, lager, oid in vagen))

            # Avvikelse mot GIS-ledningen (annan typ/dimension an filmen) - aven efter basta val
            avvik = ''
            if attr_per_oid:
                texter = []
                for lager, oid in sorted(set((l, o) for n, p, l, o in vagen), key=lambda k: (k[0], k[1])):
                    t = attr_per_oid.get((lager, oid))
                    if t:
                        tx = _avvikelse(s, t[0], t[1])[1]
                        if tx and tx not in texter:
                            texter.append(tx)
                avvik = '; '.join(texter)
                if avvik:
                    n_avvik += 1
                    avvik_lista.append('%s -> %s (%s nr %s): %s' % (a, b, txt(s.get('fil')), s.get('nr'), avvik))

            # Avbruten/ofullstandig film: bara den filmade delen fran kamerans brunn
            hela_pts = pts                     # hela vagen a -> b (svackor/bakfall skalas mot den)
            klippt = ''
            langd_film = _tal(s.get('langd_m'))
            l_karta = _langd(pts)
            if (klipp_avbrutna and (s.get('avbruten') or s.get('ofullstandig')) and langd_film
                    and 0 < langd_film < l_karta - max(tolerans, 0.5)):
                utg = normalisera(s.get('utgangsbrunn'))
                fran_b = utg == b or (utg != a and normalisera(s.get('startbrunn')) == b)
                klippt_pts = _klipp_pts(pts, l_karta - langd_film, l_karta) if fran_b else _klipp_pts(pts, 0.0, langd_film)
                if len(klippt_pts) >= 2:
                    pts = klippt_pts
                    klippt = '%.1f av %.1f m fran %s' % (langd_film, l_karta, b if fran_b else a)
                    n_klippta += 1

            arr = arcpy.Array()
            for x, y, z in pts:
                arr.add(arcpy.Point(x, y, z if har_z else None))
            ny = arcpy.Polyline(arr, sr, har_z, False)

            mask = txt(s.get('maskinell_bedomning') or 'E')[:2]
            plain = frozenset((a, b))          # manuella bedomningar i lagret ar nycklade pa litterat
            man = tidigare_manuella.get(par, '') or tidigare_manuella.get(plain, '')
            if not man:
                # Manuell bedomning ifylld i Excel (tv3_analys --manuell) foljer med via JSON-filen
                m_json = re.match(r'\s*([A-Ea-e])(?![A-Za-z\u00c5\u00c4\u00d6\u00e5\u00e4\u00f6])',
                                  txt(s.get('manuell_bedomning') or ''))
                if m_json:
                    man = m_json.group(1).upper()
            bed, bed_typ, stil = galler(mask, man)
            lager0, oid0 = vagen[0][2], vagen[0][3]
            extra = extra_per_oid.get((lager0, oid0), [None] * len(kopiera))

            insert.insertRow(utan_null([
                ny, mask, man, bed, bed_typ, stil,
                klipp(s.get('startbrunn'), 50), klipp(s.get('slutbrunn'), 50),
                klipp(s.get('klasstext'), 40),
                s.get('totalindex'), s.get('konstruktionsindex'),
                s.get('maxgrad_konstruktion'), s.get('antal_skador'), s.get('antal_anslutningar'),
                klipp(s.get('skador'), 200), klipp(s.get('driftatgard'), 100),
                s.get('langd_m'), klipp(s.get('material'), 50),
                klipp(s.get('dimension'), 20), klipp(s.get('ledningstyp'), 30),
                'Ja' if s.get('relinad') else 'Nej',
                'Ja' if s.get('avbruten') else 'Nej',
                s.get('svackdjup_cm'), s.get('lutning_promille'),
                s.get('etapp'), klipp(s.get('metod'), 10), s.get('kostnad_kr'),
                klipp(s.get('omrade'), 60), klipp(s.get('datum'), 10),
                klipp(os.path.basename(txt(s.get('tv3_fil') or '')), 100),
                s.get('nr'),
                klipp(rapport_sokvag(s), 254), klipp(video_sokvag(s), 254), klipp(video_sokvag(s, '_b'), 254),
                antal_per_par.get(par, 1), n_objekt,
                lager0[:100], oid0,
                klipp(avvik, 100), klipp(klippt, 100),
            ] + extra))
            traffade.add(par)
            traffade.add(plain)
            vagar[par] = ([(x, y) for x, y, z in hela_pts], a, b)
            n_skrivna += 1
            if n_objekt > 1:
                n_flerdelade += 1
    finally:
        del insert

    logg('  %d objekt skrivna till %s' % (n_skrivna, ut_fil))
    if n_grenror:
        logg('  %d grenrorsanslutningar klippta ut fran brunnen till den fria anden' % n_grenror)
    if n_via:
        logg('  %d strackor passerar fler brunnar an max hopp - tillatet eftersom GIS-kopplingen i'
             ' tv3_analys gick via dem (gis_via)' % n_via)
    if n_klippta:
        logg('  %d avbrutna/ofullstandiga filmer ritade bara sa langt kameran kom (falt KLIPPT)' % n_klippta)
    if attr_per_oid:
        if n_avvik:
            logg('  OBS: %d strackor ligger pa en GIS-ledning med annan typ eller dimension an filmen'
                 ' (falt GIS_AVVIK) - kontrollera littera eller kartan:' % n_avvik)
            for rad in avvik_lista[:30]:
                logg('    ' + rad)
            if len(avvik_lista) > 30:
                logg('    ... och %d till' % (len(avvik_lista) - 30))
        else:
            logg('  alla strackor ligger pa GIS-ledningar med samma typ och dimension som filmen')
    logg('  %d av %d brunnspar matchade (%d sammansatta av flera ledningsobjekt)'
         % (len(traffade), len(bedomda), n_flerdelade))
    n_hoppade = len(data.get('strackor', [])) - sum(antal_per_par.values())
    if n_hoppade:
        logg('  %d stracker i JSON-filen saknar brunnspar (tomt littera eller samma brunn i bada'
             ' andar) och kan inte laggas i kartan' % n_hoppade)
    tappade_manuella = {}
    if tidigare_manuella:
        tappade_manuella = dict((par, m) for par, m in tidigare_manuella.items() if par not in traffade)
        logg('  %d manuella bedomningar aterstallda'
             % sum(1 for par in traffade if par in tidigare_manuella))
        if tappade_manuella:
            logg('  VARNING: %d manuella bedomningar fran forra korningen hor till par som inte'
                 ' matchades nu - de finns i CSV:n over omatchade (kolumn manuell_bedomning)'
                 % len(tappade_manuella))

    # ---------------------------------------------------- 7. Omatchade
    def avst(bid):
        """Avstand fran brunnen till narmaste ledning, som text."""
        if bid not in brunns_id:
            return ''
        d = nat.narmast.get(bid)
        return ('%.1f' % d) if d is not None else ('>%.0f' % nat.sokradie)

    omatchade = []
    for par, s in bedomda.items():
        if par in traffade:
            continue
        a, b = normalisera(s.get('startbrunn')), normalisera(s.get('slutbrunn'))
        if s.get('grenror') in ('start', 'slut'):
            xy = s.get('grenror_xy')
            try:
                gx, gy = (float(xy[0]), float(xy[1])) if xy else (None, None)
            except (TypeError, ValueError, IndexError):
                gx = gy = None
            if gx is None:
                diagnos = 'grenror: ingen GIS-ledning med fri ande kopplad i analysen (kor med gis: i listfilen)'
            elif nat.fri_ande_nara(gx, gy, nat.tol + GRENROR_RADIE_EXTRA) is None:
                diagnos = 'grenror: ingen fri ledningsande vid (%.1f, %.1f) inom toleransen' % (gx, gy)
            else:
                diagnos = 'grenror: ingen vag fran brunnen till den fria anden (hoj max hopp?)'
        elif a in brunns_id and b in brunns_id:
            via = [v for v in (s.get('gis_via') or []) if v]
            diagnos = nat.diagnos(a, b, max(max_hopp, len(via) + 1), max(max_delar, 4 * (len(via) + 1)))
            if via:
                diagnos += ' (GIS-vagen i analysen gick via %s)' % ', '.join(txt(v) for v in via)
        elif a in brunns_id or b in brunns_id:
            diagnos = 'brunnen %s finns inte i brunnslagren' % (b if a in brunns_id else a)
        else:
            diagnos = 'ingen av brunnarna finns i brunnslagren'
        omatchade.append([a, b,
                          'JA' if a in brunns_id else 'NEJ',
                          'JA' if b in brunns_id else 'NEJ',
                          'JA' if (a in brunns_id or b in brunns_id) else 'NEJ',
                          avst(a), avst(b),
                          txt(s.get('maskinell_bedomning')),
                          txt(s.get('fil')), diagnos, tappade_manuella.get(par, '')])
    # Manuellt bedomda par som inte finns i JSON-filen langre (t.ex. annat urval)
    for par, m in tappade_manuella.items():
        if par not in bedomda:
            a, b = sorted(par)
            omatchade.append([a, b, '', '', '', '', '', '', '',
                              'paret finns inte i JSON-filen', m])
    omatchade.sort(key=lambda r: (r[4] != 'JA', KLASSORDNING.get(r[7], 9), r[0]))

    if csv_ut:
        with io.open(csv_ut, 'w', encoding='cp1252', errors='replace') as f:
            f.write('fran;till;fran_finns;till_finns;nagon_finns;'
                    'fran_avstand_m;till_avstand_m;maskinell_bedomning;kallfil;diagnos;'
                    'manuell_bedomning\n')
            for r in omatchade:
                f.write(';'.join(txt(v).replace(';', ',') for v in r) + '\n')
        logg('  %d omatchade par -> %s' % (len(omatchade), csv_ut))
    else:
        logg('  %d omatchade par' % len(omatchade))

    # Varfor? Bada brunnarna pa en ledning men ingen vag = ledningen ligger i ett annat
    # lager, ar bruten, eller passerar fler brunnar an max_hopp. En brunn en bit fran
    # ledningen = hoj toleransen.
    bada_brunnar = [r for r in omatchade if r[2] == 'JA' and r[3] == 'JA']
    pa_ledning = [r for r in bada_brunnar
                  if _tal(r[5]) is not None and _tal(r[5]) <= tolerans
                  and _tal(r[6]) is not None and _tal(r[6]) <= tolerans]
    nara = [r for r in bada_brunnar if r not in pa_ledning
            and _tal(r[5]) is not None and _tal(r[6]) is not None
            and max(_tal(r[5]), _tal(r[6])) <= 3 * tolerans]
    bada_brunnar = [r for r in bada_brunnar if r[9] != 'paret finns inte i JSON-filen']
    logg('  varav %d har bada brunnarna i kartan%s' % (len(bada_brunnar), ':' if bada_brunnar else ''))
    if pa_ledning:
        logg('    %d dar bada brunnarna ligger pa en ledning men ingen vag hittades'
             ' (annat ledningslager? bruten ledning? fler an %d brunnar emellan?)'
             % (len(pa_ledning), max_hopp))
    if nara:
        logg('    %d dar en brunn ligger %.0f-%.0f m fran ledningen - prova tolerans %.0f m'
             % (len(nara), tolerans, 3 * tolerans, 3 * tolerans))
    for r in bada_brunnar[:12]:
        logg('    %s - %s (klass %s): %s' % (r[0], r[1], r[7], r[9]))

    # ---------------------------------------------------- 8. GeoJSON for webb-GIS
    if geojson_ut:
        n_geo = skriv_geojson(ut_fil, [f for f in ut_falt if f != 'SHAPE@'], geojson_ut)
        logg('  %d objekt skrivna till %s (GeoJSON, WGS84, 2D)' % (n_geo, geojson_ut))

    # ---------------------------------------------------- 8b. Svackor och bakfall
    extra_lager = {}
    if svackor_ut or bakfall_ut:
        logg('Svackor och bakfall')
        extra_lager = skriv_svackor_bakfall(data, vagar, sr, svackor_ut, bakfall_ut,
                                            rapport_sokvag, svacka_min_cm, geojson_ut)

    # ---------------------------------------------------- 9. Karta
    if lagg_till_i_kartan:
        mxd = _mxd()
        if mxd is not None:
            try:
                df = arcpy.mapping.ListDataFrames(mxd)[0]
                for nyckel, namn, lyr_extra in (('bakfall', LAGERNAMN_BAKFALL, LYR_BAKFALL),
                                                ('svackor', LAGERNAMN_SVACKOR, LYR_SVACKOR)):
                    if nyckel in extra_lager:
                        l_extra = arcpy.mapping.Layer(extra_lager[nyckel][0])
                        l_extra.name = namn
                        # Symbologin satts INNAN lagret laggs i kartan - AddLayer lagger in en
                        # kopia, sa andringar pa objektet efterat nar inte kartan
                        if os.path.isfile(lyr_extra):
                            arcpy.ApplySymbologyFromLayer_management(l_extra, lyr_extra)
                        arcpy.mapping.AddLayer(df, l_extra, 'TOP')
                ny_lyr = arcpy.mapping.Layer(ut_fil)
                ny_lyr.name = LAGERNAMN
                if lyr_fil and os.path.isfile(lyr_fil):
                    arcpy.ApplySymbologyFromLayer_management(ny_lyr, lyr_fil)
                    logg('  symbologi applicerad fran %s' % lyr_fil)
                else:
                    logg('')
                    for rad in SYMBOLOGI_TIPS:
                        logg('  ' + rad)
                arcpy.mapping.AddLayer(df, ny_lyr, 'TOP')
                arcpy.RefreshTOC()
                arcpy.RefreshActiveView()
            except Exception as e:
                logg('  kunde inte lagga till lagret: %s' % txt(e))
    elif not (lyr_fil and os.path.isfile(lyr_fil)):
        logg('')
        for rad in SYMBOLOGI_TIPS:
            logg('  ' + rad)

    logg('KLART')
    return ut_fil


# =====================================================================
# Korning med execfile - anvander KONFIG
# =====================================================================

if __name__ == '__main__':
    if BARA_UPPDATERA:
        uppdatera(UT_FC)
    else:
        skapa(JSON_IN, LEDNINGSLAGER, BRUNNSLAGER, BRUNN_ID, UT_FC,
              omradeslager=OMRADESLAGER, csv_ut=CSV_UT, lyr_fil=LYR_FIL,
              tolerans=TOLERANS, marginal=MARGINAL, max_hopp=MAX_HOPP, urval=URVAL,
              kopiera_falt=KOPIERA_FALT, rapportmapp=RAPPORTMAPP, filmmapp=FILMMAPP,
              geojson_ut=GEOJSON_UT, max_delar=MAX_DELAR,
              svackor_ut=SVACKOR_UT, bakfall_ut=BAKFALL_UT, svacka_min_cm=SVACKA_MIN_CM,
              typ_falt=TYP_FALT, dim_falt=DIM_FALT, klipp_avbrutna=KLIPP_AVBRUTNA)
