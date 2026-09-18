# -*- coding: utf-8 -*-
"""
Skapar ett ledningslager for bedomda stracker ur tv3_analys.

Tva satt att kora:

  1. Som verktyg i ArcMap: lagg till arcmap/tv3_verktyg.pyt i ArcToolbox och kor
     "Skapa ledningslager" - da valjer du JSON-fil, lager och utdata i en dialog.

  2. I ArcMaps Python-fonster med kartan oppen, med installningarna i KONFIG nedan:
         execfile(r'H:\PY\tv3analys\arcmap\skapa_ledningslager.py')

Indata ar kartunderlag.json fran tv3_analys.py. For varje stracka i JSON-filen
(ett brunnspar) letas ledningen mellan brunnarna upp: alla vertexpunkter i
ledningslagret gas igenom, narmaste brunn inom TOLERANS soks upp, och det
stycke som ligger mellan de tva brunnarna klipps ut som ett eget objekt.
Detta hanterar complex edges som passerar flera brunnar utan att vara fysiskt
uppdelade.

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
LYR_FIL = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'bedomda_ledningar.lyr')

TOLERANS = 2.0     # meter mellan ledningens vertex och brunnen
MARGINAL = 100.0   # meter utanfor omradet dar brunnar anda las in (om OMRADESLAGER anges)
MAX_HOPP = 2       # hur manga brunnar en filmad stracka far passera
URVAL    = 'INTERSECT'     # 'WITHIN' om ledningen maste ligga helt inom omradet

# Falt fran ledningslagret som ska folja med till resultatet.
KOPIERA_FALT = []          # t.ex. ['DIMENSION', 'MATERIAL', 'ANLAGGNINGSAR']

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
        for enc in ('utf-8', 'cp1252', 'latin-1'):
            try:
                return v.decode(enc)
            except UnicodeDecodeError:
                continue
        return v.decode('latin-1', 'replace')
    return TEXTTYP(v)


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
    print(rad)
    sys.stdout.flush()


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

    def vag(self, a, b, max_hopp, max_delar=8):
        """Vagen med farst bitar fran brunn a till brunn b, eller None. Hogst
        max_hopp - 1 andra sokta brunnar far passeras, och hogst max_delar bitar."""
        from collections import deque
        start, mal = ('B', a), ('B', b)
        if start not in self.kanter or mal not in self.kanter:
            return None
        ko = deque([(start, [], 0)])
        sedda = set([start])
        while ko:
            nod, vagen, hopp = ko.popleft()
            if len(vagen) >= max_delar:
                continue
            for annan, pts, lager, oid in self.kanter.get(nod, ()):
                if annan == mal:
                    return vagen + [(annan, pts, lager, oid)]
                if annan in sedda:
                    continue
                h = hopp
                if annan[0] == 'B':
                    h += 1
                    if h > max_hopp - 1:
                        continue
                sedda.add(annan)
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

    def diagnos(self, a, b, max_hopp):
        """Varfor hittades ingen vag mellan a och b? Returnerar en forklaring."""
        start, mal = ('B', a), ('B', b)
        if start not in self.kanter or mal not in self.kanter:
            return 'brunnen ligger inte pa nagon ledning'
        v = self.vag(a, b, 999, 200)
        if v:
            n_br = sum(1 for nod, pts, lager, oid in v[:-1] if nod[0] == 'B')
            return ('vag finns via %d bitar och %d andra brunnar - hoj max hopp till %d'
                    % (len(v), n_br, n_br + 1))
        ka = self.komponent(start)
        kb = self.komponent(mal)
        if ka == kb:
            return 'samma natverk men vagen ar orimligt lang'
        # Narmaste avstand mellan de tva natverksdelarna (rutnat over den mindre delen)
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
    ('SKADOR',     'TEXT',   200, 'Skador (kod+grad)'),
    ('DRIFTATG',   'TEXT',   100, 'Driftåtgärd'),
    ('LANGD_M',    'DOUBLE', None, 'Längd (m)'),
    ('MATERIAL',   'TEXT',   50,  'Material'),
    ('DIMENSION',  'TEXT',   20,  'Dimension (mm)'),
    ('LEDN_TYP',   'TEXT',   30,  'Ledningstyp'),
    ('RELINAD',    'TEXT',   3,   'Relinad'),
    ('AVBRUTEN',   'TEXT',   3,   'Avbruten inspektion'),
    ('SVACKA_M',   'DOUBLE', None, 'Svackdjup (m)'),
    ('LUTNING',    'DOUBLE', None, 'Lutning (promille)'),
    ('OMRADE',     'TEXT',   60,  'Område'),
    ('DATUM',      'TEXT',   10,  'Inspektionsdatum'),
    ('TV3_FIL',    'TEXT',   100, 'TV3-fil'),
    ('RAPPORT',    'TEXT',   254, 'Inspektionsprotokoll (PDF)'),
    ('VIDEO',      'TEXT',   254, 'Videofil'),
    ('ANT_FILM',   'LONG',   None, 'Antal inspektioner av sträckan'),
    ('ANT_DELAR',  'LONG',   None, 'Antal ledningsobjekt i kartan'),
    ('SRC_LAGER',  'TEXT',   100, 'Källager'),
    ('SRC_OID',    'LONG',   None, 'Käll-OID'),
]


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


def las_kartunderlag(json_in):
    """Laser JSON-filen. Returnerar (data, bedomda, antal_per_par) dar bedomda ar
    {frozenset(brunnspar): post} med den varsta bedomningen per par."""
    if not os.path.isfile(json_in):
        raise RuntimeError('JSON-filen finns inte: %s\nKor tv3_analys.py forst.' % json_in)
    with io.open(json_in, 'r', encoding='utf-8') as f:
        data = json.loads(f.read())

    bedomda, antal_per_par = {}, {}
    for post in data.get('strackor', []):
        a, b = normalisera(post.get('startbrunn')), normalisera(post.get('slutbrunn'))
        if not a or not b or a == b:
            continue
        par = frozenset((a, b))
        antal_per_par[par] = antal_per_par.get(par, 0) + 1
        tidigare = bedomda.get(par)
        if tidigare is None or (KLASSORDNING.get(txt(post.get('maskinell_bedomning')), 9)
                                < KLASSORDNING.get(txt(tidigare.get('maskinell_bedomning')), 9)):
            bedomda[par] = post
    return data, bedomda, antal_per_par


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
          kopiera_falt=None, lagg_till_i_kartan=True):
    """Bygger ledningslagret. Returnerar sokvagen till den skrivna featureklassen."""
    kopiera_falt = kopiera_falt or []
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
    brunnar = []
    for lyr in brunn_lager:
        idfalt = hitta_falt(lyr, brunn_id)
        src, dq = kalla(lyr)
        arcpy.MakeFeatureLayer_management(src, 'lyr_br', dq)
        if omrade is not None:
            arcpy.SelectLayerByLocation_management(
                'lyr_br', 'INTERSECT', omrade, '%s Meters' % marginal, 'NEW_SELECTION')
        n = 0
        with arcpy.da.SearchCursor('lyr_br', [idfalt, 'SHAPE@XY']) as mark:
            for bid, xy in mark:
                nid = normalisera(bid)
                if nid and xy and xy[0] is not None:
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
    for par in bedomda:
        json_brunnar.update(par)
    saknade = sorted(json_brunnar - brunns_id)
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

    # Bara JSON-filens brunnar behovs for matchningen (forsta forekomsten vinner)
    sokta = {}
    for bid, x, y in brunnar:
        if bid in json_brunnar and bid not in sokta:
            sokta[bid] = (x, y)

    # ---------------------------------------------------- 4. Utdata
    d0 = arcpy.Describe(kalla(led_lager[0])[0])
    sr = d0.spatialReference
    har_z = bool(getattr(d0, 'hasZ', False))

    ut_ws = os.path.dirname(ut_fc)
    ut_namn = os.path.basename(ut_fc)
    ar_shapefil = not ut_ws.lower().endswith('.gdb')
    if ar_shapefil and not ut_namn.lower().endswith('.shp'):
        ut_namn = ut_namn + '.shp'
    ut_fil = os.path.join(ut_ws, ut_namn)
    max_text = 254 if ar_shapefil else 400

    if not os.path.isdir(ut_ws) and not arcpy.Exists(ut_ws):
        raise RuntimeError('Utdatamappen finns inte: %s' % ut_ws)

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
            if 'TV3_BEDOMNING' not in [d.name for d in arcpy.da.ListDomains(ut_ws)]:
                arcpy.CreateDomain_management(ut_ws, 'TV3_BEDOMNING', 'Prioritetsklass A-E',
                                              'TEXT', 'CODED')
                for k in KLASSER:
                    arcpy.AddCodedValueToDomain_management(
                        ut_ws, 'TV3_BEDOMNING', k, data.get('klasser', {}).get(k, k))
            arcpy.AssignDomainToField_management(ut_fil, 'MAN_BED', 'TV3_BEDOMNING')
            logg('  vardelista TV3_BEDOMNING kopplad till MAN_BED')
        except Exception as e:
            logg('  kunde inte skapa vardelista: %s' % txt(e))

    typkarta = {'String': 'TEXT', 'Integer': 'LONG', 'SmallInteger': 'SHORT',
                'Double': 'DOUBLE', 'Single': 'FLOAT', 'Date': 'DATE',
                'GUID': 'GUID', 'Blob': 'BLOB'}
    kopiera = []
    kallfalt = dict((f.name.upper(), f) for f in arcpy.ListFields(kalla(led_lager[0])[0]))
    for namn in kopiera_falt:
        f = kallfalt.get(txt(namn).upper())
        if not f:
            logg('  VARNING: faltet %s finns inte i ledningslagret - hoppas over' % txt(namn))
            continue
        typ = typkarta.get(f.type)
        if not typ:
            logg('  VARNING: falttypen %s stods inte (%s) - hoppas over' % (f.type, f.name))
            continue
        if typ == 'TEXT':
            arcpy.AddField_management(ut_fil, f.name, typ,
                                      field_length=min(f.length or 255, max_text))
        else:
            arcpy.AddField_management(ut_fil, f.name, typ)
        kopiera.append(f.name)

    ut_falt = ['SHAPE@'] + [n for n, t, l, a in EGNA_FALT] + kopiera
    ut_typer = ['SHAPE@'] + [t for n, t, l, a in EGNA_FALT] + \
               [typkarta.get(kallfalt[k.upper()].type, 'TEXT') for k in kopiera]

    def klipp(v, langd):
        return txt(v)[:min(langd, max_text)] if v is not None else ''

    def utan_null(rad):
        """Shapefiler kan inte lagra NULL i tal- och textfalt: tomma varden blir
        0 respektive '' (samma sak ArcGIS gor vid export till shapefil)."""
        if not ar_shapefil:
            return rad
        ut = []
        for v, typ in zip(rad, ut_typer):
            if v is None:
                v = '' if typ == 'TEXT' else 0
            ut.append(v)
        return ut

    if ar_shapefil:
        logg('  tomma tal (t.ex. lutning utan profil) skrivs som 0 i shapefil')

    # ---------------------------------------------------- 5. Bygg natverk av ledningsbitar
    nat = Natverk(sokta, tolerans)
    extra_per_oid = {}
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

        with arcpy.da.SearchCursor('lyr_led', ['OID@', 'SHAPE@'] + kopiera) as mark:
            for rad in mark:
                oid, geom = rad[0], rad[1]
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
                extra_per_oid[(namn, oid)] = list(rad[2:])
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

    def rapport_sokvag(post):
        """Absolut sokvag till strackans PDF (relativ till utdatamappen i JSON-filen)."""
        r = post.get('rapport')
        if not r:
            return ''
        r = txt(r).replace('/', os.sep)
        return r if os.path.isabs(r) else os.path.join(utdata_mapp, r)

    traffade = set()
    n_skrivna = n_flerdelade = 0

    insert = arcpy.da.InsertCursor(ut_fil, ut_falt)
    try:
        for par, s in bedomda.items():
            a, b = normalisera(s.get('startbrunn')), normalisera(s.get('slutbrunn'))
            vagen = nat.vag(a, b, max_hopp)
            if not vagen:
                continue
            pts = sla_ihop(vagen)
            if len(pts) < 2:
                continue

            arr = arcpy.Array()
            for x, y, z in pts:
                arr.add(arcpy.Point(x, y, z if har_z else None))
            ny = arcpy.Polyline(arr, sr, har_z, False)

            mask = txt(s.get('maskinell_bedomning') or 'E')[:2]
            man = tidigare_manuella.get(par, '')
            bed, bed_typ, stil = galler(mask, man)
            lager0, oid0 = vagen[0][2], vagen[0][3]
            extra = extra_per_oid.get((lager0, oid0), [None] * len(kopiera))

            insert.insertRow(utan_null([
                ny, mask, man, bed, bed_typ, stil,
                klipp(s.get('startbrunn'), 50), klipp(s.get('slutbrunn'), 50),
                klipp(s.get('klasstext'), 40),
                s.get('totalindex'), s.get('konstruktionsindex'),
                s.get('maxgrad_konstruktion'), s.get('antal_skador'),
                klipp(s.get('skador'), 200), klipp(s.get('driftatgard'), 100),
                s.get('langd_m'), klipp(s.get('material'), 50),
                klipp(s.get('dimension'), 20), klipp(s.get('ledningstyp'), 30),
                'Ja' if s.get('relinad') else 'Nej',
                'Ja' if s.get('avbruten') else 'Nej',
                s.get('svackdjup_m'), s.get('lutning_promille'),
                klipp(s.get('omrade'), 60), klipp(s.get('datum'), 10),
                klipp(os.path.basename(txt(s.get('tv3_fil') or '')), 100),
                klipp(rapport_sokvag(s), 254), klipp(s.get('video_sokvag') or '', 254),
                antal_per_par.get(par, 1), len(vagen),
                lager0[:100], oid0,
            ] + extra))
            traffade.add(par)
            n_skrivna += 1
            if len(vagen) > 1:
                n_flerdelade += 1
    finally:
        del insert

    logg('  %d objekt skrivna till %s' % (n_skrivna, ut_fil))
    logg('  %d av %d brunnspar matchade (%d sammansatta av flera ledningsobjekt)'
         % (len(traffade), len(bedomda), n_flerdelade))
    if tidigare_manuella:
        logg('  %d manuella bedomningar aterstallda'
             % sum(1 for par in traffade if par in tidigare_manuella))

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
        if a in brunns_id and b in brunns_id:
            diagnos = nat.diagnos(a, b, max_hopp)
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
                          txt(s.get('fil')), diagnos])
    omatchade.sort(key=lambda r: (r[4] != 'JA', KLASSORDNING.get(r[7], 9), r[0]))

    if csv_ut:
        with io.open(csv_ut, 'w', encoding='cp1252', errors='replace') as f:
            f.write('fran;till;fran_finns;till_finns;nagon_finns;'
                    'fran_avstand_m;till_avstand_m;maskinell_bedomning;kallfil;diagnos\n')
            for r in omatchade:
                f.write(';'.join(txt(v).replace(';', ',') for v in r) + '\n')
        logg('  %d omatchade par -> %s' % (len(omatchade), csv_ut))
    else:
        logg('  %d omatchade par' % len(omatchade))

    # Varfor? Bada brunnarna pa en ledning men ingen vag = ledningen ligger i ett annat
    # lager, ar bruten, eller passerar fler brunnar an max_hopp. En brunn en bit fran
    # ledningen = hoj toleransen.
    def _tal(v):
        try:
            return float(v)
        except (TypeError, ValueError):
            return None
    bada_brunnar = [r for r in omatchade if r[2] == 'JA' and r[3] == 'JA']
    pa_ledning = [r for r in bada_brunnar
                  if _tal(r[5]) is not None and _tal(r[5]) <= tolerans
                  and _tal(r[6]) is not None and _tal(r[6]) <= tolerans]
    nara = [r for r in bada_brunnar if r not in pa_ledning
            and _tal(r[5]) is not None and _tal(r[6]) is not None
            and max(_tal(r[5]), _tal(r[6])) <= 3 * tolerans]
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

    # ---------------------------------------------------- 8. Karta
    if lagg_till_i_kartan:
        mxd = _mxd()
        if mxd is not None:
            try:
                df = arcpy.mapping.ListDataFrames(mxd)[0]
                ny_lyr = arcpy.mapping.Layer(ut_fil)
                ny_lyr.name = LAGERNAMN
                arcpy.mapping.AddLayer(df, ny_lyr, 'TOP')
                if lyr_fil and os.path.isfile(lyr_fil):
                    arcpy.ApplySymbologyFromLayer_management(ny_lyr, lyr_fil)
                    logg('  symbologi applicerad fran %s' % lyr_fil)
                else:
                    logg('')
                    for rad in SYMBOLOGI_TIPS:
                        logg('  ' + rad)
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
              kopiera_falt=KOPIERA_FALT)
