# -*- coding: utf-8 -*-
"""
Skapar ett ledningslager for bedomda stracker ur tv3_analys.

Kors i ArcMaps Python-fonster med kartan oppen:
    execfile(r'H:\PY\Filmningar\skapa_ledningslager.py')

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

Efter att du fyllt i manuella bedomningar: kor skriptet igen med
BARA_UPPDATERA = True, sa raknas BEDOMNING, BED_TYP och STIL om utan att
geometrin byggs om. Manuella bedomningar bevaras aven vid en full omkorning.
"""
from __future__ import unicode_literals

import os
import re
import io
import sys
import json
import arcpy

# =====================================================================
# KONFIG - det enda du behover andra
# =====================================================================

# Del av lagernamnet i innehallsforteckningen. Flera lager tillatna.
LEDNINGSLAGER = ['A Ledning']
BRUNNSLAGER   = ['A Nedstign och övriga brunnar']

BRUNN_ID = 'EntityID'        # faltet med brunnsbeteckning i brunnslagren

# Valfritt polygonlager att begransa sokningen till. None = hela lagren.
OMRADESLAGER = None          # t.ex. 'paverkansomrade_grovt'

JSON_IN = r'H:\PY\Filmningar\tv3_resultat\kartunderlag.json'
UT_FC   = r'H:\PY\Filmningar\Karta\bedomda_ledningar'       # .gdb-vag eller mapp (= shapefil)
CSV_UT  = r'H:\PY\Filmningar\Karta\omatchade_par.csv'
LYR_FIL = r'H:\PY\Filmningar\Karta\bedomda_ledningar.lyr'   # symbologi, se avsnitt 7

TOLERANS = 2.0     # meter mellan ledningens vertex och brunnen
MARGINAL = 100.0   # meter utanfor omradet dar brunnar anda las in (om OMRADESLAGER anges)
MAX_HOPP = 2       # hur manga brunnar en filmad stracka far passera
URVAL    = 'INTERSECT'     # 'WITHIN' om ledningen maste ligga helt inom omradet

# Falt fran ledningslagret som ska folja med till resultatet.
KOPIERA_FALT = []          # t.ex. ['DIMENSION', 'MATERIAL', 'ANLAGGNINGSAR']

# True = bygg inte om geometrin, rakna bara om BEDOMNING/BED_TYP/STIL i UT_FC
# efter att manuella bedomningar fyllts i.
BARA_UPPDATERA = False

arcpy.env.overwriteOutput = True

KLASSER = ['A', 'B', 'C', 'D', 'E']
KLASSORDNING = dict((k, i) for i, k in enumerate(KLASSER))   # A = varst


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
    print(' '.join(txt(a) for a in args))
    sys.stdout.flush()


# =====================================================================

mxd = arcpy.mapping.MapDocument('CURRENT')
df = arcpy.mapping.ListDataFrames(mxd)[0]


def hitta_lager(namn):
    """Hamtar ett lager ur kartan. Exakt namn i forsta hand, annars delstrang."""
    if os.path.sep in txt(namn) or txt(namn).lower().endswith('.shp'):
        return namn

    alla = [l for l in arcpy.mapping.ListLayers(mxd) if l.isFeatureLayer]
    sokt = txt(namn).strip().lower()

    exakta = [l for l in alla if txt(l.name).strip().lower() == sokt]
    if len(exakta) == 1:
        logg('  "%s" -> %s' % (txt(namn), txt(exakta[0].name)))
        return exakta[0]
    if len(exakta) > 1:
        raise RuntimeError('Flera lager heter exakt "%s"' % txt(namn))

    delvis = [l for l in alla if sokt in txt(l.name).strip().lower()]
    if len(delvis) == 1:
        logg('  "%s" -> %s' % (txt(namn), txt(delvis[0].name)))
        return delvis[0]
    if len(delvis) > 1:
        raise RuntimeError('Flera lager matchar "%s": %s\nAnge exakt namn.'
                           % (txt(namn), ', '.join(txt(l.name) for l in delvis)))
    raise RuntimeError('Hittade inget lager som matchar "%s"' % txt(namn))


def hitta_falt(lyr, faltnamn):
    """Returnerar faltets riktiga namn, oberoende av versaler."""
    falt = arcpy.ListFields(lyr)
    for f in falt:
        if f.name.upper() == txt(faltnamn).upper():
            return f.name
    textfalt = [f.name for f in falt if f.type == 'String']
    raise RuntimeError(
        'Faltet "%s" finns inte i lagret "%s".\nTextfalt som finns: %s'
        % (txt(faltnamn), txt(getattr(lyr, 'name', lyr)), ', '.join(textfalt)))


def bygg_rutnat(brunnar, cell):
    r = {}
    for bid, x, y in brunnar:
        r.setdefault((int(x // cell), int(y // cell)), []).append((bid, x, y))
    return r


def narmaste(rutnat, cell, x, y, tol):
    cx, cy = int(x // cell), int(y // cell)
    bast, bd2 = None, tol * tol
    for dx in (-1, 0, 1):
        for dy in (-1, 0, 1):
            for bid, bx, by in rutnat.get((cx + dx, cy + dy), ()):
                d2 = (bx - x) ** 2 + (by - y) ** 2
                if d2 <= bd2:
                    bast, bd2 = bid, d2
    return bast


def traffar_langs(punkter, rutnat, cell, tol):
    """[(vertexindex, brunns_id)] i ordning langs linjen, utan dubbletter."""
    ut = []
    for i, p in enumerate(punkter):
        b = narmaste(rutnat, cell, p[0], p[1], tol)
        if b and (not ut or ut[-1][1] != b):
            ut.append((i, b))
    return ut


def delstrackor(traffar, bedomda, max_hopp):
    """[(fran_index, till_index, brunn_a, brunn_b)] for bedomda par."""
    ut = []
    for k in range(len(traffar)):
        for h in range(1, max_hopp + 1):
            j = k + h
            if j >= len(traffar):
                break
            a, b = traffar[k][1], traffar[j][1]
            if a == b:
                continue
            if frozenset((a, b)) in bedomda:
                ut.append((traffar[k][0], traffar[j][0], a, b))
                break
    return ut


def galler(mask_bed, man_bed):
    """(BEDOMNING, BED_TYP, STIL) - manuell bedomning gar fore maskinell."""
    m = (txt(man_bed) or '').strip().upper()[:1]
    if m in KLASSORDNING:
        return m, 'Manuell', '%s - Manuell' % m
    k = (txt(mask_bed) or '').strip().upper()[:1]
    if k not in KLASSORDNING:
        k = 'E'
    return k, 'Maskinell', '%s - Maskinell' % k


# ------------------------------------------------ falt i utdata

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
    ('ANT_FILM',   'LONG',   None, 'Antal inspektioner av sträckan'),
    ('SRC_LAGER',  'TEXT',   100, 'Källager'),
    ('SRC_OID',    'LONG',   None, 'Käll-OID'),
]


def rakna_om_bedomning(fc):
    """Uppdaterar BEDOMNING, BED_TYP och STIL utifran MASK_BED och MAN_BED."""
    n = 0
    with arcpy.da.UpdateCursor(fc, ['MASK_BED', 'MAN_BED', 'BEDOMNING', 'BED_TYP', 'STIL']) as mark:
        for rad in mark:
            bed, typ, stil = galler(rad[0], rad[1])
            if [rad[2], rad[3], rad[4]] != [bed, typ, stil]:
                rad[2], rad[3], rad[4] = bed, typ, stil
                mark.updateRow(rad)
                n += 1
    return n


# =====================================================================
# 0. Bara uppdatera befintligt lager
# =====================================================================

if BARA_UPPDATERA:
    mal = UT_FC if arcpy.Exists(UT_FC) else UT_FC + '.shp'
    if not arcpy.Exists(mal):
        raise SystemExit('Hittar inte %s - kor med BARA_UPPDATERA = False forst' % UT_FC)
    n = rakna_om_bedomning(mal)
    logg('%d objekt uppdaterade i %s' % (n, mal))
    manuella = 0
    with arcpy.da.SearchCursor(mal, ['BED_TYP']) as mark:
        for rad in mark:
            if rad[0] == 'Manuell':
                manuella += 1
    logg('  %d objekt har manuell bedomning' % manuella)
    arcpy.RefreshActiveView()
    raise SystemExit('KLART')


# =====================================================================
# 1. Hitta lagren
# =====================================================================

logg('Letar upp lager i kartan')
led_lager = [hitta_lager(n) for n in LEDNINGSLAGER]
brunn_lager = [hitta_lager(n) for n in BRUNNSLAGER]
omrade = hitta_lager(OMRADESLAGER) if OMRADESLAGER else None


# =====================================================================
# 2. Las kartunderlaget fran tv3_analys
# =====================================================================

logg('Laser %s' % JSON_IN)
if not os.path.isfile(JSON_IN):
    raise SystemExit('JSON-filen finns inte. Kor tv3_analys.py forst.')

with io.open(JSON_IN, 'r', encoding='utf-8') as f:
    data = json.loads(f.read())

# Samma brunnspar kan vara filmat flera ganger (t.ex. fran bada hall). Vi behaller
# den varsta bedomningen per par och raknar hur manga inspektioner som finns.
bedomda = {}
antal_per_par = {}
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

logg('  %d stracker i filen, %d unika brunnspar' % (len(data.get('strackor', [])), len(bedomda)))
if not bedomda:
    raise SystemExit('Inga brunnspar kunde lasas ur JSON-filen')


# =====================================================================
# 3. Las in brunnar
# =====================================================================

brunnar = []
for lyr in brunn_lager:
    idfalt = hitta_falt(lyr, BRUNN_ID)
    arcpy.MakeFeatureLayer_management(lyr, 'lyr_br')
    if omrade is not None:
        arcpy.SelectLayerByLocation_management(
            'lyr_br', 'INTERSECT', omrade, '%s Meters' % MARGINAL, 'NEW_SELECTION')
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
    raise SystemExit('Inga brunnar hittades - kontrollera BRUNN_ID och koordinatsystem')

rutnat = bygg_rutnat(brunnar, TOLERANS)
brunns_id = set(b[0] for b in brunnar)


# =====================================================================
# 4. Skapa utdata (manuella bedomningar fran en tidigare korning bevaras)
# =====================================================================

d0 = arcpy.Describe(led_lager[0])
sr = d0.spatialReference
har_z = bool(getattr(d0, 'hasZ', False))

ut_ws = os.path.dirname(UT_FC)
ut_namn = os.path.basename(UT_FC)

ar_shapefil = not ut_ws.lower().endswith('.gdb')
if ar_shapefil and not ut_namn.lower().endswith('.shp'):
    ut_namn = ut_namn + '.shp'
UT_FIL = os.path.join(ut_ws, ut_namn)
MAX_TEXT = 254 if ar_shapefil else 400

if not os.path.isdir(ut_ws) and not arcpy.Exists(ut_ws):
    raise SystemExit('Utdatamappen finns inte: %s' % ut_ws)

# Spara undan manuella bedomningar per brunnspar innan lagret skrivs om
tidigare_manuella = {}
if arcpy.Exists(UT_FIL):
    try:
        with arcpy.da.SearchCursor(UT_FIL, ['FRAN_BRUNN', 'TILL_BRUNN', 'MAN_BED']) as mark:
            for fran, till, man in mark:
                if man and txt(man).strip():
                    tidigare_manuella[frozenset((normalisera(fran), normalisera(till)))] = txt(man).strip()
        if tidigare_manuella:
            logg('  %d manuella bedomningar fran forra korningen bevaras' % len(tidigare_manuella))
    except Exception as e:
        logg('  kunde inte lasa tidigare manuella bedomningar: %s' % txt(e))

if ar_shapefil:
    logg('  utdata blir en shapefil - textfalt kapas till %d tecken' % MAX_TEXT)
    logg('  (skapa en filgeodatabas om du vill ha faltalias som "Maskinell bedomning")')

arcpy.CreateFeatureclass_management(
    ut_ws, ut_namn, 'POLYLINE', '', 'DISABLED',
    'ENABLED' if har_z else 'DISABLED', sr)

for namn, typ, langd, alias in EGNA_FALT:
    if typ == 'TEXT':
        arcpy.AddField_management(UT_FIL, namn, typ,
                                  field_length=min(langd, MAX_TEXT), field_alias=alias)
    else:
        arcpy.AddField_management(UT_FIL, namn, typ, field_alias=alias)

# Vardelista A-E for den manuella bedomningen (bara i filgeodatabas)
if not ar_shapefil:
    try:
        gdb = ut_ws
        if 'TV3_BEDOMNING' not in [d.name for d in arcpy.da.ListDomains(gdb)]:
            arcpy.CreateDomain_management(gdb, 'TV3_BEDOMNING', 'Prioritetsklass A-E',
                                          'TEXT', 'CODED')
            for k in KLASSER:
                arcpy.AddCodedValueToDomain_management(gdb, 'TV3_BEDOMNING', k,
                                                       data.get('klasser', {}).get(k, k))
        arcpy.AssignDomainToField_management(UT_FIL, 'MAN_BED', 'TV3_BEDOMNING')
        logg('  vardelista TV3_BEDOMNING kopplad till MAN_BED')
    except Exception as e:
        logg('  kunde inte skapa vardelista: %s' % txt(e))

TYPKARTA = {'String': 'TEXT', 'Integer': 'LONG', 'SmallInteger': 'SHORT',
            'Double': 'DOUBLE', 'Single': 'FLOAT', 'Date': 'DATE',
            'GUID': 'GUID', 'Blob': 'BLOB'}

kopiera = []
kallfalt = dict((f.name.upper(), f) for f in arcpy.ListFields(led_lager[0]))
for namn in KOPIERA_FALT:
    f = kallfalt.get(txt(namn).upper())
    if not f:
        logg('  VARNING: faltet %s finns inte i ledningslagret - hoppas over' % txt(namn))
        continue
    typ = TYPKARTA.get(f.type)
    if not typ:
        logg('  VARNING: falttypen %s stods inte (%s) - hoppas over' % (f.type, f.name))
        continue
    if typ == 'TEXT':
        arcpy.AddField_management(UT_FIL, f.name, typ,
                                  field_length=min(f.length or 255, MAX_TEXT))
    else:
        arcpy.AddField_management(UT_FIL, f.name, typ)
    kopiera.append(f.name)

ut_falt = ['SHAPE@'] + [n for n, t, l, a in EGNA_FALT] + kopiera


# =====================================================================
# 5. Matcha ledningar mot brunnsparen och skriv objekten
# =====================================================================

def klipp(v, langd):
    return txt(v)[:min(langd, MAX_TEXT)] if v is not None else ''


traffade = set()
n_skrivna = 0
n_lednkoll = 0

insert = arcpy.da.InsertCursor(UT_FIL, ut_falt)
try:
    for lyr in led_lager:
        namn = txt(getattr(lyr, 'name', lyr))
        arcpy.MakeFeatureLayer_management(lyr, 'lyr_led')
        if omrade is not None:
            arcpy.SelectLayerByLocation_management(
                'lyr_led', URVAL, omrade, '', 'NEW_SELECTION')
        antal = int(arcpy.GetCount_management('lyr_led').getOutput(0))
        logg('  %s: %d ledningar att ga igenom' % (namn, antal))

        las_falt = ['OID@', 'SHAPE@'] + kopiera
        with arcpy.da.SearchCursor('lyr_led', las_falt) as mark:
            for rad in mark:
                oid, geom = rad[0], rad[1]
                extra = list(rad[2:])
                n_lednkoll += 1
                if geom is None:
                    continue

                for del_ in geom:
                    punkter = [p for p in del_ if p is not None]
                    if len(punkter) < 2:
                        continue
                    xy = [(p.X, p.Y) for p in punkter]
                    traffar = traffar_langs(xy, rutnat, TOLERANS, TOLERANS)
                    if len(traffar) < 2:
                        continue

                    for i, j, a, b in delstrackor(traffar, bedomda, MAX_HOPP):
                        par = frozenset((a, b))
                        if par in traffade:
                            continue          # skriv varje brunnspar en gang
                        s = bedomda[par]

                        arr = arcpy.Array()
                        for p in punkter[i:j + 1]:
                            arr.add(arcpy.Point(p.X, p.Y, p.Z if har_z else None))
                        ny = arcpy.Polyline(arr, sr, har_z, False)

                        mask = txt(s.get('maskinell_bedomning') or 'E')[:2]
                        man = tidigare_manuella.get(par, '')
                        bed, bed_typ, stil = galler(mask, man)

                        insert.insertRow([
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
                            antal_per_par.get(par, 1),
                            namn[:100], oid,
                        ] + extra)
                        traffade.add(par)
                        n_skrivna += 1

        arcpy.Delete_management('lyr_led')
finally:
    del insert

logg('  %d ledningar genomgangna' % n_lednkoll)
logg('  %d objekt skrivna till %s' % (n_skrivna, UT_FIL))
logg('  %d av %d brunnspar matchade' % (len(traffade), len(bedomda)))
if tidigare_manuella:
    aterstallda = sum(1 for par in traffade if par in tidigare_manuella)
    logg('  %d manuella bedomningar aterstallda' % aterstallda)


# =====================================================================
# 6. Rapport over omatchade par
# =====================================================================

omatchade = []
for par, s in bedomda.items():
    if par in traffade:
        continue
    a, b = normalisera(s.get('startbrunn')), normalisera(s.get('slutbrunn'))
    omatchade.append([a, b,
                      'JA' if a in brunns_id else 'NEJ',
                      'JA' if b in brunns_id else 'NEJ',
                      'JA' if (a in brunns_id or b in brunns_id) else 'NEJ',
                      txt(s.get('maskinell_bedomning')),
                      txt(s.get('fil'))])

omatchade.sort(key=lambda r: (r[4] != 'JA', KLASSORDNING.get(r[5], 9), r[0]))

with io.open(CSV_UT, 'w', encoding='cp1252', errors='replace') as f:
    f.write('fran;till;fran_finns;till_finns;nagon_finns;maskinell_bedomning;kallfil\n')
    for r in omatchade:
        f.write(';'.join(txt(v).replace(';', ',') for v in r) + '\n')

n_intressanta = sum(1 for r in omatchade if r[4] == 'JA')
logg('  %d omatchade par -> %s' % (len(omatchade), CSV_UT))
logg('  varav %d har minst en brunn i kartan (dessa ar vard att granska)' % n_intressanta)


# =====================================================================
# 7. Lagg till i kartan
# =====================================================================

try:
    ny_lyr = arcpy.mapping.Layer(UT_FIL)
    ny_lyr.name = 'Bedomda ledningar'
    arcpy.mapping.AddLayer(df, ny_lyr, 'TOP')
    if os.path.isfile(LYR_FIL):
        arcpy.ApplySymbologyFromLayer_management(ny_lyr, LYR_FIL)
        logg('  symbologi applicerad fran %s' % LYR_FIL)
    else:
        logg('')
        logg('  Ingen .lyr-fil an. Satt symbologin en gang:')
        logg('    Egenskaper > Symbology > Categories > Unique values, Value Field: STIL')
        logg('    Add All Values ger tio kategorier:')
        logg('      A/B/C/D/E - Maskinell  ->  STRECKAD linje i klassens farg')
        logg('      A/B/C/D/E - Manuell    ->  HELDRAGEN linje i klassens farg')
        logg('    Fargar: A rott, B orange, C gult, D gront, E gratt.')
        logg('    Spara sedan lagret som %s - nasta korning applicerar det automatiskt.' % LYR_FIL)
    arcpy.RefreshTOC()
    arcpy.RefreshActiveView()
except Exception as e:
    logg('  kunde inte lagga till lagret: %s' % txt(e))

logg('KLART')
