# -*- coding: utf-8 -*-
"""
Exportera omrade: kopierar alla VA-objekt som beror ett eller flera omraden (polygoner i ett
lager i kartan) till en lokal filgeodatabas, ett featureklass per kallager med alla falt
(ArcMap 10.x, Python 2.7). Steg 1 i PLAN_ledningsinventering.md.

Hela objekt kopieras (inte klippta) sa att langder, vattengangar och brunnspar forblir riktiga -
utom vattenledningarna (V-lager eller gruppen Vatten), som klipps vid omradesgransen (+ buffert) nar
klipp_vatten ar satt: de kan annars stracka sig hundratals meter utanfor; KLIPPT_M anger hur mycket av
objektet som togs bort. Avloppet klipps aldrig (brunn till brunn). Varje objekt far OMRADE (polygonens namn, flera med ';'), KORSAR_GRANS (1 nar objektet sticker
ut ur omradet) och KALLAGER (kallagrets namn). Med buffert > 0 tas aven objekt strax utanfor med
sa att natet hanger ihop vid gransen. export_info.json i mappen beskriver exporten.

Kors som verktyget "Exportera omrade" i tv3_verktyg.pyt, eller i ArcMaps Python-fonster:

    import sys; sys.path.insert(0, r'C:\\...\\arcmap\\moduler')
    import omrade_export
    omrade_export.exportera_omrade('Omraden', ['Avlopp\\A Ledning', 'Avlopp\\A Servis'],
                                   r'C:\\Temp\\export', 'alsten', namnfalt='NAMN')

Testas utan ArcMap med en latsas-arcpy (scratchpad/test_omrade.py).
"""
from __future__ import unicode_literals

import os
import io
import re
import json
import datetime

import arcpy

from skapa_ledningslager import txt, logg, hitta_lager, kalla, TEXTTYP, _mxd

VERSION = 1
FALT_OMRADE = 'OMRADE'          # polygonens namn (flera med ';')
FALT_KORSAR = 'KORSAR_GRANS'    # 1 = objektet ligger delvis utanfor omradet
FALT_KALLA = 'KALLAGER'         # kallagrets namn i kartan
FALT_KLIPPT = 'KLIPPT_M'        # vattenledningar: meter av objektet utanfor omradet som klipptes bort
KLIPP_VATTEN = True             # vattenledningar klipps vid omradesgransen (+ buffert)
INFOFIL = 'export_info.json'
LAGERSUFFIX = ' (lokal)'        # namn i kartan pa de lokala lagren
TMP_OMRADE = 'in_memory/tv3_omrade'
TMP_LAGER = 'tv3_omrade_export_tmp'
ANDBRUNN_AVSTAND_M = 2.0        # avloppets punktlager (brunnar, ledningsandar) inom detta fran en exporterad
                                # avloppsledning tas med aven utanfor omradet - ledningen ska ga brunn till
                                # brunn aven nar den korsar gransen; None stanger av

_AO = {'\u00e5': 'a', '\u00e4': 'a', '\u00f6': 'o', '\u00c5': 'A', '\u00c4': 'A', '\u00d6': 'O',
       '\u00e9': 'e', '\u00fc': 'u'}


def _fc_namn(namn, tagna):
    """Giltigt namn pa en featureklass i en filgeodatabas: bokstaver/siffror/understreck, borjar
    med bokstav, hogst 60 tecken, unikt bland `tagna` (lopnummer vid krock)."""
    n = txt(namn)
    for k, v in _AO.items():
        n = n.replace(k, v)
    n = re.sub(r'[^A-Za-z0-9_]+', '_', n).strip('_')
    if not n:
        n = 'lager'
    if not re.match(r'[A-Za-z]', n):
        n = 'L_' + n
    n = n[:60]
    bas, i = n, 1
    while n.lower() in tagna:
        i += 1
        n = '%s_%d' % (bas[:57], i)
    tagna.add(n.lower())
    return n


def _lagernamn(lyr):
    return txt(getattr(lyr, 'name', None) or os.path.basename(txt(lyr)))


def _har_urval(lyr):
    try:
        fids = arcpy.Describe(lyr).FIDSet
        return bool(fids and txt(fids).strip())
    except Exception:
        return False


def _polygoner(lyr, namnfalt=None, bara_valda=True):
    """[(namn, geometri)] ur omradeslagret. Finns ett urval i kartan och bara_valda ar satt anvands
    bara de valda polygonerna (cursor pa lagerobjektet), annars alla (datakallan med lagrets
    definitionsfraga)."""
    src, dq = kalla(lyr)
    anvand_urval = bara_valda and not isinstance(lyr, (TEXTTYP, bytes)) and _har_urval(lyr)
    kalla_cursor = lyr if anvand_urval else src
    falt = ['SHAPE@'] + ([namnfalt] if namnfalt else [])
    ut = []
    with arcpy.da.SearchCursor(kalla_cursor, falt, None if anvand_urval else dq) as cur:
        for i, rad in enumerate(cur):
            geom = rad[0]
            if geom is None:
                continue
            namn = txt(rad[1]).strip() if namnfalt and rad[1] is not None else ''
            ut.append((namn or 'Omrade %d' % (i + 1), geom))
    if anvand_urval:
        logg('  %d valda polygoner i omradeslagret anvands' % len(ut))
    else:
        logg('  %d polygoner i omradeslagret' % len(ut))
    return ut


def _ar_avlopp(lyr):
    """Avloppslager: i gruppen Avlopp eller med namn 'A ...' (SVOA)."""
    kort = txt(getattr(lyr, 'name', '') or '').strip().lower()
    langt = txt(getattr(lyr, 'longName', '') or '').strip().lower()
    return kort.startswith('a ') or langt.startswith('avlopp')


def _ar_vatten(lyr):
    """Vattenlager: i gruppen Vatten eller med namn 'V ...' (SVOA)."""
    kort = txt(getattr(lyr, 'name', '') or '').strip().lower()
    langt = txt(getattr(lyr, 'longName', '') or '').strip().lower()
    return kort.startswith('v ') or langt.startswith('vatten')


def _geometrityp(namn):
    try:
        return txt(getattr(arcpy.Describe(kalla(hitta_lager(namn))[0]), 'shapeType', '') or '')
    except Exception:
        return ''


def _union(geometrier):
    u = None
    for g in geometrier:
        u = g if u is None else u.union(g)
    return u


def exportera_omrade(omradeslager, lager, ut_mapp, gdb_namn='omrade_export', namnfalt=None,
                     buffert=0.0, bara_valda=True, lagg_i_kartan=True, klipp_vatten=None):
    """Kopierar objekten i `lager` (lista med kartans lagernamn) som beror polygonerna i
    `omradeslager` till <ut_mapp>/<gdb_namn>.gdb. Returnerar en sammanfattning (dict), som ocksa
    skrivs till export_info.json i mappen."""
    if isinstance(lager, (TEXTTYP, bytes)):
        lager = [lager]
    lager = [l for l in lager if txt(l).strip()]
    if not lager:
        raise RuntimeError('Inga lager att exportera')
    buffert = float(buffert or 0)
    if klipp_vatten is None:
        klipp_vatten = KLIPP_VATTEN
    ut_mapp = txt(ut_mapp)
    if not os.path.isdir(ut_mapp):
        os.makedirs(ut_mapp)
    gdb_namn = txt(gdb_namn or 'omrade_export')
    if gdb_namn.lower().endswith('.gdb'):
        gdb_namn = gdb_namn[:-4]
    if not re.match(r'^[A-Za-z][A-Za-z0-9_\-]*$', gdb_namn):
        raise RuntimeError('Geodatabasens namn far bara innehalla A-Z, siffror, understreck och '
                           'bindestreck och ska borja med en bokstav: %s' % gdb_namn)
    gdb = os.path.join(ut_mapp, gdb_namn + '.gdb')

    logg('Letar upp omradeslagret')
    omr = hitta_lager(omradeslager)
    polys = _polygoner(omr, namnfalt, bara_valda)
    if not polys:
        raise RuntimeError('Inga polygoner i omradeslagret (ar nagot valt i ett tomt urval?)')
    union = _union([g for _, g in polys])
    sok = union.buffer(buffert) if buffert > 0 else union
    if buffert > 0:
        logg('  buffert %.1f m runt omradena' % buffert)

    if not arcpy.Exists(gdb):
        arcpy.CreateFileGDB_management(ut_mapp, gdb_namn + '.gdb')
        logg('Skapade %s' % gdb)
    else:
        logg('Skriver till %s (befintliga lager med samma namn skrivs over)' % gdb)

    # Sokpolygonen som en tillfallig featureklass - SelectLayerByLocation tar inte geometriobjekt
    # i alla versioner
    try:
        arcpy.Delete_management(TMP_OMRADE)
    except Exception:
        pass
    arcpy.CopyFeatures_management([sok], TMP_OMRADE)

    tagna = set()
    resultat = []
    hoppade = []
    # linjelagren forst, sa att avloppsbrunnarna vid ledningarnas andar kan hamtas efterat
    lager = sorted(lager, key=lambda n: 0 if _geometrityp(n) == 'Polyline' else 1)
    avlopp_linjer = []
    sr_namn = ''
    mxd = _mxd() if lagg_i_kartan else None
    df = None
    if mxd is not None:
        try:
            df = arcpy.mapping.ListDataFrames(mxd)[0]
        except Exception:
            df = None
    try:
        for namn in lager:
            try:
                lyr = hitta_lager(namn)
                src, dq = kalla(lyr)
                try:
                    d = arcpy.Describe(src)
                    geometri = txt(getattr(d, 'shapeType', '') or '')
                    if not sr_namn:
                        sr_namn = txt(getattr(getattr(d, 'spatialReference', None), 'name', '') or '')
                except Exception as e:
                    logg('  hoppar "%s": kan inte beskrivas (%s)' % (txt(namn), txt(e)))
                    continue
                if not geometri:
                    logg('  hoppar "%s": inget featurelager' % txt(namn))
                    continue
                lagernamn = _lagernamn(lyr)
                fc_namn = _fc_namn(lagernamn, tagna)
                ut_fc = os.path.join(gdb, fc_namn)
                try:
                    arcpy.Delete_management(TMP_LAGER)
                except Exception:
                    pass
                _tillfalligt_lager(lyr, src, dq, lagernamn)
                arcpy.SelectLayerByLocation_management(TMP_LAGER, 'INTERSECT', TMP_OMRADE)
                antal = int(arcpy.GetCount_management(TMP_LAGER).getOutput(0))
                if ANDBRUNN_AVSTAND_M and geometri == 'Point' and avlopp_linjer and _ar_avlopp(lyr):
                    for fc in avlopp_linjer:
                        arcpy.SelectLayerByLocation_management(TMP_LAGER, 'WITHIN_A_DISTANCE', fc,
                                                               '%s Meters' % ANDBRUNN_AVSTAND_M, 'ADD_TO_SELECTION')
                    fler = int(arcpy.GetCount_management(TMP_LAGER).getOutput(0)) - antal
                    if fler:
                        logg('  %s: %d utanfor omradet vid avloppsledningar som korsar gransen tas med'
                             % (lagernamn, fler))
                    antal += fler
                if antal == 0:
                    logg('  %s: inga objekt inom omradet' % lagernamn)
                    arcpy.Delete_management(TMP_LAGER)
                    resultat.append({'lager': lagernamn, 'langt_namn': txt(getattr(lyr, 'longName', lagernamn)),
                                     'kalla': txt(src), 'fc': None, 'antal': 0, 'korsar_grans': 0,
                                     'geometri': geometri})
                    continue
                if arcpy.Exists(ut_fc):
                    _ta_bort_ur_kartan(mxd, ut_fc)
                    arcpy.Delete_management(ut_fc)
                arcpy.CopyFeatures_management(TMP_LAGER, ut_fc)
                arcpy.Delete_management(TMP_LAGER)

                # Omrade, grans och kallager per objekt
                befintliga = set(f.name.upper() for f in arcpy.ListFields(ut_fc))
                if FALT_OMRADE not in befintliga:
                    arcpy.AddField_management(ut_fc, FALT_OMRADE, 'TEXT', field_length=254, field_alias='Omrade')
                if FALT_KORSAR not in befintliga:
                    arcpy.AddField_management(ut_fc, FALT_KORSAR, 'SHORT', field_alias='Korsar omradesgrans')
                if FALT_KALLA not in befintliga:
                    arcpy.AddField_management(ut_fc, FALT_KALLA, 'TEXT', field_length=100, field_alias='Kallager')
                klipp = sok if (klipp_vatten and geometri == 'Polyline' and _ar_vatten(lyr)) else None
                if klipp is not None and FALT_KLIPPT not in befintliga:
                    arcpy.AddField_management(ut_fc, FALT_KLIPPT, 'DOUBLE', field_alias='Klippt vid omradesgransen (m)')
                korsar, klippta = _fyll_omradesfalt(ut_fc, gdb, polys, union, lagernamn, klipp)
                if klipp is not None:
                    antal -= klippta['bort']
                    logg('  %s: %d vattenledningar klippta vid omradesgransen (%.0f m utanfor togs bort)%s'
                         % (lagernamn, klippta['klippta'], klippta['meter'],
                            (', %d helt utanfor togs bort' % klippta['bort']) if klippta['bort'] else ''))
                if geometri == 'Polyline' and _ar_avlopp(lyr):
                    avlopp_linjer.append(ut_fc)
                logg('  %s: %d objekt -> %s%s' % (lagernamn, antal, fc_namn,
                                                   (' (%d korsar omradesgransen)' % korsar) if korsar else ''))
                post = {'lager': lagernamn, 'langt_namn': txt(getattr(lyr, 'longName', lagernamn)),
                        'kalla': txt(src), 'fc': ut_fc, 'antal': antal, 'korsar_grans': korsar,
                        'geometri': geometri}
                if klipp is not None:
                    post['klippt'] = {'ledningar': klippta['klippta'], 'meter': round(klippta['meter'], 1),
                                      'bort': klippta['bort']}
                resultat.append(post)
                if df is not None:
                    _lagg_in(df, ut_fc, lyr, lagernamn)
            except Exception as e:                      # ett trasigt lager ska inte stoppa exporten
                logg('  OBS hoppar over "%s": %s' % (txt(namn), txt(e)))
                hoppade.append(txt(namn))
                resultat.append({'lager': txt(namn), 'langt_namn': txt(namn), 'kalla': '', 'fc': None,
                                 'antal': 0, 'korsar_grans': 0, 'geometri': '', 'fel': txt(e)})
                try:
                    arcpy.Delete_management(TMP_LAGER)
                except Exception:
                    pass
    finally:
        try:
            arcpy.Delete_management(TMP_OMRADE)
        except Exception:
            pass

    info = {
        'version': VERSION,
        'datum': datetime.datetime.now().strftime('%Y-%m-%d %H:%M'),
        'omradeslager': txt(getattr(omr, 'longName', None) or getattr(omr, 'name', None) or omradeslager),
        'omraden': [n for n, _ in polys],
        'buffert_m': buffert,
        'vatten_klippt': bool(klipp_vatten),
        'gdb': gdb,
        'koordinatsystem': sr_namn,
        'lager': resultat,
    }
    infofil = os.path.join(ut_mapp, INFOFIL)
    with io.open(infofil, 'w', encoding='utf-8') as f:
        f.write(json.dumps(info, ensure_ascii=False, indent=1))
    n_tot = sum(r['antal'] for r in resultat)
    logg('%d objekt i %d lager exporterade till %s' % (n_tot, sum(1 for r in resultat if r['fc']), gdb))
    logg('  export_info.json skriven i %s' % ut_mapp)
    if hoppade:
        logg('OBS %d lager kunde inte exporteras (se ovan): %s' % (len(hoppade), ', '.join(hoppade)))
    if mxd is not None and df is not None:
        try:
            arcpy.RefreshTOC()
            arcpy.RefreshActiveView()
        except Exception:
            pass
    return info


def _utan_tabellprefix(dq):
    """sde_geopipe.GNG.pipeSewerWellP.Status -> Status (falt i en kopplad tabell, join). Text inom
    apostrofer ('Gata.Nord') lamnas orord."""
    delar = re.split(r"('(?:[^']|'')*')", dq)
    return ''.join(d if d.startswith("'") else re.sub(r"\b(?:[A-Za-z_]\w*\.)+([A-Za-z_]\w*)\b", r"\1", d)
                   for d in delar)


def _tillfalligt_lager(lyr, src, dq, lagernamn):
    """MakeFeatureLayer for exporten. Lager med en koppling (join) i kartan har en definitionsfraga med
    tabellprefix (sde_geopipe.GNG.pipeSewerWellP.Status ...) som inte galler datakallan ensam (ERROR
    000358). Da provas fragan utan prefix och sist lagret i kartan sjalvt (med kopplingen)."""
    forsok = [(src, dq, '')]
    if dq and _utan_tabellprefix(dq) != dq:
        forsok.append((src, _utan_tabellprefix(dq), 'definitionsfragan utan tabellprefix'))
    if lyr is not None and not isinstance(lyr, (TEXTTYP, bytes)):
        forsok.append((lyr, None, 'lagret i kartan med sin koppling (join)'))
    fel = []
    for kalla_, fraga, hur in forsok:
        try:
            arcpy.Delete_management(TMP_LAGER)
        except Exception:
            pass
        urval = None
        if kalla_ is lyr:
            # ett lager i kartan med urval ger bara de valda objekten - urvalet tas bort under tiden
            try:
                urval = list(lyr.getSelectionSet() or [])
                if urval:
                    lyr.setSelectionSet('NEW', [])
            except Exception:
                urval = None
        try:
            if fraga:
                arcpy.MakeFeatureLayer_management(kalla_, TMP_LAGER, fraga)
            else:
                arcpy.MakeFeatureLayer_management(kalla_, TMP_LAGER)
        except Exception as e:
            fel.append(txt(e).strip().split('\n')[0])
            continue
        finally:
            if urval:
                try:
                    lyr.setSelectionSet('NEW', urval)
                except Exception:
                    pass
        if hur:
            logg('  %s: definitionsfragan gick inte att anvanda direkt - %s' % (lagernamn, hur))
        return
    raise RuntimeError('kan inte skapa lagret (%s)' % ' / '.join(fel))


def _i_sr(geom, sr):
    """Geometrin projicerad till sr nar koordinatsystemen skiljer (omradeslagret kan ha ett annat)."""
    try:
        if sr is not None and txt(getattr(geom.spatialReference, 'name', '')) not in ('', txt(sr.name)):
            return geom.projectAs(sr)
    except Exception:
        pass
    return geom


def _fyll_omradesfalt(ut_fc, gdb, polys, union, lagernamn, klipp=None):
    """Fyller OMRADE, KORSAR_GRANS och KALLAGER och, med klipp (polygon), klipper linjerna vid den
    (KLIPPT_M = bortklippt langd, objekt helt utanfor tas bort). Returnerar (antal som korsar
    omradesgransen, {klippta, meter, bort}).
    SVOA:s featureklasser har ett klasstillagg som foljer med kopian, och da gar de bara att uppdatera i
    en redigeringssession ("cannot be updated outside an edit session") - da provas en sadan. Gar inte
    det heller lamnas falten tomma (exporten i sig ar klar) och det loggas."""
    try:
        sr = arcpy.Describe(ut_fc).spatialReference
    except Exception:
        sr = None
    union_ = _i_sr(union, sr)
    polys_ = [(nm, _i_sr(p, sr)) for nm, p in polys]
    klipp_ = _i_sr(klipp, sr) if klipp is not None else None

    def skriv():
        n = 0
        st = {'klippta': 0, 'meter': 0.0, 'bort': 0}
        falt = ['SHAPE@', FALT_OMRADE, FALT_KORSAR, FALT_KALLA] + ([FALT_KLIPPT] if klipp_ is not None else [])
        with arcpy.da.UpdateCursor(ut_fc, falt) as cur:
            for rad in cur:
                g = rad[0]
                namn_i = [nm for nm, p in polys_ if g is not None and not p.disjoint(g)]
                inne = 1 if (g is None or g.within(union_)) else 0
                if not inne:
                    n += 1
                ny = [g, ';'.join(namn_i)[:254], 0 if inne else 1, lagernamn[:100]]
                if klipp_ is not None:
                    bort_m = 0.0
                    if g is not None and not inne:
                        L0 = g.length
                        g2 = g.intersect(klipp_, 2)
                        if g2 is None or not g2.length > 0:
                            cur.deleteRow()
                            st['bort'] += 1
                            continue
                        if g2.length < L0 - 0.01:
                            bort_m = L0 - g2.length
                            st['klippta'] += 1
                            st['meter'] += bort_m
                            ny[0] = g2
                    ny.append(round(bort_m, 2))
                cur.updateRow(ny)
        return n, st
    try:
        return skriv()
    except RuntimeError as e:
        if 'edit session' not in txt(e).lower():
            raise
    ed = None
    try:
        ed = arcpy.da.Editor(gdb)
        ed.startEditing(False, False)
        ed.startOperation()
        n = skriv()
        ed.stopOperation()
        ed.stopEditing(True)
        logg('  %s: falten fylldes i en redigeringssession (klasstillagg i kallan)' % lagernamn)
        return n
    except Exception as e:
        if ed is not None:
            try:
                ed.stopEditing(False)
            except Exception:
                pass
        logg('  OBS %s: kunde inte fylla %s/%s/%s (%s) - objekten ar exporterade, falten ar tomma%s'
             % (lagernamn, FALT_OMRADE, FALT_KORSAR, FALT_KALLA, txt(e),
                ' och ledningarna INTE klippta' if klipp is not None else ''))
        return 0, {'klippta': 0, 'meter': 0.0, 'bort': 0}


def _ta_bort_ur_kartan(mxd, ut_fc):
    """Tar bort ett tidigare exporterat lager ur kartan sa att featureklassen kan skrivas over."""
    if mxd is None:
        return
    try:
        for l in arcpy.mapping.ListLayers(mxd):
            if getattr(l, 'supports', lambda x: False)('DATASOURCE') and \
                    os.path.normcase(txt(l.dataSource)) == os.path.normcase(txt(ut_fc)):
                arcpy.mapping.RemoveLayer(arcpy.mapping.ListDataFrames(mxd)[0], l)
    except Exception as e:
        logg('  kunde inte ta bort det gamla lagret ur kartan: %s' % txt(e))


def _lagg_in(df, ut_fc, kall_lyr, lagernamn):
    """Lagger det lokala lagret i kartan med kallagrets symbologi."""
    try:
        ny = arcpy.mapping.Layer(ut_fc)
        ny.name = lagernamn + LAGERSUFFIX
        if not isinstance(kall_lyr, (TEXTTYP, bytes)):
            try:
                arcpy.mapping.UpdateLayer(df, ny, kall_lyr, True)     # bara symbologin
                ny.name = lagernamn + LAGERSUFFIX
            except Exception as e:
                logg('  symbologin kunde inte kopieras fran %s: %s' % (lagernamn, txt(e)))
        arcpy.mapping.AddLayer(df, ny, 'TOP')
    except Exception as e:
        logg('  kunde inte lagga %s i kartan: %s' % (lagernamn, txt(e)))
