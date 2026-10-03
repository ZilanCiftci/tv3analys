# -*- coding: utf-8 -*-
"""
Python Toolbox for ArcMap 10.x med verktygen fran tv3_analys.

OBS: Filen ska vara ren ASCII. ArcMap laser .pyt-filer med Windows-teckentabellen
oavsett kodningsrad, sa a-ring, a-prickar och o-prickar i etiketter skrivs som
unicode-koder: u00e5, u00e4 och u00f6 efter ett omvant snedstreck. Tack vare
unicode_literals blir de riktiga tecken i dialogen.

Lagg till i ArcToolbox: hogerklicka > Add Toolbox > valj denna .pyt-fil.
Logiken ligger i skapa_ledningslager.py i samma mapp; den har filen ar bara
dialogerna. Modulen laddas om vid varje korning.

Lagervalen ar rullistor med kartans lagernamn (inte lagerparametrar): ArcMap
tolkar '/' i ett lagernamn som sokvag nar namnet skickas som text, vilket gor
att t.ex. "A Rensbrunn/tillsynsbrunn" annars inte hittas. Namnen slas upp
till lagerobjekt via arcpy.mapping i skapa_ledningslager.py.
"""
from __future__ import unicode_literals

import os
import sys
import arcpy

HAR = os.path.dirname(os.path.abspath(__file__))
if HAR not in sys.path:
    sys.path.insert(0, HAR)

# .lyr-fil med symbologi som anvands om ingen annan anges (sparas fran ArcMap, se handledningen 7.2)
STANDARD_LYR = os.path.join(HAR, 'bedomda_ledningar.lyr')

# Lager som fylls i automatiskt om de finns i kartan (exakt namn, skiftlage spelar ingen roll)
STANDARD_LEDNING = ['A Ledning']
STANDARD_BRUNN = ['A Nedstign och \u00f6vriga brunnar', 'A Rensbrunn/tillsynsbrunn',
                  'A Platsgjuten brunnspunkt']
STANDARD_CSV = 'brunnsfel.csv'          # foreslas bredvid JSON-filen, som shapefilen
STANDARD_MARKPROFIL = 'markprofil.json'  # foreslas bredvid kartunderlag.json
STANDARD_UPPSTROMS = 'uppstroms.csv'     # batchresultat fran Uppstroms, foreslas bredvid lagret
STANDARD_VG_FRAN = ['VG_UPP', 'VG_FRAN', 'VATTENGANG_UPP', 'VGUPP']    # gissningar pa faltnamn
STANDARD_VG_TILL = ['VG_NED', 'VG_TILL', 'VATTENGANG_NED', 'VGNED']


def _ladda_modul(namn='skapa_ledningslager'):
    """Importerar (och laddar om) en modul i arcmap-mappen sa att andringar slar igenom."""
    modul = __import__(namn)
    try:
        reload(modul)                 # Python 2
    except NameError:
        import importlib
        importlib.reload(modul)       # Python 3
    return modul


def _filter(param, lista):
    """Satter filterlista om parametertypen har ett filter (utdata-filer saknar det)."""
    try:
        if param.filter is not None:
            param.filter.list = lista
    except Exception:
        pass


def _kartlager():
    """Featurelagren i den oppna kartan: lista med (langt namn, lagerobjekt) i TOC-ordning.
    Langt namn = Grupp\\Lager, sa att lager med samma namn i olika grupper kan skiljas."""
    ut = []
    try:
        mxd = arcpy.mapping.MapDocument('CURRENT')
        for l in arcpy.mapping.ListLayers(mxd):
            try:
                if l.isFeatureLayer:
                    ut.append((l.longName, l))
            except Exception:
                pass
    except Exception:
        pass
    return ut


def _langa_namn(namnlista, kartlager):
    """Langa namn for de lager i namnlista som finns i kartan (matchar kort eller langt namn)."""
    sokta = [n.strip().lower() for n in namnlista]
    ut = []
    for langt, l in kartlager:
        if l.name.strip().lower() in sokta or langt.strip().lower() in sokta:
            if langt not in ut:
                ut.append(langt)
    return ut


def _lagerparam(displayName, name, multi, kartlager, parameterType='Required'):
    """Rullista (GPString) med kartans lager. Fri text tillats ocksa, t.ex. en sokvag."""
    p = arcpy.Parameter(displayName=displayName, name=name, datatype='GPString',
                        parameterType=parameterType, direction='Input', multiValue=multi)
    namn = [langt for langt, l in kartlager]
    if namn:
        _filter(p, namn)
    return p


def _satt_varden(param, varden):
    """Fyller i en (multivalue-)parameter."""
    if not varden:
        return
    try:
        param.values = list(varden)
    except Exception:
        param.value = ';'.join(varden)


def _lagerlista(param):
    """Multivalue-parameter -> lista med namn (citattecken bortskalade)."""
    text = param.valueAsText or ''
    return [v.strip().strip("'") for v in text.split(';') if v.strip()]


def _lagerobjekt(namn):
    """Lagerobjekt i kartan for ett kort eller langt namn, annars None."""
    sokt = namn.strip().strip("'").lower()
    for langt, l in _kartlager():
        if langt.strip().lower() == sokt or l.name.strip().lower() == sokt:
            return l
    return None


def _faltnamn(param):
    """Faltnamnen i det (forsta) lager som en lagerparameter pekar pa, annars []."""
    namn = _lagerlista(param)
    if not namn:
        return []
    l = _lagerobjekt(namn[0])
    try:
        return [f.name for f in arcpy.ListFields(l.dataSource if l else namn[0])]
    except Exception:
        return []


def _forsta_traff(kandidater, faltnamn):
    """Forsta kandidaten som finns bland faltnamnen (skiftlage spelar ingen roll)."""
    upp = dict((f.upper(), f) for f in faltnamn)
    for k in kandidater:
        if k.upper() in upp:
            return upp[k.upper()]
    return None


def _rasterlager():
    """Rasterlagren i kartan (namn), for markhojder ur en hojdmodell."""
    ut = []
    try:
        mxd = arcpy.mapping.MapDocument('CURRENT')
        for l in arcpy.mapping.ListLayers(mxd):
            try:
                if l.isRasterLayer:
                    ut.append(l.longName)
            except Exception:
                pass
    except Exception:
        pass
    return ut


def _kolla_geometri(param, tillatna, vad):
    """Varnar om nagot valt lager har fel geometrityp eller inte finns i kartan."""
    if not param.valueAsText:
        return
    fel, saknas = [], []
    for namn in _lagerlista(param):
        l = _lagerobjekt(namn)
        if l is None:
            if not arcpy.Exists(namn):
                saknas.append(namn)
            continue
        try:
            typ = arcpy.Describe(l.dataSource).shapeType
        except Exception:
            continue
        if typ not in tillatna:
            fel.append('%s (%s)' % (l.name, typ))
    if saknas:
        param.setErrorMessage('Finns inte i kartan: %s' % ', '.join(saknas))
    elif fel:
        param.setWarningMessage('%s bor vara %s: %s' % (vad, '/'.join(tillatna), ', '.join(fel)))


class Toolbox(object):
    def __init__(self):
        self.label = 'tv3_analys'
        self.alias = 'tv3'
        self.tools = [SkapaLedningslager, UppdateraBedomning, Markprofil, Uppstroms,
                      ExporteraKartbild, SkapaProjekteringslager, Projekteringsprofil]


class SkapaLedningslager(object):
    def __init__(self):
        self.label = 'Skapa ledningslager'
        self.description = (
            'Skapar ett ledningslager av kartunderlag.json fr\u00e5n tv3_analys. '
            'Varje brunnspar i filen letas upp i brunnslagren och ledningen mellan '
            'brunnarna klipps ut som ett eget objekt med f\u00e4lten Maskinell bed\u00f6mning '
            'och Manuell bed\u00f6mning. Manuella bed\u00f6mningar fr\u00e5n en tidigare k\u00f6rning bevaras. '
            'Valfritt skrivs ocks\u00e5 svackor (punkter) och bakfall (linjer) ur inklinometerprofilerna.')
        self.canRunInBackground = False

    def getParameterInfo(self):
        kartlager = _kartlager()

        json_in = arcpy.Parameter(
            displayName='Kartunderlag (kartunderlag.json fr\u00e5n tv3_analys)',
            name='json_in', datatype='DEFile', parameterType='Required', direction='Input')
        _filter(json_in, ['json'])

        ledning = _lagerparam('Ledningslager', 'ledningslager', True, kartlager)
        brunn = _lagerparam('Brunnslager (v\u00e4lj alla lager d\u00e4r brunnar kan ligga)',
                            'brunnslager', True, kartlager)

        brunn_id = arcpy.Parameter(
            displayName='F\u00e4lt med brunnsbeteckning i brunnslagren', name='brunn_id',
            datatype='GPString', parameterType='Required', direction='Input')
        brunn_id.value = 'EntityID'

        ut_fc = arcpy.Parameter(
            displayName='Utdata (featureklass i geodatabas, eller shapefil)', name='ut_fc',
            datatype='DEFeatureClass', parameterType='Required', direction='Output')

        omrade = _lagerparam('Begr\u00e4nsa till omr\u00e5de (polygonlager, valfritt)',
                             'omradeslager', False, kartlager, 'Optional')

        lyr_fil = arcpy.Parameter(
            displayName='Symbologi (.lyr-fil, valfritt)', name='lyr_fil',
            datatype='DELayer', parameterType='Optional', direction='Input')
        if os.path.isfile(STANDARD_LYR):
            lyr_fil.value = STANDARD_LYR

        csv_ut = arcpy.Parameter(
            displayName='Rapport \u00f6ver omatchade brunnspar (.csv, valfritt)', name='csv_ut',
            datatype='DEFile', parameterType='Optional', direction='Output')
        _filter(csv_ut, ['csv'])

        tolerans = arcpy.Parameter(
            displayName='Tolerans mellan ledningens vertex och brunnen (m)', name='tolerans',
            datatype='GPDouble', parameterType='Required', direction='Input',
            category='Matchning')
        tolerans.value = 2.0

        max_hopp = arcpy.Parameter(
            displayName='Max antal brunnar en str\u00e4cka f\u00e5r passera, inkl. slutbrunnen (2 = en mellanbrunn)',
            name='max_hopp',
            datatype='GPLong', parameterType='Required', direction='Input',
            category='Matchning')
        max_hopp.value = 2

        marginal = arcpy.Parameter(
            displayName='Marginal utanf\u00f6r omr\u00e5det d\u00e4r brunnar \u00e4nd\u00e5 l\u00e4ses in (m)',
            name='marginal', datatype='GPDouble', parameterType='Required', direction='Input',
            category='Matchning')
        marginal.value = 100.0

        kopiera = arcpy.Parameter(
            displayName='F\u00e4lt fr\u00e5n ledningslagret som ska f\u00f6lja med', name='kopiera_falt',
            datatype='GPString', parameterType='Optional', direction='Input',
            multiValue=True, category='Matchning')

        _satt_varden(ledning, _langa_namn(STANDARD_LEDNING, kartlager))
        _satt_varden(brunn, _langa_namn(STANDARD_BRUNN, kartlager))

        rapportmapp = arcpy.Parameter(
            displayName='Mapp med PDF-rapporterna, som den h\u00e4r datorn ser den (valfritt)',
            name='rapportmapp', datatype='DEFolder', parameterType='Optional', direction='Input',
            category='Hyperl\u00e4nkar (n\u00e4r ArcMap k\u00f6rs p\u00e5 en annan dator, t.ex. Citrix)')
        filmmapp = arcpy.Parameter(
            displayName='Mapp med filmerna, som den h\u00e4r datorn ser den (valfritt)',
            name='filmmapp', datatype='DEFolder', parameterType='Optional', direction='Input',
            category='Hyperl\u00e4nkar (n\u00e4r ArcMap k\u00f6rs p\u00e5 en annan dator, t.ex. Citrix)')

        geojson_ut = arcpy.Parameter(
            displayName='GeoJSON f\u00f6r webb-GIS (WGS84, 2D; valfritt)', name='geojson_ut',
            datatype='DEFile', parameterType='Optional', direction='Output')
        _filter(geojson_ut, ['geojson', 'json'])

        svackor_ut = arcpy.Parameter(
            displayName='Svackor (punktlager, valfritt)', name='svackor_ut',
            datatype='DEFeatureClass', parameterType='Optional', direction='Output',
            category='Svackor och bakfall (ur inklinometerprofilerna)')
        bakfall_ut = arcpy.Parameter(
            displayName='Bakfall (linjelager, valfritt)', name='bakfall_ut',
            datatype='DEFeatureClass', parameterType='Optional', direction='Output',
            category='Svackor och bakfall (ur inklinometerprofilerna)')

        return [json_in, ledning, brunn, brunn_id, ut_fc, omrade, lyr_fil, csv_ut,
                tolerans, max_hopp, marginal, kopiera, rapportmapp, filmmapp, geojson_ut,
                svackor_ut, bakfall_ut]

    def isLicensed(self):
        return True

    def updateParameters(self, parameters):
        # Foresla utdata och CSV-rapport bredvid JSON-filen
        if parameters[0].altered and parameters[0].valueAsText:
            mapp = os.path.dirname(parameters[0].valueAsText)
            if not parameters[4].altered:
                parameters[4].value = os.path.join(mapp, 'bedomda_ledningar.shp')
            if not parameters[7].altered:
                parameters[7].value = os.path.join(mapp, STANDARD_CSV)
        return

    def updateMessages(self, parameters):
        if parameters[0].valueAsText and not os.path.isfile(parameters[0].valueAsText):
            parameters[0].setErrorMessage('Filen finns inte. Kor tv3_analys.py forst.')
        _kolla_geometri(parameters[1], ('Polyline',), 'Ledningslager')
        _kolla_geometri(parameters[2], ('Point',), 'Brunnslager')
        _kolla_geometri(parameters[5], ('Polygon',), 'Omradeslager')
        if parameters[8].value is not None and not parameters[8].value > 0:
            parameters[8].setErrorMessage('Storre an 0')
        if parameters[9].value is not None and parameters[9].value < 1:
            parameters[9].setErrorMessage('Minst 1')
        return

    def execute(self, parameters, messages):
        m = _ladda_modul()
        omrade = _lagerlista(parameters[5])
        ut = m.skapa(
            parameters[0].valueAsText,
            _lagerlista(parameters[1]),
            _lagerlista(parameters[2]),
            parameters[3].valueAsText,
            parameters[4].valueAsText,
            omradeslager=omrade[0] if omrade else None,
            csv_ut=parameters[7].valueAsText or None,
            lyr_fil=parameters[6].valueAsText or None,
            tolerans=float(parameters[8].value),
            marginal=float(parameters[10].value),
            max_hopp=int(parameters[9].value),
            kopiera_falt=_lagerlista(parameters[11]),
            lagg_till_i_kartan=False,     # ArcMap lagger sjalv till utdata-parametern i kartan
            rapportmapp=parameters[12].valueAsText or None,
            filmmapp=parameters[13].valueAsText or None,
            geojson_ut=parameters[14].valueAsText or None,
            svackor_ut=parameters[15].valueAsText or None,
            bakfall_ut=parameters[16].valueAsText or None,
        )
        parameters[4].value = ut
        if parameters[6].valueAsText and os.path.isfile(parameters[6].valueAsText):
            parameters[4].symbology = parameters[6].valueAsText
        for i, lyr in ((15, m.LYR_SVACKOR), (16, m.LYR_BAKFALL)):
            if parameters[i].valueAsText and os.path.isfile(lyr):
                parameters[i].symbology = lyr
        return


class UppdateraBedomning(object):
    def __init__(self):
        self.label = 'Uppdatera bed\u00f6mning'
        self.description = (
            'R\u00e4knar om G\u00e4llande bed\u00f6mning, Bed\u00f6mningstyp och STIL efter att du fyllt i '
            'Manuell bed\u00f6mning. Geometrin r\u00f6rs inte.')
        self.canRunInBackground = False

    def getParameterInfo(self):
        lager = _lagerparam('Ledningslager fr\u00e5n "Skapa ledningslager"', 'lager',
                            False, _kartlager())
        return [lager]

    def isLicensed(self):
        return True

    def updateMessages(self, parameters):
        _kolla_geometri(parameters[0], ('Polyline',), 'Lagret')
        if parameters[0].valueAsText:
            l = _lagerobjekt(parameters[0].valueAsText)
            try:
                namn = [f.name.upper() for f in arcpy.ListFields(l.dataSource if l else parameters[0].valueAsText)]
                if 'MAN_BED' not in namn or 'MASK_BED' not in namn:
                    parameters[0].setErrorMessage(
                        'Lagret saknar faltet MASK_BED/MAN_BED - valj lagret fran "Skapa ledningslager".')
            except Exception:
                pass
        return

    def execute(self, parameters, messages):
        m = _ladda_modul()
        m.uppdatera(parameters[0].valueAsText.strip("'"))
        try:
            arcpy.RefreshActiveView()
        except Exception:
            pass
        return


class Markprofil(object):
    def __init__(self):
        self.label = 'Markprofil'
        self.description = (
            'Tar ut markh\u00f6jden l\u00e4ngs varje str\u00e4cka i ledningslagret fr\u00e5n "Skapa '
            'ledningslager" ur ett punktlager med markh\u00f6jder, och vatteng\u00e5ngsniv\u00e5erna i '
            'b\u00e5da \u00e4ndarna ur det ursprungliga ledningslagret. Skriver markprofil.json som '
            'tv3_analys.py anv\u00e4nder f\u00f6r att rita mark mot ledning i protokollen och '
            'r\u00e4kna t\u00e4ckning.')
        self.canRunInBackground = False

    def getParameterInfo(self):
        kartlager = _kartlager()

        bedomda = _lagerparam('Ledningslager fr\u00e5n "Skapa ledningslager"', 'bedomda',
                              False, kartlager)
        ledning = _lagerparam('Ursprungligt ledningslager (med vatteng\u00e5ng)', 'ledningslager',
                              True, kartlager)
        vg_fran = arcpy.Parameter(
            displayName='F\u00e4lt med vatteng\u00e5ng vid ledningens startpunkt (uppstr\u00f6ms)',
            name='vg_fran', datatype='GPString', parameterType='Required', direction='Input')
        vg_till = arcpy.Parameter(
            displayName='F\u00e4lt med vatteng\u00e5ng vid ledningens slutpunkt (nedstr\u00f6ms)',
            name='vg_till', datatype='GPString', parameterType='Required', direction='Input')
        mark = _lagerparam('Markh\u00f6jder (punktlager)', 'marklager', False, kartlager)
        z_falt = arcpy.Parameter(
            displayName='F\u00e4lt med markh\u00f6jd (tomt = punkternas Z)', name='z_falt',
            datatype='GPString', parameterType='Optional', direction='Input')
        json_ut = arcpy.Parameter(
            displayName='Utdata (markprofil.json, l\u00e4ggs bredvid kartunderlag.json)',
            name='json_ut', datatype='DEFile', parameterType='Required', direction='Output')
        _filter(json_ut, ['json'])

        intervall = arcpy.Parameter(
            displayName='Avst\u00e5nd mellan markh\u00f6jdsproven l\u00e4ngs ledningen (m)',
            name='intervall', datatype='GPDouble', parameterType='Required', direction='Input',
            category='Inst\u00e4llningar')
        intervall.value = 1.0
        sokradie = arcpy.Parameter(
            displayName='S\u00f6kradie f\u00f6r markh\u00f6jdspunkter (m)', name='sokradie',
            datatype='GPDouble', parameterType='Required', direction='Input',
            category='Inst\u00e4llningar')
        sokradie.value = 5.0
        tolerans = arcpy.Parameter(
            displayName='Tolerans mellan str\u00e4ckans \u00e4nde och ledningens \u00e4nde (m)',
            name='tolerans', datatype='GPDouble', parameterType='Required', direction='Input',
            category='Inst\u00e4llningar')
        tolerans.value = 2.0
        hojdsystem = arcpy.Parameter(
            displayName='H\u00f6jdsystem i kartan (skrivs i filen)', name='hojdsystem',
            datatype='GPString', parameterType='Required', direction='Input',
            category='Inst\u00e4llningar')
        hojdsystem.value = 'RH2000'

        _satt_varden(ledning, _langa_namn(STANDARD_LEDNING, kartlager))
        return [bedomda, ledning, vg_fran, vg_till, mark, z_falt, json_ut,
                intervall, sokradie, tolerans, hojdsystem]

    def isLicensed(self):
        return True

    def updateParameters(self, parameters):
        # Faltlistor ur valda lager, och gissning pa vattengangsfalten
        if parameters[1].altered and parameters[1].valueAsText:
            falt = _faltnamn(parameters[1])
            if falt:
                _filter(parameters[2], falt)
                _filter(parameters[3], falt)
                if not parameters[2].altered:
                    parameters[2].value = _forsta_traff(STANDARD_VG_FRAN, falt)
                if not parameters[3].altered:
                    parameters[3].value = _forsta_traff(STANDARD_VG_TILL, falt)
        if parameters[4].altered and parameters[4].valueAsText:
            falt = _faltnamn(parameters[4])
            if falt:
                _filter(parameters[5], [''] + falt)
        if parameters[0].altered and parameters[0].valueAsText and not parameters[6].altered:
            l = _lagerobjekt(parameters[0].valueAsText)
            try:
                mapp = os.path.dirname(l.dataSource if l else parameters[0].valueAsText)
                if mapp.lower().endswith('.gdb'):          # inte inuti geodatabasmappen
                    mapp = os.path.dirname(mapp)
                if mapp:
                    parameters[6].value = os.path.join(mapp, STANDARD_MARKPROFIL)
            except Exception:
                pass
        return

    def updateMessages(self, parameters):
        _kolla_geometri(parameters[0], ('Polyline',), 'Ledningslagret')
        _kolla_geometri(parameters[1], ('Polyline',), 'Ledningslager')
        _kolla_geometri(parameters[4], ('Point',), 'Markh\u00f6jder')
        if parameters[0].valueAsText:
            namn = [f.upper() for f in _faltnamn(parameters[0])]
            if namn and 'FRAN_BRUNN' not in namn:
                parameters[0].setErrorMessage(
                    'Lagret saknar faltet FRAN_BRUNN - valj lagret fran "Skapa ledningslager".')
            elif namn and 'NR' not in namn:
                parameters[0].setWarningMessage(
                    'Lagret saknar faltet NR (skapat med en aldre version) - kor "Skapa '
                    'ledningslager" igen for saker koppling till TV3-filen.')
        for i in (7, 8, 9):
            if parameters[i].value is not None and not parameters[i].value > 0:
                parameters[i].setErrorMessage('Storre an 0')
        return

    def execute(self, parameters, messages):
        m = _ladda_modul('markprofil')
        _ladda_modul('skapa_ledningslager')
        m.markprofil(
            parameters[0].valueAsText.strip("'"),
            _lagerlista(parameters[1]),
            parameters[2].valueAsText,
            parameters[3].valueAsText,
            parameters[4].valueAsText.strip("'"),
            parameters[6].valueAsText,
            z_falt=(parameters[5].valueAsText or '').strip() or None,
            intervall=float(parameters[7].value),
            sokradie=float(parameters[8].value),
            tolerans=float(parameters[9].value),
            hojdsystem=parameters[10].valueAsText or 'RH2000',
        )
        return


class Uppstroms(object):
    def __init__(self):
        self.label = 'Uppstr\u00f6ms'
        self.description = (
            'Allt som ligger uppstr\u00f6ms om en brunn eller en markerad ledning: ledningar, '
            'brunnar, serviser, total l\u00e4ngd och l\u00e4ngsta gren. Fl\u00f6desriktningen tas ur '
            'ett riktningsattribut i ledningslagret, med vatteng\u00e5ngsf\u00e4lten och ritad '
            'riktning som reserv. I batchl\u00e4get r\u00e4knas serviser och l\u00e4ngd uppstr\u00f6ms '
            'f\u00f6r varje str\u00e4cka i lagret fr\u00e5n "Skapa ledningslager" och skrivs till '
            'lagret (ANT_SERV_U, L_UPPSTR) och en CSV f\u00f6r tv3_analys.py.')
        self.canRunInBackground = False

    def getParameterInfo(self):
        kartlager = _kartlager()
        K_RIKT = 'Fl\u00f6desriktning'
        K_SERV = 'Serviser'
        K_STOPP = 'Stopp (passeras inte)'
        K_BEGR = 'Begr\u00e4nsning'
        K_BATCH = 'Batch: per bed\u00f6md str\u00e4cka'

        ledning = _lagerparam('Ledningslager', 'ledningslager', True, kartlager)
        brunn = _lagerparam('Brunnslager (alla lager d\u00e4r brunnar kan ligga)', 'brunnslager',
                            True, kartlager)
        brunn_id = arcpy.Parameter(
            displayName='F\u00e4lt med brunnsbeteckning i brunnslagren', name='brunn_id',
            datatype='GPString', parameterType='Required', direction='Input')
        brunn_id.value = 'EntityID'
        startbrunn = arcpy.Parameter(
            displayName='Startbrunn (tomt = den markerade ledningen i kartan)', name='startbrunn',
            datatype='GPString', parameterType='Optional', direction='Input')
        ut_fc = arcpy.Parameter(
            displayName='Utdata: lager med allt uppstr\u00f6ms (valfritt)', name='ut_fc',
            datatype='DEFeatureClass', parameterType='Optional', direction='Output')
        csv_ut = arcpy.Parameter(
            displayName='Sammanfattning / batchresultat (.csv, valfritt)', name='csv_ut',
            datatype='DEFile', parameterType='Optional', direction='Output')
        _filter(csv_ut, ['csv'])

        riktn = arcpy.Parameter(
            displayName='F\u00e4lt med fl\u00f6desriktning (valfritt)', name='riktningsfalt',
            datatype='GPString', parameterType='Optional', direction='Input', category=K_RIKT)
        med = arcpy.Parameter(
            displayName='V\u00e4rde(n) som betyder med ritad riktning (flera med ;)', name='med_varden',
            datatype='GPString', parameterType='Optional', direction='Input', category=K_RIKT)
        med.value = 'MED'
        mot = arcpy.Parameter(
            displayName='V\u00e4rde(n) som betyder mot ritad riktning (flera med ;)', name='mot_varden',
            datatype='GPString', parameterType='Optional', direction='Input', category=K_RIKT)
        mot.value = 'MOT'
        vg_fran = arcpy.Parameter(
            displayName='F\u00e4lt med vatteng\u00e5ng vid startpunkten (reserv, valfritt)', name='vg_fran',
            datatype='GPString', parameterType='Optional', direction='Input', category=K_RIKT)
        vg_till = arcpy.Parameter(
            displayName='F\u00e4lt med vatteng\u00e5ng vid slutpunkten (reserv, valfritt)', name='vg_till',
            datatype='GPString', parameterType='Optional', direction='Input', category=K_RIKT)

        servis = _lagerparam('Servislager (valfritt)', 'servislager', False, kartlager, 'Optional')
        servis.category = K_SERV
        servis_falt = arcpy.Parameter(
            displayName='...eller f\u00e4lt i ledningslagret som anger servis', name='servis_falt',
            datatype='GPString', parameterType='Optional', direction='Input', category=K_SERV)
        servis_varden = arcpy.Parameter(
            displayName='V\u00e4rde(n) som betyder servis (flera med ;)', name='servis_varden',
            datatype='GPString', parameterType='Optional', direction='Input', category=K_SERV)

        stopp_falt = arcpy.Parameter(
            displayName='F\u00e4lt i ledningslagret f\u00f6r stopp (t.ex. ledningstyp)', name='stopp_falt',
            datatype='GPString', parameterType='Optional', direction='Input', category=K_STOPP)
        stopp_varden = arcpy.Parameter(
            displayName='V\u00e4rde(n) som inte passeras, t.ex. tryckledning (flera med ;)',
            name='stopp_varden', datatype='GPString', parameterType='Optional', direction='Input',
            category=K_STOPP)
        stopp_brunnar = arcpy.Parameter(
            displayName='Brunnar d\u00e4r s\u00f6kningen stannar, t.ex. pumpstationer (littera, flera med ;)',
            name='stopp_brunnar', datatype='GPString', parameterType='Optional', direction='Input',
            category=K_STOPP)

        omrade = _lagerparam('Begr\u00e4nsa till omr\u00e5de (polygonlager, valfritt)', 'omradeslager',
                             False, kartlager, 'Optional')
        omrade.category = K_BEGR
        sokradie = arcpy.Parameter(
            displayName='S\u00f6kradie kring startbrunnen (m, tomt = hela lagret)', name='sokradie',
            datatype='GPDouble', parameterType='Optional', direction='Input', category=K_BEGR)
        tolerans = arcpy.Parameter(
            displayName='Tolerans mellan ledning och brunn/servis\u00e4nde (m)', name='tolerans',
            datatype='GPDouble', parameterType='Required', direction='Input', category=K_BEGR)
        tolerans.value = 1.0

        bedomda = _lagerparam('Ledningslager fr\u00e5n "Skapa ledningslager" (batchl\u00e4ge: alla '
                              'str\u00e4ckor, startbrunn ignoreras)', 'bedomda', False, kartlager, 'Optional')
        bedomda.category = K_BATCH

        _satt_varden(ledning, _langa_namn(STANDARD_LEDNING, kartlager))
        _satt_varden(brunn, _langa_namn(STANDARD_BRUNN, kartlager))
        return [ledning, brunn, brunn_id, startbrunn, ut_fc, csv_ut,
                riktn, med, mot, vg_fran, vg_till,
                servis, servis_falt, servis_varden,
                stopp_falt, stopp_varden, stopp_brunnar,
                omrade, sokradie, tolerans, bedomda]

    def isLicensed(self):
        return True

    def updateParameters(self, parameters):
        if parameters[0].altered and parameters[0].valueAsText:
            falt = _faltnamn(parameters[0])
            if falt:
                for i in (6, 9, 10, 12, 14):
                    _filter(parameters[i], [''] + falt)
                if not parameters[9].altered:
                    parameters[9].value = _forsta_traff(STANDARD_VG_FRAN, falt)
                if not parameters[10].altered:
                    parameters[10].value = _forsta_traff(STANDARD_VG_TILL, falt)
        if parameters[20].altered and parameters[20].valueAsText and not parameters[5].altered:
            l = _lagerobjekt(parameters[20].valueAsText)
            try:
                mapp = os.path.dirname(l.dataSource if l else parameters[20].valueAsText)
                if mapp.lower().endswith('.gdb'):
                    mapp = os.path.dirname(mapp)
                if mapp:
                    parameters[5].value = os.path.join(mapp, STANDARD_UPPSTROMS)
            except Exception:
                pass
        return

    def updateMessages(self, parameters):
        _kolla_geometri(parameters[0], ('Polyline',), 'Ledningslager')
        _kolla_geometri(parameters[1], ('Point',), 'Brunnslager')
        _kolla_geometri(parameters[11], ('Polyline',), 'Servislager')
        _kolla_geometri(parameters[17], ('Polygon',), 'Omradeslager')
        _kolla_geometri(parameters[20], ('Polyline',), 'Bedomda ledningar')
        if parameters[19].value is not None and not parameters[19].value > 0:
            parameters[19].setErrorMessage('Storre an 0')
        if parameters[20].valueAsText:
            namn = [f.upper() for f in _faltnamn(parameters[20])]
            if namn and 'FRAN_BRUNN' not in namn:
                parameters[20].setErrorMessage(
                    'Lagret saknar faltet FRAN_BRUNN - valj lagret fran "Skapa ledningslager".')
            if not parameters[5].valueAsText:
                parameters[5].setErrorMessage('Batchlaget behover en CSV-fil att skriva resultatet till.')
        elif not parameters[3].valueAsText and not parameters[4].valueAsText:
            parameters[4].setWarningMessage(
                'Utan utdatalager visas resultatet bara i loggen (och i CSV:n om den anges).')
        return

    def execute(self, parameters, messages):
        _ladda_modul('skapa_ledningslager')
        m = _ladda_modul('natverk')
        servis = _lagerlista(parameters[11])
        omrade = _lagerlista(parameters[17])
        gemensamt = dict(
            riktningsfalt=parameters[6].valueAsText or None,
            med_varden=parameters[7].valueAsText or None,
            mot_varden=parameters[8].valueAsText or None,
            vg_fran=parameters[9].valueAsText or None,
            vg_till=parameters[10].valueAsText or None,
            servislager=servis or None,
            servis_falt=parameters[12].valueAsText or None,
            servis_varden=parameters[13].valueAsText or None,
            stopp_falt=parameters[14].valueAsText or None,
            stopp_varden=parameters[15].valueAsText or None,
            stopp_brunnar=_lagerlista(parameters[16]) or None,
            omradeslager=omrade[0] if omrade else None,
            tolerans=float(parameters[19].value),
        )
        if parameters[20].valueAsText:
            m.uppstroms_batch(parameters[20].valueAsText.strip("'"), _lagerlista(parameters[0]),
                              _lagerlista(parameters[1]), parameters[2].valueAsText,
                              parameters[5].valueAsText, **gemensamt)
            try:
                arcpy.RefreshActiveView()
            except Exception:
                pass
            return
        ut = m.uppstroms(_lagerlista(parameters[0]), _lagerlista(parameters[1]), parameters[2].valueAsText,
                         startbrunn=parameters[3].valueAsText or None,
                         ut_fc=parameters[4].valueAsText or None,
                         csv_ut=parameters[5].valueAsText or None,
                         sokradie=float(parameters[18].value) if parameters[18].value else None,
                         lagg_till_i_kartan=False, **gemensamt)
        if ut.get('lager'):
            parameters[4].value = ut['lager']
        return


class ExporteraKartbild(object):
    def __init__(self):
        self.label = 'Exportera kartbild'
        self.description = (
            'Sparar kartans aktuella vy som PNG i utdatamappen fr\u00e5n tv3_analys (kartbild.png), '
            's\u00e5 att tv3_pptx.py kan l\u00e4gga in kartan i presentationen. Zooma och t\u00e4nd de '
            'lager som ska synas innan verktyget k\u00f6rs.')
        self.canRunInBackground = False

    def getParameterInfo(self):
        mapp = arcpy.Parameter(
            displayName='Utdatamapp fr\u00e5n tv3_analys (d\u00e4r kartunderlag.json ligger)', name='mapp',
            datatype='DEFolder', parameterType='Required', direction='Input')
        namn = arcpy.Parameter(
            displayName='Filnamn', name='namn', datatype='GPString', parameterType='Required',
            direction='Input')
        namn.value = 'kartbild.png'
        upplosning = arcpy.Parameter(
            displayName='Uppl\u00f6sning (dpi)', name='dpi', datatype='GPLong', parameterType='Required',
            direction='Input')
        upplosning.value = 200
        bredd = arcpy.Parameter(
            displayName='Bildbredd (pixlar, 0 = enligt dataramen)', name='bredd', datatype='GPLong',
            parameterType='Required', direction='Input')
        bredd.value = 2400
        ut = arcpy.Parameter(
            displayName='Kartbild', name='ut', datatype='DEFile', parameterType='Derived',
            direction='Output')
        return [mapp, namn, upplosning, bredd, ut]

    def isLicensed(self):
        return True

    def updateMessages(self, parameters):
        if parameters[1].valueAsText and not parameters[1].valueAsText.lower().endswith('.png'):
            parameters[1].setErrorMessage('Filnamnet ska sluta pa .png')
        return

    def execute(self, parameters, messages):
        mxd = arcpy.mapping.MapDocument('CURRENT')
        df = arcpy.mapping.ListDataFrames(mxd)[0]
        ut = os.path.join(parameters[0].valueAsText, parameters[1].valueAsText)
        dpi = int(parameters[2].value)
        bredd = int(parameters[3].value or 0)
        if bredd > 0:
            # Hojden foljer dataramens proportioner
            hojd = int(bredd * df.elementHeight / df.elementWidth)
            arcpy.mapping.ExportToPNG(mxd, ut, df, df_export_width=bredd, df_export_height=hojd,
                                      resolution=dpi, world_file=False)
        else:
            arcpy.mapping.ExportToPNG(mxd, ut, df, resolution=dpi, world_file=False)
        arcpy.AddMessage('Kartbild skriven: %s' % ut)
        parameters[4].value = ut
        return


class SkapaProjekteringslager(object):
    def __init__(self):
        self.label = 'Skapa projekteringslager'
        self.description = (
            'Skapar tomma ritlager f\u00f6r ett nytt VA-str\u00e5k i en filgeodatabas: <prefix>_Ledning '
            '(linje, med vatteng\u00e5ng uppstr\u00f6ms/nedstr\u00f6ms, typ, dimension, material) och '
            '<prefix>_Brunn (punkt, med lockniv\u00e5 och bottenniv\u00e5). Rita sedan med vanlig '
            'redigering och k\u00f6r "Projekteringsprofil".')
        self.canRunInBackground = False

    def getParameterInfo(self):
        kartlager = _kartlager()
        gdb = arcpy.Parameter(
            displayName='Filgeodatabas (skapas om den inte finns)', name='gdb',
            datatype='DEWorkspace', parameterType='Required', direction='Output')
        prefix = arcpy.Parameter(
            displayName='Prefix p\u00e5 lagernamnen', name='prefix',
            datatype='GPString', parameterType='Required', direction='Input')
        prefix.value = 'Proj'
        sr_lager = _lagerparam('Lager att ta koordinatsystemet fr\u00e5n (tomt = kartans)', 'sr_lager',
                               False, kartlager, 'Optional')
        _satt_varden(sr_lager, _langa_namn(STANDARD_LEDNING, kartlager)[:1])
        return [gdb, prefix, sr_lager]

    def isLicensed(self):
        return True

    def updateMessages(self, parameters):
        import re
        if parameters[1].valueAsText and not re.match(r'^[A-Za-z][A-Za-z0-9_]*$', parameters[1].valueAsText):
            parameters[1].setErrorMessage('Bara A-Z, siffror och understreck, borja med en bokstav')
        return

    def execute(self, parameters, messages):
        _ladda_modul('skapa_ledningslager')
        m = _ladda_modul('projektering')
        m.skapa_projekteringslager(
            parameters[0].valueAsText, prefix=parameters[1].valueAsText or 'Proj',
            sr_lager=(parameters[2].valueAsText or '').strip("'") or None,
            lagg_till_i_kartan=True)
        return


class Projekteringsprofil(object):
    def __init__(self):
        self.label = 'Projekteringsprofil'
        self.description = (
            'Ritar profilen l\u00e4ngs ett ritat VA-str\u00e5k: markyta, vatteng\u00e5ng och hj\u00e4ssa per '
            'ledning, brunnar fr\u00e5n botten till lock, fall i promille och sektioner. Alla '
            'ledningstyper i samma diagram. \u00c4r n\u00e5gra ledningar valda i kartan ritas bara de. '
            'Utdata: PDF, PNG, CSV och JSON per str\u00e5k.')
        self.canRunInBackground = False

    def getParameterInfo(self):
        kartlager = _kartlager()
        ledning = _lagerparam('Ledningslager (<prefix>_Ledning)', 'ledningslager', False, kartlager)
        brunn = _lagerparam('Brunnslager (<prefix>_Brunn)', 'brunnslager', False, kartlager, 'Optional')
        mark = _lagerparam('Markh\u00f6jder (punktlager eller raster)', 'marklager', False, kartlager, 'Optional')
        _filter(mark, [langt for langt, l in kartlager] + _rasterlager())
        z_falt = arcpy.Parameter(
            displayName='F\u00e4lt med markh\u00f6jd (tomt = punkternas Z / rastrets v\u00e4rde)', name='z_falt',
            datatype='GPString', parameterType='Optional', direction='Input')
        ut_mapp = arcpy.Parameter(
            displayName='Utdatamapp', name='ut_mapp', datatype='DEFolder',
            parameterType='Required', direction='Input')
        namn = arcpy.Parameter(
            displayName='Namn p\u00e5 str\u00e5ket (filnamn)', name='namn',
            datatype='GPString', parameterType='Required', direction='Input')
        namn.value = 'profil'
        startbrunn = arcpy.Parameter(
            displayName='Startbrunn (valfritt, annars uppstr\u00f6ms \u00e4nde)', name='startbrunn',
            datatype='GPString', parameterType='Optional', direction='Input')
        bara_valda = arcpy.Parameter(
            displayName='Bara valda ledningar (om n\u00e5got \u00e4r valt)', name='bara_valda',
            datatype='GPBoolean', parameterType='Optional', direction='Input')
        bara_valda.value = True
        intervall = arcpy.Parameter(
            displayName='Avst\u00e5nd mellan markh\u00f6jdsproven (m)', name='intervall',
            datatype='GPDouble', parameterType='Required', direction='Input', category='Inst\u00e4llningar')
        intervall.value = 1.0
        sokradie = arcpy.Parameter(
            displayName='S\u00f6kradie f\u00f6r markh\u00f6jdspunkter (m)', name='sokradie',
            datatype='GPDouble', parameterType='Required', direction='Input', category='Inst\u00e4llningar')
        sokradie.value = 5.0
        tolerans = arcpy.Parameter(
            displayName='Tolerans ledning\u00e4nde mot brunn/annan ledning (m)', name='tolerans',
            datatype='GPDouble', parameterType='Required', direction='Input', category='Inst\u00e4llningar')
        tolerans.value = 1.0
        hojdsystem = arcpy.Parameter(
            displayName='H\u00f6jdsystem (skrivs p\u00e5 ritningen)', name='hojdsystem',
            datatype='GPString', parameterType='Required', direction='Input', category='Inst\u00e4llningar')
        hojdsystem.value = 'RH2000'
        return [ledning, brunn, mark, z_falt, ut_mapp, namn, startbrunn, bara_valda,
                intervall, sokradie, tolerans, hojdsystem]

    def isLicensed(self):
        return True

    def updateParameters(self, parameters):
        if parameters[2].altered and parameters[2].valueAsText:
            falt = _faltnamn(parameters[2])
            if falt:
                _filter(parameters[3], [''] + falt)
        if parameters[0].altered and parameters[0].valueAsText and not parameters[4].altered:
            l = _lagerobjekt(parameters[0].valueAsText)
            try:
                ds = l.dataSource if l else parameters[0].valueAsText
                mapp = os.path.dirname(ds)
                if mapp.lower().endswith('.gdb'):
                    mapp = os.path.dirname(mapp)
                if mapp:
                    parameters[4].value = mapp
            except Exception:
                pass
        return

    def updateMessages(self, parameters):
        _kolla_geometri(parameters[0], ('Polyline',), 'Ledningslagret')
        _kolla_geometri(parameters[1], ('Point',), 'Brunnslagret')
        if parameters[0].valueAsText:
            namn = [f.upper() for f in _faltnamn(parameters[0])]
            if namn and ('VG_UPP' not in namn or 'VG_NED' not in namn):
                parameters[0].setErrorMessage('Lagret saknar VG_UPP/VG_NED - skapa lagren med '
                                              '"Skapa projekteringslager".')
        for i in (8, 9, 10):
            if parameters[i].value is not None and not parameters[i].value > 0:
                parameters[i].setErrorMessage('Storre an 0')
        return

    def execute(self, parameters, messages):
        _ladda_modul('skapa_ledningslager')
        m = _ladda_modul('projektering')
        filer = m.profil(
            parameters[0].valueAsText.strip("'"),
            (parameters[1].valueAsText or '').strip("'") or None,
            (parameters[2].valueAsText or '').strip("'") or None,
            parameters[4].valueAsText,
            namn=parameters[5].valueAsText or 'profil',
            z_falt=(parameters[3].valueAsText or '').strip() or None,
            startbrunn=parameters[6].valueAsText or None,
            intervall=float(parameters[8].value),
            sokradie=float(parameters[9].value),
            tolerans=float(parameters[10].value),
            hojdsystem=parameters[11].valueAsText or 'RH2000',
            bara_valda=bool(parameters[7].value),
        )
        for f in filer:
            if f.lower().endswith('.pdf'):
                try:
                    os.startfile(f)          # oppna ritningen (Windows)
                except Exception:
                    pass
        return
