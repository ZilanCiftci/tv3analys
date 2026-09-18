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
        self.tools = [SkapaLedningslager, UppdateraBedomning]


class SkapaLedningslager(object):
    def __init__(self):
        self.label = 'Skapa ledningslager'
        self.description = (
            'Skapar ett ledningslager av kartunderlag.json fr\u00e5n tv3_analys. '
            'Varje brunnspar i filen letas upp i brunnslagren och ledningen mellan '
            'brunnarna klipps ut som ett eget objekt med f\u00e4lten Maskinell bed\u00f6mning '
            'och Manuell bed\u00f6mning. Manuella bed\u00f6mningar fr\u00e5n en tidigare k\u00f6rning bevaras.')
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
            displayName='Max antal brunnar en str\u00e4cka f\u00e5r passera', name='max_hopp',
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

        return [json_in, ledning, brunn, brunn_id, ut_fc, omrade, lyr_fil, csv_ut,
                tolerans, max_hopp, marginal, kopiera]

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
        )
        parameters[4].value = ut
        if parameters[6].valueAsText and os.path.isfile(parameters[6].valueAsText):
            parameters[4].symbology = parameters[6].valueAsText
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
