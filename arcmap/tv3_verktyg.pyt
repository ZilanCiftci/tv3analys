# -*- coding: utf-8 -*-
"""
Python Toolbox for ArcMap 10.x med verktygen fran tv3_analys.

OBS: Filen ska vara ren ASCII. ArcMap laser .pyt-filer med Windows-teckentabellen
oavsett kodningsrad, sa a-ring, a-prickar och o-prickar i etiketter skrivs som
unicode-koder: u00e5, u00e4 och u00f6 efter ett omvant snedstreck. Tack vare
unicode_literals blir de riktiga tecken i dialogen.

Lagg till i ArcToolbox: hogerklicka > Add Toolbox > valj denna .pyt-fil.
Logiken ligger i skapa_ledningslager.py i samma mapp; den har filen ar bara
dialogen. Andra i skapa_ledningslager.py och kor verktyget igen - modulen
laddas om vid varje korning.
"""
from __future__ import unicode_literals

import os
import sys
import arcpy

HAR = os.path.dirname(os.path.abspath(__file__))
if HAR not in sys.path:
    sys.path.insert(0, HAR)

# .lyr-fil med symbologi som anvands om ingen annan anges (skapas av "Skapa symbologi")
STANDARD_LYR = os.path.join(HAR, 'bedomda_ledningar.lyr')

# Lager som fylls i automatiskt om de finns i kartan (exakt namn, skiftlage spelar ingen roll)
STANDARD_LEDNING = ['A Ledning']
STANDARD_BRUNN = ['A Nedstign och \u00f6vriga brunnar', 'A Rensbrunn/tillsynsbrunn']
STANDARD_CSV = 'brunnsfel.csv'          # foreslas bredvid JSON-filen, som shapefilen


def _lager_i_kartan(namnlista):
    """Namnen i namnlista som finns som lager i den oppna kartan (aven i grupplager)."""
    try:
        mxd = arcpy.mapping.MapDocument('CURRENT')
        finns = {}
        for l in arcpy.mapping.ListLayers(mxd):
            try:
                if l.isFeatureLayer:
                    finns[l.name.strip().lower()] = l.name
            except Exception:
                pass
        return [finns[n.strip().lower()] for n in namnlista if n.strip().lower() in finns]
    except Exception:
        return []


def _satt_lager(param, namn):
    """Fyller i en (multivalue-)lagerparameter med lagernamn."""
    if not namn:
        return
    try:
        param.values = list(namn)
    except Exception:
        param.value = ';'.join(namn)


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


def _kolla_geometri(param, tillatna, vad):
    """Varnar om nagot valt lager har fel geometrityp. Inget filter anvands pa
    lagerparametrarna eftersom ArcMaps geometrifilter doljer lager i geometriska
    natverk (complex edges/junctions)."""
    if not param.valueAsText:
        return
    fel = []
    for lyr in _lagerlista(param):
        try:
            typ = arcpy.Describe(lyr).shapeType
        except Exception:
            continue
        if typ not in tillatna:
            fel.append('%s (%s)' % (getattr(lyr, 'name', lyr), typ))
    if fel:
        param.setWarningMessage('%s bor vara %s: %s' % (vad, '/'.join(tillatna), ', '.join(fel)))


def _lagerlista(param):
    """Multivalue-parameter -> lista med lagerobjekt eller namn."""
    if param.values:
        return list(param.values)
    text = param.valueAsText or ''
    return [v.strip().strip("'") for v in text.split(';') if v.strip()]


class Toolbox(object):
    def __init__(self):
        self.label = 'tv3_analys'
        self.alias = 'tv3'
        self.tools = [SkapaLedningslager, UppdateraBedomning, SkapaSymbologi]


class SkapaLedningslager(object):
    def __init__(self):
        self.label = 'Skapa ledningslager'
        self.description = (
            'Skapar ett ledningslager av kartunderlag.json fr\u00e5n tv3_analys. '
            'Varje brunnspar i filen letas upp i brunnslagret och ledningen mellan '
            'brunnarna klipps ut som ett eget objekt med f\u00e4lten Maskinell bed\u00f6mning '
            'och Manuell bed\u00f6mning. Manuella bed\u00f6mningar fr\u00e5n en tidigare k\u00f6rning bevaras.')
        self.canRunInBackground = False

    def getParameterInfo(self):
        json_in = arcpy.Parameter(
            displayName='Kartunderlag (kartunderlag.json fr\u00e5n tv3_analys)',
            name='json_in', datatype='DEFile', parameterType='Required', direction='Input')
        _filter(json_in, ['json'])

        ledning = arcpy.Parameter(
            displayName='Ledningslager', name='ledningslager',
            datatype='GPFeatureLayer', parameterType='Required', direction='Input',
            multiValue=True)

        brunn = arcpy.Parameter(
            displayName='Brunnslager (v\u00e4lj alla lager d\u00e4r brunnar kan ligga)', name='brunnslager',
            datatype='GPFeatureLayer', parameterType='Required', direction='Input',
            multiValue=True)

        brunn_id = arcpy.Parameter(
            displayName='F\u00e4lt med brunnsbeteckning i brunnslagret', name='brunn_id',
            datatype='GPString', parameterType='Required', direction='Input')
        brunn_id.value = 'EntityID'

        ut_fc = arcpy.Parameter(
            displayName='Utdata (featureklass i geodatabas, eller shapefil)', name='ut_fc',
            datatype='DEFeatureClass', parameterType='Required', direction='Output')

        omrade = arcpy.Parameter(
            displayName='Begr\u00e4nsa till omr\u00e5de (polygonlager, valfritt)', name='omradeslager',
            datatype='GPFeatureLayer', parameterType='Optional', direction='Input')

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
            displayName='Marginal utanf\u00f6r omr\u00e5det d\u00e4r brunnar \u00e4nd\u00e5 l\u00e4ses in (m)', name='marginal',
            datatype='GPDouble', parameterType='Required', direction='Input',
            category='Matchning')
        marginal.value = 100.0

        kopiera = arcpy.Parameter(
            displayName='F\u00e4lt fr\u00e5n ledningslagret som ska f\u00f6lja med', name='kopiera_falt',
            datatype='GPString', parameterType='Optional', direction='Input',
            multiValue=True, category='Matchning')

        _satt_lager(ledning, _lager_i_kartan(STANDARD_LEDNING))
        _satt_lager(brunn, _lager_i_kartan(STANDARD_BRUNN))

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
        ut = m.skapa(
            parameters[0].valueAsText,
            _lagerlista(parameters[1]),
            _lagerlista(parameters[2]),
            parameters[3].valueAsText,
            parameters[4].valueAsText,
            omradeslager=parameters[5].value if parameters[5].valueAsText else None,
            csv_ut=parameters[7].valueAsText or None,
            lyr_fil=parameters[6].valueAsText or None,
            tolerans=float(parameters[8].value),
            marginal=float(parameters[10].value),
            max_hopp=int(parameters[9].value),
            kopiera_falt=_lagerlista(parameters[11]) if parameters[11].valueAsText else [],
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
        lager = arcpy.Parameter(
            displayName='Ledningslager fr\u00e5n "Skapa ledningslager"', name='lager',
            datatype='GPFeatureLayer', parameterType='Required', direction='Input')
        ut = arcpy.Parameter(
            displayName='Uppdaterat lager', name='ut', datatype='GPFeatureLayer',
            parameterType='Derived', direction='Output')
        ut.parameterDependencies = [lager.name]
        ut.schema.clone = True
        return [lager, ut]

    def isLicensed(self):
        return True

    def updateMessages(self, parameters):
        _kolla_geometri(parameters[0], ('Polyline',), 'Lagret')
        if parameters[0].valueAsText:
            try:
                namn = [f.name.upper() for f in arcpy.ListFields(parameters[0].valueAsText)]
                if 'MAN_BED' not in namn or 'MASK_BED' not in namn:
                    parameters[0].setErrorMessage(
                        'Lagret saknar falten MASK_BED/MAN_BED - valj lagret fran "Skapa ledningslager".')
            except Exception:
                pass
        return

    def execute(self, parameters, messages):
        m = _ladda_modul()
        lyr = parameters[0].value
        sokvag = arcpy.Describe(lyr).catalogPath if lyr is not None else parameters[0].valueAsText
        m.uppdatera(sokvag)
        parameters[1].value = parameters[0].value
        try:
            arcpy.RefreshActiveView()
        except Exception:
            pass
        return


class SkapaSymbologi(object):
    def __init__(self):
        self.label = 'Skapa symbologi (.lyr)'
        self.description = (
            'Bygger symbologin f\u00f6r ledningslagret - f\u00e4rg efter prioritetsklass, '
            'heldragen linje f\u00f6r manuell bed\u00f6mning och streckad f\u00f6r maskinell - '
            's\u00e4tter den p\u00e5 lagret i kartan och sparar den som .lyr-fil. '
            'Kr\u00e4ver att comtypes finns i ArcMaps Python (pip install "comtypes<1.2").')
        self.canRunInBackground = False

    def getParameterInfo(self):
        lager = arcpy.Parameter(
            displayName='Ledningslager i kartan (fr\u00e5n "Skapa ledningslager")', name='lager',
            datatype='GPFeatureLayer', parameterType='Required', direction='Input')
        lyr_ut = arcpy.Parameter(
            displayName='Spara som (.lyr)', name='lyr_ut',
            datatype='DELayer', parameterType='Required', direction='Output')
        lyr_ut.value = STANDARD_LYR
        return [lager, lyr_ut]

    def isLicensed(self):
        return True

    def updateMessages(self, parameters):
        _kolla_geometri(parameters[0], ('Polyline',), 'Lagret')
        if parameters[0].valueAsText:
            try:
                namn = [f.name.upper() for f in arcpy.ListFields(parameters[0].valueAsText)]
                if 'STIL' not in namn:
                    parameters[0].setErrorMessage(
                        'Lagret saknar faltet STIL - valj lagret fran "Skapa ledningslager".')
            except Exception:
                pass
        return

    def execute(self, parameters, messages):
        m = _ladda_modul('skapa_lyr')
        lyr = parameters[0].value
        namn = getattr(lyr, 'name', None) or parameters[0].valueAsText
        ut = m.skapa_lyr(namn, parameters[1].valueAsText)
        parameters[1].value = ut
        return
