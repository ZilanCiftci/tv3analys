# -*- coding: utf-8 -*-
"""
Python Toolbox for ArcMap 10.x med verktygen fran tv3_analys.

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


def _ladda_modul():
    """Importerar (och laddar om) skapa_ledningslager sa att andringar slar igenom."""
    import skapa_ledningslager
    try:
        reload(skapa_ledningslager)                 # Python 2
    except NameError:
        import importlib
        importlib.reload(skapa_ledningslager)       # Python 3
    return skapa_ledningslager


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
        self.tools = [SkapaLedningslager, UppdateraBedomning]


class SkapaLedningslager(object):
    def __init__(self):
        self.label = 'Skapa ledningslager'
        self.description = (
            'Skapar ett ledningslager av kartunderlag.json från tv3_analys. '
            'Varje brunnspar i filen letas upp i brunnslagret och ledningen mellan '
            'brunnarna klipps ut som ett eget objekt med fälten Maskinell bedömning '
            'och Manuell bedömning. Manuella bedömningar från en tidigare körning bevaras.')
        self.canRunInBackground = False

    def getParameterInfo(self):
        json_in = arcpy.Parameter(
            displayName='Kartunderlag (kartunderlag.json från tv3_analys)',
            name='json_in', datatype='DEFile', parameterType='Required', direction='Input')
        json_in.filter.list = ['json']

        ledning = arcpy.Parameter(
            displayName='Ledningslager', name='ledningslager',
            datatype='GPFeatureLayer', parameterType='Required', direction='Input',
            multiValue=True)
        ledning.filter.list = ['Polyline']

        brunn = arcpy.Parameter(
            displayName='Brunnslager', name='brunnslager',
            datatype='GPFeatureLayer', parameterType='Required', direction='Input',
            multiValue=True)
        brunn.filter.list = ['Point']

        brunn_id = arcpy.Parameter(
            displayName='Fält med brunnsbeteckning i brunnslagret', name='brunn_id',
            datatype='GPString', parameterType='Required', direction='Input')
        brunn_id.value = 'EntityID'

        ut_fc = arcpy.Parameter(
            displayName='Utdata (featureklass i geodatabas, eller shapefil)', name='ut_fc',
            datatype='DEFeatureClass', parameterType='Required', direction='Output')

        omrade = arcpy.Parameter(
            displayName='Begränsa till område (polygonlager, valfritt)', name='omradeslager',
            datatype='GPFeatureLayer', parameterType='Optional', direction='Input')
        omrade.filter.list = ['Polygon']

        lyr_fil = arcpy.Parameter(
            displayName='Symbologi (.lyr-fil, valfritt)', name='lyr_fil',
            datatype='DELayer', parameterType='Optional', direction='Input')

        csv_ut = arcpy.Parameter(
            displayName='Rapport över omatchade brunnspar (.csv, valfritt)', name='csv_ut',
            datatype='DEFile', parameterType='Optional', direction='Output')
        csv_ut.filter.list = ['csv']

        tolerans = arcpy.Parameter(
            displayName='Tolerans mellan ledningens vertex och brunnen (m)', name='tolerans',
            datatype='GPDouble', parameterType='Required', direction='Input',
            category='Matchning')
        tolerans.value = 2.0

        max_hopp = arcpy.Parameter(
            displayName='Max antal brunnar en sträcka får passera', name='max_hopp',
            datatype='GPLong', parameterType='Required', direction='Input',
            category='Matchning')
        max_hopp.value = 2

        marginal = arcpy.Parameter(
            displayName='Marginal utanför området där brunnar ändå läses in (m)', name='marginal',
            datatype='GPDouble', parameterType='Required', direction='Input',
            category='Matchning')
        marginal.value = 100.0

        kopiera = arcpy.Parameter(
            displayName='Fält från ledningslagret som ska följa med', name='kopiera_falt',
            datatype='GPString', parameterType='Optional', direction='Input',
            multiValue=True, category='Matchning')

        return [json_in, ledning, brunn, brunn_id, ut_fc, omrade, lyr_fil, csv_ut,
                tolerans, max_hopp, marginal, kopiera]

    def isLicensed(self):
        return True

    def updateParameters(self, parameters):
        # Foresla utdatanamn bredvid JSON-filen
        if parameters[0].altered and not parameters[4].altered and parameters[0].valueAsText:
            mapp = os.path.dirname(parameters[0].valueAsText)
            parameters[4].value = os.path.join(mapp, 'bedomda_ledningar.shp')
        return

    def updateMessages(self, parameters):
        if parameters[0].valueAsText and not os.path.isfile(parameters[0].valueAsText):
            parameters[0].setErrorMessage('Filen finns inte. Kor tv3_analys.py forst.')
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
        self.label = 'Uppdatera bedömning'
        self.description = (
            'Räknar om Gällande bedömning, Bedömningstyp och STIL efter att du fyllt i '
            'Manuell bedömning. Geometrin rörs inte.')
        self.canRunInBackground = False

    def getParameterInfo(self):
        lager = arcpy.Parameter(
            displayName='Ledningslager från "Skapa ledningslager"', name='lager',
            datatype='GPFeatureLayer', parameterType='Required', direction='Input')
        lager.filter.list = ['Polyline']
        ut = arcpy.Parameter(
            displayName='Uppdaterat lager', name='ut', datatype='GPFeatureLayer',
            parameterType='Derived', direction='Output')
        ut.parameterDependencies = [lager.name]
        ut.schema.clone = True
        return [lager, ut]

    def isLicensed(self):
        return True

    def updateMessages(self, parameters):
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
