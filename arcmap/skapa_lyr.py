# -*- coding: utf-8 -*-
"""
Skapar symbologin for ledningslagret och sparar den som .lyr-fil:
    farg efter prioritetsklass (A rott, B orange, C gult, D gront, E gratt)
    heldragen linje = manuell bedomning, streckad linje = maskinell bedomning

Renderaren byggs med ArcObjects (via comtypes) och satts pa lagret i den oppna
kartan; darefter sparas lagret som .lyr med arcpy.mapping. Lagg .lyr-filen i
arcmap-mappen som bedomda_ledningar.lyr sa anvander verktyget "Skapa
ledningslager" den automatiskt.

comtypes (ren Python, MIT-licens) foljer med i arcmap/lib och laddas darifran om det
inte redan finns i ArcMaps Python - ingen installation behovs.

Kors i ArcMaps Python-fonster med ledningslagret i kartan:
    execfile(r'H:\PY\tv3analys\arcmap\skapa_lyr.py')
eller via verktyget "Skapa symbologi (.lyr)" i tv3_verktyg.pyt.

Forsta korningen tar en stund: comtypes genererar Python-omslag for ArcObjects
(esriCarto ar stort). Darefter gar det pa nagra sekunder.
"""
from __future__ import unicode_literals

import os
import sys
import arcpy

# =====================================================================
# KONFIG - galler vid korning med execfile
# =====================================================================

LAGER  = 'bedomda_ledningar'                          # lagrets namn i innehallsforteckningen
LYR_UT = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'bedomda_ledningar.lyr')

# =====================================================================

FALT = 'STIL'
BREDD_MANUELL   = 3.0     # punkter
BREDD_MASKINELL = 2.5

# Samma farger som i Excel och diagrammen
KLASS_FARG = {'A': (208, 59, 59), 'B': (236, 131, 90), 'C': (250, 178, 25),
              'D': (12, 163, 12), 'E': (191, 191, 191)}
KLASS_TEXT = {'A': 'A – Åtgärd snarast', 'B': 'B – Planera renovering', 'C': 'C – Bevaka',
              'D': 'D – Inga skador', 'E': 'E – Ej bedömd'}

# (varde i STIL-faltet, etikett i teckenforklaringen, streckad?, bredd)
KATEGORIER = []
for _k in 'ABCDE':
    KATEGORIER.append(('%s - Manuell' % _k, '%s (manuell bedömning)' % KLASS_TEXT[_k], False, BREDD_MANUELL))
    KATEGORIER.append(('%s - Maskinell' % _k, '%s (maskinell bedömning)' % KLASS_TEXT[_k], True, BREDD_MASKINELL))

try:
    TEXTTYP = unicode
except NameError:
    TEXTTYP = str


def txt(v):
    if v is None:
        return ''
    if isinstance(v, TEXTTYP):
        return v
    if isinstance(v, bytes):
        try:
            return v.decode('utf-8')
        except UnicodeDecodeError:
            return v.decode('cp1252', 'replace')
    return TEXTTYP(v)


def logg(*args):
    rad = ' '.join(txt(a) for a in args)
    try:
        arcpy.AddMessage(rad)
    except Exception:
        pass
    print(rad)
    sys.stdout.flush()


# ------------------------------------------------ ArcObjects

def _arcobjects():
    """Laddar ArcObjects-biblioteken via comtypes. Returnerar (comtypes.client, moduler)."""
    comtypes = None
    try:
        import comtypes.client
    except ImportError:
        # comtypes foljer med i arcmap/lib sa att ingen installation behovs
        lib = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'lib')
        if lib not in sys.path:
            sys.path.insert(0, lib)
        try:
            import comtypes.client
        except ImportError:
            comtypes = None
    if comtypes is None or not hasattr(comtypes, 'client'):
        pip = os.path.join(sys.prefix, 'Scripts', 'pip.exe')
        raise RuntimeError(
            'comtypes saknas i ArcMaps Python (%s) och gick inte att ladda fran arcmap/lib.\n'
            'Installera en gang i Kommandotolken:\n'
            '  "%s" install "comtypes<1.2"\n'
            'Saknar du adminrattigheter, lagg till --user:\n'
            '  "%s" install --user "comtypes<1.2"\n'
            'Starta sedan om ArcMap och kor verktyget igen.'
            % (sys.prefix, pip, pip))
    com = os.path.join(arcpy.GetInstallInfo()['InstallDir'], 'com')
    moduler = {}
    for namn in ('esriSystem', 'esriGeometry', 'esriDisplay', 'esriGeoDatabase',
                 'esriCarto', 'esriFramework', 'esriArcMapUI'):
        olb = os.path.join(com, namn + '.olb')
        if not os.path.isfile(olb):
            raise RuntimeError('Hittar inte %s' % olb)
        moduler[namn] = comtypes.client.GetModule(olb)
    return comtypes.client, moduler


def _hitta_lager_ao(pMap, m, lagernamn):
    """Letar upp lagret (aven i grupplager) i den oppna kartan via ArcObjects."""
    sokt = txt(lagernamn).strip().lower()

    def sok(pLayers, antal):
        for i in range(antal):
            pLayer = pLayers[i]
            if txt(pLayer.Name).strip().lower() == sokt:
                return pLayer
            try:
                pComp = pLayer.QueryInterface(m['esriCarto'].ICompositeLayer)
                traff = sok(pComp.Layer, pComp.Count)
                if traff is not None:
                    return traff
            except Exception:
                pass
        return None

    return sok(pMap.Layer, pMap.LayerCount)


def bygg_renderare(client, m):
    """Unique values-renderare pa STIL med tio kategorier."""
    esriCarto, esriDisplay = m['esriCarto'], m['esriDisplay']
    pUVR = client.CreateObject(esriCarto.UniqueValueRenderer, interface=esriCarto.IUniqueValueRenderer)
    pUVR.FieldCount = 1
    pUVR.Field[0] = FALT
    pUVR.UseDefaultSymbol = False

    for varde, etikett, streckad, bredd in KATEGORIER:
        r, g, b = KLASS_FARG[varde[0]]
        pFarg = client.CreateObject(esriDisplay.RgbColor, interface=esriDisplay.IRgbColor)
        pFarg.Red, pFarg.Green, pFarg.Blue = r, g, b
        pSym = client.CreateObject(esriDisplay.SimpleLineSymbol, interface=esriDisplay.ISimpleLineSymbol)
        pSym.Color = pFarg
        pSym.Style = esriDisplay.esriSLSDash if streckad else esriDisplay.esriSLSSolid
        pSym.Width = bredd
        pUVR.AddValue(varde, 'Bedömning', pSym.QueryInterface(esriDisplay.ISymbol))
        pUVR.Label[varde] = etikett
    return pUVR


def skapa_lyr(lagernamn=LAGER, lyr_ut=LYR_UT):
    """Satter symbologin pa lagret i kartan och sparar det som .lyr. Returnerar sokvagen."""
    logg('Laddar ArcObjects (forsta gangen tar det en stund) ...')
    client, m = _arcobjects()
    esriFramework, esriArcMapUI, esriCarto = m['esriFramework'], m['esriArcMapUI'], m['esriCarto']

    pApp = client.CreateObject(esriFramework.AppRef, interface=esriFramework.IApplication)
    pDoc = pApp.Document.QueryInterface(esriArcMapUI.IMxDocument)
    pMap = pDoc.FocusMap

    pLayer = _hitta_lager_ao(pMap, m, lagernamn)
    if pLayer is None:
        raise RuntimeError('Hittar inget lager "%s" i kartan. Kor "Skapa ledningslager" forst.' % lagernamn)
    logg('  lager: %s' % txt(pLayer.Name))

    pGFL = pLayer.QueryInterface(esriCarto.IGeoFeatureLayer)
    pGFL.Renderer = bygg_renderare(client, m).QueryInterface(esriCarto.IFeatureRenderer)
    logg('  symbologi satt: %d kategorier pa faltet %s' % (len(KATEGORIER), FALT))

    try:
        pDoc.UpdateContents()
        arcpy.RefreshTOC()
        arcpy.RefreshActiveView()
    except Exception:
        pass

    # Spara som .lyr via arcpy.mapping - samma dokument, sa renderaren foljer med
    mxd = arcpy.mapping.MapDocument('CURRENT')
    kandidater = [l for l in arcpy.mapping.ListLayers(mxd)
                  if txt(l.name).strip().lower() == txt(lagernamn).strip().lower()]
    if not kandidater:
        raise RuntimeError('arcpy.mapping hittar inte lagret "%s"' % lagernamn)
    mapp = os.path.dirname(lyr_ut)
    if mapp and not os.path.isdir(mapp):
        os.makedirs(mapp)
    if os.path.isfile(lyr_ut):
        os.remove(lyr_ut)
    kandidater[0].saveACopy(lyr_ut)
    logg('  sparad: %s' % lyr_ut)
    logg('KLART')
    return lyr_ut


if __name__ == '__main__':
    skapa_lyr(LAGER, LYR_UT)
