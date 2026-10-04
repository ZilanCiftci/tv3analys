# -*- coding: utf-8 -*-
"""
Listar alla lager i den oppna kartan med deras falt och exempelvarden (ArcMap 10.x, Python 2.7).
Anvands for att se vilka faltnamn kartan har (vattengang, dimension, material, lockniva, DUF ...)
innan Exportera GIS-data, Markprofil och Uppstroms stalls in.

Kors som verktyget "Lista falt" i tv3_verktyg.pyt, eller i ArcMaps Python-fonster:

    UTFIL = r'C:\Temp\kartans_falt.txt'
    execfile(r'C:\...\arcmap\lista_falt.py')

Utdata: en textfil (UTF-8) med, per lager: namn, datakalla, geometrityp, antal objekt, och per
falt: namn, alias, typ, langd samt upp till EXEMPEL distinkta varden ur de forsta RADER raderna.
"""
from __future__ import unicode_literals

import io
import os

try:
    import arcpy
except ImportError:          # for test utan ArcMap
    arcpy = None

UTFIL = globals().get('UTFIL') or os.path.join(os.path.expanduser('~'), 'kartans_falt.txt')
RADER = globals().get('RADER') or 2000      # rader som lases per lager for exempelvarden
EXEMPEL = globals().get('EXEMPEL') or 6     # distinkta exempelvarden per falt
BARA_LAGER = globals().get('BARA_LAGER') or []   # tomt = alla lager, annars lista med (del av) lagernamn


def _u(v):
    if v is None:
        return ''
    if isinstance(v, bytes):
        try:
            return v.decode('utf-8')
        except UnicodeDecodeError:
            return v.decode('cp1252', 'replace')
    try:
        return unicode(v)       # noqa: F821 (Python 2)
    except NameError:
        return str(v)


def lista_lager(lyr):
    """Rader med beskrivning av ett lager: falt och exempelvarden."""
    ut = []
    try:
        kalla = lyr.dataSource
    except Exception:
        kalla = '(ingen datakalla)'
    ut.append('=' * 100)
    ut.append('LAGER: %s' % _u(lyr.longName if hasattr(lyr, 'longName') else lyr.name))
    ut.append('  datakalla: %s' % _u(kalla))
    try:
        d = arcpy.Describe(lyr)
        ut.append('  geometri: %s   koordinatsystem: %s' % (_u(getattr(d, 'shapeType', '?')),
                                                            _u(getattr(getattr(d, 'spatialReference', None), 'name', '?'))))
    except Exception as e:
        ut.append('  (Describe misslyckades: %s)' % _u(e))
    try:
        antal = int(arcpy.GetCount_management(lyr).getOutput(0))
        ut.append('  antal objekt: %d' % antal)
    except Exception:
        pass
    try:
        falt = [f for f in arcpy.ListFields(lyr) if f.type not in ('Geometry',)]
    except Exception as e:
        ut.append('  (ListFields misslyckades: %s)' % _u(e))
        return ut
    namn = [f.name for f in falt]
    exempel = dict((n, []) for n in namn)
    try:
        with arcpy.da.SearchCursor(lyr, namn) as cur:
            for i, rad in enumerate(cur):
                if i >= RADER:
                    break
                for n, v in zip(namn, rad):
                    if v is None or v == '':
                        continue
                    lst = exempel[n]
                    if len(lst) < EXEMPEL and v not in lst:
                        lst.append(v)
    except Exception as e:
        ut.append('  (kunde inte lasa exempelvarden: %s)' % _u(e))
    ut.append('  %-28s %-32s %-10s %5s  %s' % ('FALT', 'ALIAS', 'TYP', 'LANGD', 'EXEMPEL'))
    for f in falt:
        ex = ', '.join(_u(v) for v in exempel.get(f.name, []))
        if len(ex) > 90:
            ex = ex[:87] + '...'
        ut.append('  %-28s %-32s %-10s %5s  %s' % (_u(f.name), _u(f.aliasName), _u(f.type),
                                                   _u(f.length) if f.type == 'String' else '', ex))
    return ut


def lista(utfil=UTFIL, bara=BARA_LAGER):
    mxd = arcpy.mapping.MapDocument('CURRENT')
    rader = ['Falt i kartan %s' % _u(mxd.filePath or '(osparad)'), '']
    n = 0
    for df in arcpy.mapping.ListDataFrames(mxd):
        rader.append('DATARAM: %s' % _u(df.name))
        for lyr in arcpy.mapping.ListLayers(mxd, '', df):
            if lyr.isGroupLayer:
                continue
            if bara and not any(_u(b).lower() in _u(lyr.longName).lower() for b in bara):
                continue
            try:
                rader.extend(lista_lager(lyr))
            except Exception as e:
                rader.append('LAGER: %s  (fel: %s)' % (_u(lyr.name), _u(e)))
            rader.append('')
            n += 1
    with io.open(utfil, 'w', encoding='utf-8') as f:
        f.write('\n'.join(rader) + '\n')
    msg = '%d lager listade till %s' % (n, utfil)
    try:
        arcpy.AddMessage(msg)
    except Exception:
        pass
    print(msg)
    return utfil


if __name__ == '__main__' or (arcpy is not None and globals().get('UTFIL')):
    if arcpy is not None:
        lista()
