# -*- coding: utf-8 -*-
"""
Batchexport av kartor (PDF) per atgardsstracka eller etapp fran ArcMap 10.x (Python 2.7 + arcpy).

For varje stracka i lagret fran "Skapa ledningslager" (urval: strackor med atgard, klass A/B,
markerade i kartan eller alla) centreras layoutens dataram pa strackan och skalan satts till
den minsta i SKALOR (1:200, 1:300, 1:400 ...) dar strackan med marginal ryms i dataramen.
Ryms den inte i nagon skala anvands den storsta och det loggas. Strackan markeras (urval i
lagret, eller ett eget markeringslager vars definitionsfraga satts), layoutens textelement
TITEL/UNDERTITEL/SKALA fylls i om de finns, och sidan exporteras med ExportToPDF. Alla sidor
kan dessutom slas ihop till en samlad PDF, och sokvagen skrivs till faltet KARTA i lagret
(hyperlank, som RAPPORT).

Skalvalet ar oberoende av sidenheter: dataramens bredd i meter vid skala S raknas som
df.extent.width * S / df.scale (i layoutvyn fyller utbredningen ramen exakt).

Testas utan ArcMap med latsas-arcpy (scratchpad/test_kartexport.py).
"""
from __future__ import unicode_literals

import os
import re

import arcpy

from skapa_ledningslager import txt, logg, hitta_lager, kalla, _tal, _mxd, TEXTTYP

SKALOR = (200, 300, 400, 500, 750, 1000, 1500, 2000)
MARGINAL_M = 10.0          # fritt utrymme runt strackan i varje riktning (meter)
URVAL = ('atgard', 'AB', 'valda', 'alla')
KLASSER = ('A', 'B')


def _falt(src):
    return dict((f.name.upper(), f.name) for f in arcpy.ListFields(src))


def _utbredning(geom):
    """(xmin, ymin, xmax, ymax) for en geometri."""
    e = geom.extent
    return e.XMin, e.YMin, e.XMax, e.YMax


def _sla_ihop(a, b):
    if a is None:
        return b
    return min(a[0], b[0]), min(a[1], b[1]), max(a[2], b[2]), max(a[3], b[3])


def dataram_matt(df):
    """(bredd, hojd) for dataramen i meter per skalenhet: ramens bredd i kartmeter vid skala
    1:S ar bredd * S. Las en gang innan skalan andras (utbredning och skala hor ihop)."""
    try:
        ram_b, ram_h, s0 = float(df.extent.width), float(df.extent.height), float(df.scale)
    except Exception:
        ram_b = ram_h = s0 = None
    if not ram_b or not ram_h or not s0:
        raise RuntimeError('Kan inte lasa dataramens utbredning och skala - oppna layoutvyn och kor igen')
    return ram_b / s0, ram_h / s0


def valj_skala(utb, ram, skalor=SKALOR, marginal=MARGINAL_M):
    """Minsta skala i skalor dar utbredningen (med marginal) ryms i dataramen (ram fran
    dataram_matt). Returnerar (skala, ryms). Ryms den inte i nagon skala returneras den
    storsta och False."""
    xmin, ymin, xmax, ymax = utb
    bredd = (xmax - xmin) + 2 * marginal
    hojd = (ymax - ymin) + 2 * marginal
    for s in sorted(skalor):
        if bredd <= ram[0] * s and hojd <= ram[1] * s:
            return int(s), True
    return int(max(skalor)), False


def _centrera(df, utb, skala, ram):
    xmin, ymin, xmax, ymax = utb
    cx, cy = (xmin + xmax) / 2.0, (ymin + ymax) / 2.0
    b, h = ram[0] * skala, ram[1] * skala
    df.extent = arcpy.Extent(cx - b / 2.0, cy - h / 2.0, cx + b / 2.0, cy + h / 2.0)
    df.scale = skala


def _satt_text(mxd, namn, text):
    """Fyller layoutens textelement med namnet namn (skiftlage spelar ingen roll), om det finns."""
    try:
        for el in arcpy.mapping.ListLayoutElements(mxd, 'TEXT_ELEMENT'):
            if txt(el.name).strip().upper() == namn.upper():
                el.text = text if text else ' '
                return True
    except Exception as e:
        logg('  kunde inte satta textelementet %s: %s' % (namn, txt(e)))
    return False


def _filnamn(post, per_etapp):
    def ren(v):
        return re.sub(r'[^A-Za-z0-9_.\-]+', '_', txt(v or '')).strip('_') or 'x'
    if per_etapp:
        return 'etapp_%02d.pdf' % int(post.get('ETAPP') or 0)
    fil = os.path.splitext(os.path.basename(txt(post.get('TV3_FIL') or '')))[0]
    return '%s_%s_%s_%s-%s.pdf' % (ren(post.get('BEDOMNING') or 'X'), ren(fil), txt(post.get('NR') or ''),
                                   ren(post.get('FRAN_BRUNN')), ren(post.get('TILL_BRUNN')))


def _titel(post, per_etapp, antal=1):
    if per_etapp:
        return 'Etapp %s (%d sträckor)' % (txt(post.get('ETAPP') or '?'), antal)
    nr = post.get('NR')
    return 'Sträcka %s%s → %s' % (('%s: ' % txt(nr)) if nr is not None else '',
                                            txt(post.get('FRAN_BRUNN') or '?'), txt(post.get('TILL_BRUNN') or '?'))


def _undertitel(post, skala, per_etapp, langd=None):
    delar = []
    if not per_etapp and post.get('BEDOMNING'):
        delar.append('Klass %s' % txt(post['BEDOMNING']))
    if post.get('ETAPP') is not None and not per_etapp:
        delar.append('Etapp %s' % txt(post['ETAPP']))
    if post.get('METOD'):
        delar.append(txt(post['METOD']))
    L = langd if langd is not None else _tal(post.get('LANGD_M'))
    if L:
        delar.append('%.0f m' % L)
    if not per_etapp:
        mat = ' '.join(x for x in (txt(post.get('MATERIAL') or ''), txt(post.get('DIMENSION') or '')) if x)
        if mat:
            delar.append(mat)
    delar.append('Skala 1:%d' % skala)
    return '  ·  '.join(delar)


def exportera(bedomda, ut_mapp, urval='atgard', skalor=SKALOR, marginal=MARGINAL_M, dpi=200,
              markeringslager=None, samlad=True, per_etapp=False, skriv_falt=True,
              kartmapp=None, bara_valda_klasser=KLASSER):
    """Exporterar en PDF per stracka (eller per etapp) till ut_mapp. Returnerar lista med
    (filnamn, skala, ryms). urval: 'atgard' (METOD ifyllt), 'AB' (BEDOMNING i A/B),
    'valda' (markerade i kartan), 'alla'. markeringslager: lager i kartan som pekar pa samma
    featureklass och vars definitionsfraga satts till den aktuella strackan (tydligare an
    urvalsfargen). kartmapp: mapp som skrivs i faltet KARTA i stallet for ut_mapp (Citrix)."""
    mxd = _mxd()
    if mxd is None:
        raise RuntimeError('Verktyget kors i ArcMap med ett oppet kartdokument')
    if not os.path.isdir(ut_mapp):
        os.makedirs(ut_mapp)
    skalor = sorted(int(s) for s in skalor if _tal(s) and float(s) > 0) or list(SKALOR)
    df = arcpy.mapping.ListDataFrames(mxd)[0]
    ram = dataram_matt(df)
    logg('Dataramen ar %.0f x %.0f m i skala 1:%d' % (ram[0] * skalor[0] if skalor else 0,
                                                     ram[1] * skalor[0] if skalor else 0, skalor[0] if skalor else 0))

    lyr = hitta_lager(bedomda) if isinstance(bedomda, (TEXTTYP, bytes)) else bedomda
    src, dq = kalla(lyr)
    falt = _falt(src)
    for f in ('FRAN_BRUNN', 'TILL_BRUNN'):
        if f not in falt:
            raise RuntimeError('Lagret saknar faltet %s - valj lagret fran "Skapa ledningslager"' % f)
    oidfalt = arcpy.Describe(src).OIDFieldName
    las = ['OID@', 'SHAPE@'] + [falt[f] for f in ('FRAN_BRUNN', 'TILL_BRUNN', 'BEDOMNING', 'ETAPP', 'METOD',
                                                  'LANGD_M', 'MATERIAL', 'DIMENSION', 'NR', 'TV3_FIL') if f in falt]
    namn = ['OID@', 'SHAPE@'] + [f for f in ('FRAN_BRUNN', 'TILL_BRUNN', 'BEDOMNING', 'ETAPP', 'METOD',
                                             'LANGD_M', 'MATERIAL', 'DIMENSION', 'NR', 'TV3_FIL') if f in falt]
    valda_oid = None
    if urval == 'valda':
        try:
            valda_oid = set(lyr.getSelectionSet() or [])
        except Exception:
            valda_oid = set()
        if not valda_oid:
            raise RuntimeError('Inga strackor ar markerade i kartan')

    poster = []
    with arcpy.da.SearchCursor(src, las, dq) as mark:
        for rad in mark:
            post = dict(zip(namn, rad))
            if post['SHAPE@'] is None:
                continue
            if urval == 'atgard' and not txt(post.get('METOD') or '').strip():
                continue
            if urval == 'AB' and txt(post.get('BEDOMNING') or '').strip().upper()[:1] not in bara_valda_klasser:
                continue
            if urval == 'valda' and post['OID@'] not in valda_oid:
                continue
            poster.append(post)
    logg('%d strackor att exportera (urval: %s)' % (len(poster), urval))
    if not poster:
        return []

    # Grupper: en per stracka, eller en per etapp
    grupper = []
    if per_etapp:
        per = {}
        for p in poster:
            per.setdefault(p.get('ETAPP'), []).append(p)
        for et in sorted(per, key=lambda v: (v is None, v)):
            if et is None:
                logg('  %d strackor utan etapp hoppas over i etapplaget' % len(per[et]))
                continue
            grupper.append(per[et])
    else:
        grupper = [[p] for p in poster]

    # Markeringslager: egen kopia av lagret med tydlig symbologi; annars urval i lagret
    mark_lyr = None
    if markeringslager:
        try:
            mark_lyr = hitta_lager(markeringslager)
        except Exception as e:
            logg('  markeringslagret hittades inte (%s) - urval anvands i stallet' % txt(e))

    def markera(oids):
        where = '%s IN (%s)' % (arcpy.AddFieldDelimiters(src, oidfalt), ','.join(str(o) for o in oids))
        if mark_lyr is not None:
            mark_lyr.definitionQuery = where
            try:
                mark_lyr.visible = True
            except Exception:
                pass
        else:
            arcpy.SelectLayerByAttribute_management(lyr, 'NEW_SELECTION', where)

    gammal_dq = getattr(mark_lyr, 'definitionQuery', None) if mark_lyr is not None else None
    ut = []
    filer = []
    karta_per_oid = {}
    try:
        for grupp in grupper:
            utb = None
            for p in grupp:
                utb = _sla_ihop(utb, _utbredning(p['SHAPE@']))
            skala, ryms = valj_skala(utb, ram, skalor, marginal)
            _centrera(df, utb, skala, ram)
            markera([p['OID@'] for p in grupp])
            p0 = grupp[0]
            langd = sum(_tal(p.get('LANGD_M')) or 0 for p in grupp) if per_etapp else None
            _satt_text(mxd, 'TITEL', _titel(p0, per_etapp, len(grupp)))
            _satt_text(mxd, 'UNDERTITEL', _undertitel(p0, skala, per_etapp, langd))
            _satt_text(mxd, 'SKALA', 'Skala 1:%d' % skala)
            try:
                arcpy.RefreshActiveView()
            except Exception:
                pass
            fil = os.path.join(ut_mapp, _filnamn(p0, per_etapp))
            arcpy.mapping.ExportToPDF(mxd, fil, 'PAGE_LAYOUT', resolution=int(dpi),
                                      image_quality='BEST', georef_info=True)
            filer.append(fil)
            ut.append((fil, skala, ryms))
            for p in grupp:
                karta_per_oid[p['OID@']] = fil if not kartmapp else os.path.join(txt(kartmapp), os.path.basename(fil))
            logg('  %s  1:%d%s' % (os.path.basename(fil), skala, '' if ryms else '  (ryms inte - storsta skalan)'))
    finally:
        # Aterstall markering
        try:
            if mark_lyr is not None:
                mark_lyr.definitionQuery = gammal_dq or ''
            else:
                arcpy.SelectLayerByAttribute_management(lyr, 'CLEAR_SELECTION')
        except Exception:
            pass

    if samlad and filer:
        samlad_fil = os.path.join(ut_mapp, 'kartor_etapper.pdf' if per_etapp else 'kartor_strackor.pdf')
        try:
            if os.path.exists(samlad_fil):
                os.remove(samlad_fil)
            pdf = arcpy.mapping.PDFDocumentCreate(samlad_fil)
            for f in filer:
                pdf.appendPages(f)
            pdf.saveAndClose()
            logg('  samlad PDF: %s (%d sidor)' % (samlad_fil, len(filer)))
        except Exception as e:
            logg('  kunde inte skapa samlad PDF: %s' % txt(e))

    if skriv_falt and karta_per_oid:
        try:
            if 'KARTA' not in falt:
                arcpy.AddField_management(src, 'KARTA', 'TEXT', field_length=254, field_alias='Karta (PDF)')
                falt['KARTA'] = 'KARTA'
            n = 0
            with arcpy.da.UpdateCursor(src, ['OID@', falt['KARTA']]) as mark:
                for rad in mark:
                    if rad[0] in karta_per_oid:
                        rad[1] = karta_per_oid[rad[0]][:254]
                        mark.updateRow(rad)
                        n += 1
            logg('  faltet KARTA skrivet for %d strackor (hyperlank: Display > Support Hyperlinks using field)' % n)
        except Exception as e:
            logg('  kunde inte skriva faltet KARTA: %s' % txt(e))

    ej = [f for f, s, r in ut if not r]
    if ej:
        logg('  %d kartor rymdes inte i storsta skalan 1:%d - korta ned sidan eller lagg till storre skalor' % (len(ej), max(skalor)))
    logg('KLART')
    return ut
