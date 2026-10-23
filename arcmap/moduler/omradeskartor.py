# -*- coding: utf-8 -*-
"""
Kartor for omradet: PDF-kartor ur ledningsinventeringen (ledningsresultat.json fran ledning_analys.py)
i ArcMap 10.x (Python 2.7 + arcpy).

  * Omradeskarta - en per polygon i omradeslagret, i den storsta skala dar hela omradet ryms pa sidan
    (liggande eller staende mall, den som ger storst skala). En informationsruta (textelementet INFO
    i mallen) visar ledningslangden per ledningstyp, totalt och den totala schaktlangden. Med kartblad
    visas bladindelningen med bladnummer.
  * Schaktgrupp - en karta per schaktgrupp (ledningar i samma schakt), gruppens ledningar markerade
    med markeringssymbolen (markering.lyr, bred gul genomskinlig linje) och etiketter med ledningens id.
  * Kartblad - omradet delat i blad i en fast skala (t.ex. 1:1000) med lite overlapp, varje blad med
    ledningarnas id som etiketter. Bladnumren skrivs till ledningarna (falt KARTBLAD och KARTA med
    sokvagen till bladets PDF) och till kartblad.csv bredvid ledningsresultat.json - ledning_analys.py
    lagger dem i Excel-filen nar den kors igen.

Informationsrutan: finns inget textelement med namnet INFO i mallen skapas en kopia av textelementet
TITEL (annars det forsta textelementet) till hoger om kartan och dataramen gors smalare pa
omradeskartorna. Etiketterna tar faltet LEDN_ID (verktyget Uppdatera inventeringslager).

Mallarna sparas inte; i den oppna kartan aterstalls allt (tillfalliga lager, etiketter, dataramens
storlek) efter exporten. Testas med latsas-arcpy (scratchpad/test_omradeskartor.py).
"""
from __future__ import unicode_literals

import io
import math
import os
import re

import arcpy

from skapa_ledningslager import txt, logg, hitta_lager, kalla, _tal, _mxd
import kartexport as ke
import omrade_export as oe
import inventeringslager as il

SKALOR_OMRADE = (500, 750, 1000, 1500, 2000, 2500, 3000, 4000, 5000, 7500, 10000, 15000, 20000, 25000, 50000)
SKALOR_GRUPP = (200, 300, 400, 500, 750, 1000, 1500, 2000, 2500, 3000, 4000, 5000, 7500, 10000)
MARGINAL_OMRADE_M = 20.0      # fritt utrymme runt omradet
MARGINAL_GRUPP_M = 15.0       # fritt utrymme runt schaktgruppen
KARTBLAD_SKALA = 1000
KARTBLAD_OVERLAPP = 0.05      # andel av bladets bredd/hojd som grannbladen overlappar
INFO_NAMN = 'INFO'            # textelement i mallen for informationsrutan
INFO_BREDD_CM = 6.5           # bredd pa rutan nar den skapas av verktyget (dataramen gors smalare)
INFO_TEXTSTORLEK = 9
ETIKETT_FALT = 'LEDN_ID'      # ledningens id (Uppdatera inventeringslager)
MARKERING_NAMN = 'Aktuell schaktgrupp (export)'
KARTBLAD_NAMN = 'Kartblad (export)'
KARTBLAD_LYR = os.path.join(ke.ARCMAP_MAPP, 'kartblad.lyr')   # valfri symbol for bladgranserna
TMP_GRUPPER = 'tv3_schaktgrupper'
TMP_KARTBLAD = 'tv3_kartblad'
KARTBLAD_CSV = 'kartblad.csv'
SYSTEM_NAMN = {'S': 'Spillvatten', 'D': 'Dagvatten', 'K': 'Kombinerat', 'V': 'Vatten'}
SYSTEM_ORDNING = 'SDKV'
FALT_KARTOR = [('KARTBLAD', 'TEXT', 100, 'Kartblad'), ('KARTA', 'TEXT', 254, 'Kartblad (PDF)'),
               ('GRUPPKARTA', 'TEXT', 254, 'Schaktgrupp (PDF)')]

_AO = {'å': 'a', 'ä': 'a', 'ö': 'o', 'Å': 'A', 'Ä': 'A', 'Ö': 'O',
       'é': 'e', 'ü': 'u'}


# ------------------------------------------------------------------ text och tal
def _m(v):
    """1234.5 -> '1 235 m'."""
    return ('{:,.0f}'.format(float(v)).replace(',', ' ')) + ' m'


def _ren(namn):
    """Filnamnsdel: aao for åäö, annat an A-Z/0-9/-/_ blir _."""
    s = ''.join(_AO.get(c, c) for c in txt(namn))
    return re.sub(r'[^A-Za-z0-9_\-]+', '_', s).strip('_') or 'x'


def info_text(namn, ledningar):
    """Informationsrutan for ett omrade: ledningslangd per ledningstyp, totalt och schaktlangd."""
    per = {}
    typer = {}
    for l in ledningar:
        s = txt(l.get('system') or '').strip().upper()[:1] or '?'
        per[s] = per.get(s, 0.0) + (_tal(l.get('langd_m')) or 0.0)
        t = txt(l.get('ledningstyp') or '').strip()
        if t:
            typer.setdefault(s, [])
            if t not in typer[s]:
                typer[s].append(t)
    rader = ['Område %s' % txt(namn), '', 'Ledningslängd']
    for s in sorted(per, key=lambda k: (SYSTEM_ORDNING.find(k) % 9, k)):
        etikett = SYSTEM_NAMN.get(s, 'Okänd typ' if s == '?' else s)
        if typer.get(s):
            etikett += ' (%s)' % ', '.join(sorted(typer[s]))
        rader.append('%s: %s' % (etikett, _m(per[s])))
    rader.append('Totalt: %s' % _m(sum(per.values())))
    schakt = [_tal(l.get('schakt_langd_m')) for l in ledningar]
    rader.append('')
    if ledningar and all(v is not None for v in schakt):
        rader.append('Schaktlängd totalt: %s' % _m(sum(schakt)))
    else:
        rader.append('Schaktlängd: okänd (kör ledning_analys.py igen)')
    return '\n'.join(rader)


# ------------------------------------------------------------------ kartblad
def rutnat(utb, blad_b, blad_h, overlapp=KARTBLAD_OVERLAPP):
    """Bladindelning som tacker utbredningen (xmin, ymin, xmax, ymax), centrerad over den.
    Returnerar [(rad, kolumn, (xmin, ymin, xmax, ymax))] radvis uppifran, vanster till hoger."""
    xmin, ymin, xmax, ymax = utb
    sx, sy = blad_b * (1 - overlapp), blad_h * (1 - overlapp)
    W, H = xmax - xmin, ymax - ymin
    nx = 1 if W <= blad_b else int(math.ceil((W - blad_b) / sx - 1e-9)) + 1
    ny = 1 if H <= blad_h else int(math.ceil((H - blad_h) / sy - 1e-9)) + 1
    tot_b, tot_h = blad_b + (nx - 1) * sx, blad_h + (ny - 1) * sy
    x0 = (xmin + xmax) / 2.0 - tot_b / 2.0
    ytopp = (ymin + ymax) / 2.0 + tot_h / 2.0
    ut = []
    for r in range(ny):
        for c in range(nx):
            x = x0 + c * sx
            y = ytopp - r * sy
            ut.append((r, c, (x, y - blad_h, x + blad_b, y)))
    return ut


def grannar(blad):
    """{bladnummer: [grannarnas nummer]} for blad = [(nr, rad, kolumn, rekt)]."""
    pos = dict(((r, c), nr) for nr, r, c, _ in blad)
    ut = {}
    for nr, r, c, _ in blad:
        ut[nr] = sorted(pos[(r + dr, c + dc)] for dr in (-1, 0, 1) for dc in (-1, 0, 1)
                        if (dr or dc) and (r + dr, c + dc) in pos)
    return ut


def _polygon(rekt, sr=None):
    x0, y0, x1, y1 = rekt
    arr = arcpy.Array([arcpy.Point(x0, y0), arcpy.Point(x0, y1), arcpy.Point(x1, y1), arcpy.Point(x1, y0),
                       arcpy.Point(x0, y0)])
    return arcpy.Polygon(arr, sr)


def _ramlinje(rekt, sr=None):
    x0, y0, x1, y1 = rekt
    arr = arcpy.Array([arcpy.Point(x0, y0), arcpy.Point(x0, y1), arcpy.Point(x1, y1), arcpy.Point(x1, y0),
                       arcpy.Point(x0, y0)])
    return arcpy.Polyline(arr, sr)


def _utb(geom):
    return ke._utbredning(geom)


def _skar(a, b):
    """Overlappar tva utbredningar (xmin, ymin, xmax, ymax)?"""
    return a[0] <= b[2] and b[0] <= a[2] and a[1] <= b[3] and b[1] <= a[3]


def _traffar(rekt, poly, geometrier, sr):
    """Behovs bladet: overlappar det omradet eller nagon av omradets ledningar?"""
    if not _skar(rekt, _utb(poly)) and not any(_skar(rekt, _utb(g)) for g in geometrier):
        return False
    rp = _polygon(rekt, sr)
    if _skar(rekt, _utb(poly)) and not poly.disjoint(rp):
        return True
    return any(_skar(rekt, _utb(g)) and not rp.disjoint(g) for g in geometrier)


# ------------------------------------------------------------------ layout
class Sida(object):
    """Ett kartdokument att exportera ur (mall eller den oppna kartan) med dataram, informationsruta,
    tillfalliga lager och etiketter som aterstalls efterat."""

    def __init__(self, mxd, namn, oppen=None, kopiera_synliga=True):
        self.mxd = mxd
        self.namn = namn
        self.tillagda = []
        self._etiketter = []
        self.info_el = None
        self._klon = None
        self._df_orig = None
        try:
            if txt(mxd.activeView).upper() != 'PAGE_LAYOUT':
                mxd.activeView = 'PAGE_LAYOUT'
        except Exception:
            pass
        self.df = ke._huvudram(mxd)
        if oppen is not None and mxd is not oppen and kopiera_synliga:
            f = ke.Layout._kopiera_synliga
            getattr(f, '__func__', f)(self, oppen)
        self.ram = ke.dataram_matt(self.df)
        self.ram_info = self.ram
        self.liggande = self.ram[0] >= self.ram[1]
        self._forbered_info()

    # --- textelement
    def _textelement(self, namn=None):
        try:
            alla = arcpy.mapping.ListLayoutElements(self.mxd, 'TEXT_ELEMENT')
        except Exception:
            return None
        if namn is None:
            return alla[0] if alla else None
        for el in alla:
            if txt(el.name).strip().upper() == namn.upper():
                return el
        return None

    def _forbered_info(self):
        el = self._textelement(INFO_NAMN)
        if el is not None:
            self.info_el = el
            return
        kalla_el = self._textelement('TITEL') or self._textelement()
        if kalla_el is None:
            logg('  %s: inget textelement att gora informationsruta av - lagg in ett textelement med namnet '
                 '%s i mallen (Properties > Size and Position > Element Name)' % (self.namn, INFO_NAMN))
            return
        try:
            ps = self.mxd.pageSize
            storst = max(float(ps.width), float(ps.height))
            enhet = 10.0 if storst > 200 else (1.0 if storst > 25 else 1 / 2.54)    # mm, cm eller tum
            df = self.df
            x, y, w, h = (float(df.elementPositionX), float(df.elementPositionY), float(df.elementWidth),
                          float(df.elementHeight))
            bredd, mellan = INFO_BREDD_CM * enhet, 0.4 * enhet
            if bredd + mellan > 0.5 * w:
                logg('  %s: dataramen ar for smal for en informationsruta - lagg in textelementet %s i mallen'
                     % (self.namn, INFO_NAMN))
                return
            klon = kalla_el.clone('_info')
            try:
                klon.name = INFO_NAMN
            except Exception:
                pass
            try:
                klon.fontSize = INFO_TEXTSTORLEK
            except Exception:
                pass
            klon.text = ' '
            self._klon = self.info_el = klon
            self._df_orig = (x, y, w, h)
            self._info_x = x + w - bredd
            self._info_topp = y + h
            self._smal = w - bredd - mellan
            self.ram_info = (self.ram[0] * self._smal / w, self.ram[1])
            self._gom_klon()
            logg('  %s: informationsrutan laggs till hoger om kartan (inget textelement %s i mallen)'
                 % (self.namn, INFO_NAMN))
        except Exception as e:
            self._klon = self.info_el = None
            self.ram_info = self.ram
            logg('  %s: kunde inte skapa informationsrutan (%s) - lagg in ett textelement med namnet %s i mallen'
                 % (self.namn, txt(e), INFO_NAMN))

    def _gom_klon(self):
        x, y, w, h = self._df_orig
        try:
            self._klon.text = ' '
            self._klon.elementPositionX = x - 10 * w          # utanfor sidan - skrivs inte ut
        except Exception:
            pass

    def visa_info(self, text):
        """Informationsrutan med text (None = ingen ruta). Anropas fore centreringen - dataramen kan
        andra bredd."""
        if self.info_el is None:
            return
        if self._klon is None:
            try:
                self.info_el.text = text or ' '
            except Exception as e:
                logg('  kunde inte fylla %s: %s' % (INFO_NAMN, txt(e)))
            return
        x, y, w, h = self._df_orig
        try:
            if text:
                self.df.elementWidth = self._smal
                self._klon.text = text
                self._klon.elementPositionX = self._info_x
                self._klon.elementPositionY = self._info_topp - float(self._klon.elementHeight)
            else:
                self.df.elementWidth = w
                self._gom_klon()
        except Exception as e:
            logg('  kunde inte placera informationsrutan: %s' % txt(e))

    def har_info(self):
        return self.info_el is not None

    def ram_for(self, info):
        return self.ram_info if (info and self._klon is not None) else self.ram

    # --- tillfalliga lager
    def lagg_in(self, fc, namn, lyr_fil=None, transparens=None):
        """Lagger ett tillfalligt lager (featureklassen fc) overst i dataramen; tas bort i aterstall."""
        try:
            ny = arcpy.mapping.Layer(fc)
            ny.name = namn
            if lyr_fil:
                try:
                    arcpy.ApplySymbologyFromLayer_management(ny, lyr_fil)
                except Exception as e:
                    logg('  kunde inte satta symbolen ur %s: %s' % (os.path.basename(txt(lyr_fil)), txt(e)))
            arcpy.mapping.AddLayer(self.df, ny, 'TOP')
        except Exception as e:
            logg('  kunde inte lagga in lagret "%s" i %s: %s' % (namn, self.namn, txt(e)))
            return None
        lag = None
        try:
            for l in arcpy.mapping.ListLayers(self.mxd, '', self.df):
                if txt(l.name).strip().lower() == namn.lower() and not getattr(l, 'isGroupLayer', False):
                    lag = l
                    break
        except Exception:
            pass
        if lag is None:
            logg('  OBS: lagret "%s" hittades inte i %s efter att det lagts in - ta bort det for hand' % (namn, self.namn))
            return None
        if lyr_fil:
            try:
                arcpy.mapping.UpdateLayer(self.df, lag, arcpy.mapping.Layer(lyr_fil), True)
            except Exception:
                pass
        if transparens is not None:
            try:
                lag.transparency = int(transparens)
            except Exception:
                pass
        self.tillagda.append(lag)
        return lag

    # --- etiketter
    def etiketter(self, lagernamn, falt):
        """Satter etiketter med faltet falt pa lagren (namn) - aterstalls i etiketter_av/aterstall."""
        for namn in lagernamn:
            l = ke._lager_i_doc(self.mxd, namn)
            if l is None:
                continue
            try:
                if not l.supports('LABELCLASSES'):
                    continue
                finns = dict((f.name.upper(), f.name) for f in arcpy.ListFields(kalla(l)[0]))
                if falt.upper() not in finns:
                    continue
                klasser = list(l.labelClasses)
                self._etiketter.append((l, l.showLabels, [(k, k.expression, k.showClassLabels) for k in klasser]))
                for i, k in enumerate(klasser):
                    k.expression = '[%s]' % finns[falt.upper()]
                    k.showClassLabels = (i == 0)
                l.showLabels = True
            except Exception as e:
                logg('  kunde inte satta etiketter pa %s: %s' % (txt(namn), txt(e)))

    def etiketter_av(self):
        for l, visa, klasser in reversed(self._etiketter):
            try:
                for k, uttryck, visa_k in klasser:
                    k.expression = uttryck
                    k.showClassLabels = visa_k
                l.showLabels = visa
            except Exception:
                pass
        self._etiketter = []

    def aterstall(self):
        self.etiketter_av()
        for lag in self.tillagda:
            try:
                arcpy.mapping.RemoveLayer(self.df, lag)
            except Exception as e:
                logg('  kunde inte ta bort det tillfalliga lagret "%s" ur %s: %s' % (txt(lag.name), self.namn, txt(e)))
        self.tillagda = []
        if self._klon is not None:
            try:
                x, y, w, h = self._df_orig
                self.df.elementWidth = w
                self._klon.delete()
            except Exception:
                pass
            self._klon = None

    def stang(self):
        self.df = self.mxd = self.info_el = None


def valj_sida(sidor, utb, skalor, marginal, info=False):
    """(sida, skala, ryms): den sida som ger minsta skala (vid lika den dar utbredningen fyller ramen
    bast; ryms den inte: minst overskjutning)."""
    bast = None
    for s in sidor:
        ram = s.ram_for(info)
        skala, ryms = ke.valj_skala(utb, ram, skalor, marginal)
        b = (utb[2] - utb[0]) + 2 * marginal
        h = (utb[3] - utb[1]) + 2 * marginal
        fyll = max(b / (ram[0] * skala), h / (ram[1] * skala))
        nyckel = (not ryms, skala, fyll if not ryms else -fyll)
        if bast is None or nyckel < bast[0]:
            bast = (nyckel, s, skala, ryms)
    return bast[1], bast[2], bast[3]


# ------------------------------------------------------------------ hjalp
def _tmp_fc(namn, typ, sr, falt):
    ws = arcpy.env.scratchGDB
    fc = os.path.join(ws, namn)
    if arcpy.Exists(fc):
        arcpy.Delete_management(fc)
    arcpy.CreateFeatureclass_management(ws, namn, typ, spatial_reference=sr)
    for f, t, n in falt:
        if n:
            arcpy.AddField_management(fc, f, t, field_length=n)
        else:
            arcpy.AddField_management(fc, f, t)
    return fc


def _exportera_pdf(sida, fil, dpi, oppen):
    if sida.mxd is oppen:
        try:
            arcpy.RefreshActiveView()
        except Exception:
            pass
    if os.path.exists(fil):
        os.remove(fil)
    arcpy.mapping.ExportToPDF(sida.mxd, fil, 'PAGE_LAYOUT', resolution=int(dpi), image_quality='BEST',
                              georef_info=True)


def _samla(filer, ut):
    if not filer:
        return
    try:
        if os.path.exists(ut):
            os.remove(ut)
        pdf = arcpy.mapping.PDFDocumentCreate(ut)
        for f in filer:
            pdf.appendPages(f)
        pdf.saveAndClose()
        logg('  samlad PDF: %s (%d sidor)' % (ut, len(filer)))
    except Exception as e:
        logg('  kunde inte skapa samlad PDF %s: %s' % (os.path.basename(ut), txt(e)))


def _sql_text(v):
    return "'%s'" % txt(v).replace("'", "''")


def _skriv_csv(path, rader):
    with io.open(path, 'w', encoding='utf-8-sig', newline='') as f:
        f.write('id;omrade;kartblad;karta;schaktgrupp;gruppkarta\r\n')
        for r in rader:
            f.write(';'.join(txt(v if v is not None else '').replace(';', ',') for v in r) + '\r\n')


# ------------------------------------------------------------------ export
def exportera(resultat_json, omradeslager, ut_mapp, namnfalt=None, bara_valda=True, omraden=True,
              schaktgrupper=True, kartblad=True, kartblad_skala=KARTBLAD_SKALA, mall_liggande=None,
              mall_staende=None, kopiera_synliga=True, markering_lyr=None, transparens=ke.MARKERING_TRANSPARENS,
              dpi=200, samlad=True, skriv_falt=True, kartmapp=None, kartblad_lyr=None):
    """Exporterar omradeskartor, kartor per schaktgrupp och kartblad (se modulens beskrivning).
    Returnerar {'omraden': [filer], 'schaktgrupper': [filer], 'kartblad': {omrade: [filer]}}."""
    oppen = _mxd()
    if oppen is None:
        raise RuntimeError('Verktyget kors i ArcMap med ett oppet kartdokument')
    if not (omraden or schaktgrupper or kartblad):
        raise RuntimeError('Valj minst en sorts karta')
    if not os.path.isdir(ut_mapp):
        os.makedirs(ut_mapp)
    data = il.las_resultat(resultat_json)
    ledningar = data['ledningar']
    per = il.per_objekt(ledningar)
    if not per:
        raise RuntimeError('ledningsresultat.json har inga ledningar med lager och objekt-id')
    if markering_lyr and not os.path.isfile(txt(markering_lyr)):
        logg('Markeringssymbolen finns inte: %s' % txt(markering_lyr))
        markering_lyr = None
    if schaktgrupper and not markering_lyr:
        logg('Ingen markeringssymbol (.lyr) - schaktgruppen visas med ett tillfalligt lager i ArcMaps standardsymbol. '
             'Skapa symbolen en gang: hogerklicka ett ledningslager > Properties > Symbology > Single symbol, gul linje '
             'ca 8 pt > OK, hogerklicka lagret > Save As Layer File... som %s (genomskinligheten satts av verktyget).'
             % txt(ke.MARKERING_LYR))
    if kartblad_lyr is None and os.path.isfile(KARTBLAD_LYR):
        kartblad_lyr = KARTBLAD_LYR
    elif kartblad_lyr and not os.path.isfile(txt(kartblad_lyr)):
        kartblad_lyr = None

    # --- sidor (mallar eller den oppna kartan)
    sidor = []
    mallar = [(m, n) for m, n in ((mall_liggande, 'liggande'), (mall_staende, 'staende')) if m]
    try:
        if mallar:
            for mall, namn in mallar:
                if not os.path.isfile(txt(mall)):
                    raise RuntimeError('Mallen finns inte: %s' % txt(mall))
                sidor.append(Sida(arcpy.mapping.MapDocument(txt(mall)),
                                  '%s (%s)' % (namn, os.path.basename(txt(mall))), oppen, kopiera_synliga))
        else:
            sidor.append(Sida(oppen, 'oppna kartan'))
        for s in sidor:
            logg('%s: dataramen ar %.0f x %.0f m i skala 1:1000' % (s.namn, s.ram[0] * 1000, s.ram[1] * 1000))
        sr = sidor[0].df.spatialReference

        # --- geometrier for objekten i resultatet (i dataramens koordinatsystem)
        geom = {}
        lagernamn = sorted(set(k[0] for k in per))
        for namn in lagernamn:
            try:
                src, dq = kalla(hitta_lager(namn))
            except Exception as e:
                logg('  OBS: lagret "%s" finns inte i kartan (%s) - dess ledningar kommer inte med' % (namn, txt(e)))
                continue
            oids = set(oid for l, oid in per if l == namn)
            with arcpy.da.SearchCursor(src, ['OID@', 'SHAPE@'], dq, spatial_reference=sr) as cur:
                for oid, g in cur:
                    if oid in oids and g is not None:
                        geom[(namn, oid)] = g
            saknas = len(oids) - sum(1 for k in geom if k[0] == namn)
            if saknas:
                logg('  OBS: %d objekt i resultatet finns inte i %s (har lagret exporterats om?)' % (saknas, namn))
        logg('%d av %d ledningsobjekt hittades i kartan' % (len(geom), len(per)))
        if not geom:
            raise RuntimeError('Inga av ledningarna i resultatet finns i kartans lager - lagg in de lokala lagren '
                               'som anvandes i Ledningsinventering')

        # --- omraden och vilka ledningar som ligger i dem
        polys = [(n, oe._i_sr(g, sr)) for n, g in oe._polygoner(hitta_lager(omradeslager), namnfalt, bara_valda)]
        if not polys:
            raise RuntimeError('Inga polygoner i omradeslagret')
        omr_obj = dict((n, []) for n, _ in polys)
        for k, delar in per.items():
            namn_i = set()
            for l in delar:
                for n in txt(l.get('omrade') or '').split(';'):
                    if n.strip():
                        namn_i.add(n.strip())
            traff = [n for n, _ in polys if n in namn_i]
            if not traff and k in geom:            # inget omradesnamn (eller annat omradeslager): geometrin avgor
                traff = [n for n, p in polys if _skar(_utb(p), _utb(geom[k])) and not p.disjoint(geom[k])]
            for n in traff:
                omr_obj[n].append(k)

        ut = {'omraden': [], 'schaktgrupper': [], 'kartblad': {}}
        kartblad_per_obj = {}      # objekt -> [(omrade, nr, fil)]
        gruppkarta = {}            # schaktgrupp -> fil

        # --- kartblad: rutnat per omrade (i forvag, sa att omradeskartan kan visa indelningen)
        blad_per_omr = {}
        if kartblad:
            for namn, poly in polys:
                bast = None
                for s in sidor:
                    b, h = s.ram[0] * kartblad_skala, s.ram[1] * kartblad_skala
                    kvar = []
                    for r, c, rekt in rutnat(_utb(poly), b, h):
                        if _traffar(rekt, poly, [geom[k] for k in omr_obj[namn] if k in geom], sr):
                            kvar.append((r, c, rekt))
                    if bast is None or len(kvar) < len(bast[1]) or (len(kvar) == len(bast[1]) and s.liggande):
                        bast = (s, kvar)
                s, kvar = bast
                blad_per_omr[namn] = (s, [(i + 1, r, c, rekt) for i, (r, c, rekt) in enumerate(kvar)])
                logg('Kartblad %s: %d blad i 1:%d (%s)' % (namn, len(kvar), kartblad_skala, s.namn))
            fc_blad = _tmp_fc(TMP_KARTBLAD, 'POLYLINE', sr, [('OMRADE', 'TEXT', 200), ('BLAD', 'LONG', None),
                                                              ('ETIKETT', 'TEXT', 20)])
            with arcpy.da.InsertCursor(fc_blad, ['SHAPE@', 'OMRADE', 'BLAD', 'ETIKETT']) as cur:
                for namn, (s, blad) in blad_per_omr.items():
                    for nr, r, c, rekt in blad:
                        cur.insertRow([_ramlinje(rekt, sr), namn[:200], nr, '%d' % nr])
            for s in sidor:
                lag = s.lagg_in(fc_blad, KARTBLAD_NAMN, kartblad_lyr)
                if lag is not None:
                    s.bladlager = lag
                    try:
                        lag.visible = False
                        if lag.supports('LABELCLASSES'):
                            for k in lag.labelClasses:
                                k.expression = '[ETIKETT]'
                                k.showClassLabels = True
                            lag.showLabels = True
                    except Exception:
                        pass

        if schaktgrupper:
            grupper = {}
            for k, delar in per.items():
                for l in delar:
                    g = l.get('schaktgrupp')
                    if g is not None and k in geom:
                        grupper.setdefault(int(g), set()).add(k)
            if grupper:
                fc_grp = _tmp_fc(TMP_GRUPPER, 'POLYLINE', sr, [('GRUPP', 'LONG', None)])
                with arcpy.da.InsertCursor(fc_grp, ['SHAPE@', 'GRUPP']) as cur:
                    for g, nycklar in grupper.items():
                        for k in nycklar:
                            cur.insertRow([geom[k], g])
                for s in sidor:
                    lag = s.lagg_in(fc_grp, MARKERING_NAMN, markering_lyr, transparens if markering_lyr else None)
                    if lag is not None:
                        s.markeringslager = lag
                        try:
                            lag.visible = False
                        except Exception:
                            pass
        else:
            grupper = {}

        def stall_in(sida, blad_dq=None, grupp=None, etiketter=False):
            """Synlighet for de tillfalliga lagren och etiketterna infor en export."""
            lag = getattr(sida, 'bladlager', None)
            if lag is not None:
                try:
                    lag.definitionQuery = blad_dq or ''
                    lag.visible = blad_dq is not None
                except Exception:
                    pass
            lag = getattr(sida, 'markeringslager', None)
            if lag is not None:
                try:
                    lag.definitionQuery = 'GRUPP = %d' % grupp if grupp is not None else ''
                    lag.visible = grupp is not None
                except Exception:
                    pass
            sida.etiketter_av()
            if etiketter:
                sida.etiketter(lagernamn, ETIKETT_FALT)

        harid = False
        for namn in lagernamn:
            try:
                harid = harid or ETIKETT_FALT in [f.name.upper() for f in arcpy.ListFields(kalla(hitta_lager(namn))[0])]
            except Exception:
                pass
        if (schaktgrupper or kartblad) and not harid:
            logg('  OBS: ledningslagren saknar faltet %s - kor Uppdatera inventeringslager forst sa far kartorna '
                 'etiketter med ledningens id' % ETIKETT_FALT)

        # --- omradeskartor
        if omraden:
            logg('Omradeskartor')
            for namn, poly in polys:
                leds = [l for k in omr_obj[namn] for l in per[k]]
                text = info_text(namn, leds)
                sida, skala, ryms = valj_sida(sidor, _utb(poly), SKALOR_OMRADE, MARGINAL_OMRADE_M, info=True)
                for s in sidor:
                    s.visa_info(None)
                sida.visa_info(text)
                stall_in(sida, blad_dq=('OMRADE = %s' % _sql_text(namn)) if kartblad else None)
                ke._centrera(sida.df, _utb(poly), skala, sida.ram_for(True))
                ke._satt_text(sida.mxd, 'TITEL', 'Område %s' % namn)
                ke._satt_text(sida.mxd, 'UNDERTITEL', 'Översikt' + (' med kartblad' if kartblad else ''))
                ke._satt_text(sida.mxd, 'SKALA', 'Skala 1:%d' % skala)
                if not sida.har_info():
                    logg('  (ingen informationsruta i %s)' % sida.namn)
                fil = os.path.join(ut_mapp, 'omrade_%s.pdf' % _ren(namn))
                _exportera_pdf(sida, fil, dpi, oppen)
                ut['omraden'].append(fil)
                logg('  %s  1:%d  %d ledningar%s' % (os.path.basename(fil), skala, len(leds),
                                                     '' if ryms else '  (ryms inte - storsta skalan)'))
            for s in sidor:
                s.visa_info(None)
            if samlad and len(ut['omraden']) > 1:
                _samla(ut['omraden'], os.path.join(ut_mapp, 'kartor_omraden.pdf'))

        # --- kartblad
        if kartblad:
            logg('Kartblad')
            for namn, poly in polys:
                sida, blad = blad_per_omr[namn]
                gr = grannar(blad)
                filer = []
                if omraden:
                    filer.append(os.path.join(ut_mapp, 'omrade_%s.pdf' % _ren(namn)))
                for nr, r, c, rekt in blad:
                    stall_in(sida, blad_dq='OMRADE = %s' % _sql_text(namn), etiketter=True)
                    ke._centrera(sida.df, rekt, kartblad_skala, sida.ram)
                    ke._satt_text(sida.mxd, 'TITEL', 'Område %s – kartblad %d av %d' % (namn, nr, len(blad)))
                    ke._satt_text(sida.mxd, 'UNDERTITEL', ('Angränsande blad: %s' % ', '.join(str(g) for g in gr[nr]))
                                  if gr[nr] else '')
                    ke._satt_text(sida.mxd, 'SKALA', 'Skala 1:%d' % kartblad_skala)
                    fil = os.path.join(ut_mapp, 'kartblad_%s_%02d.pdf' % (_ren(namn), nr))
                    _exportera_pdf(sida, fil, dpi, oppen)
                    filer.append(fil)
                    ut['kartblad'].setdefault(namn, []).append(fil)
                    rp = _polygon(rekt, sr)
                    for k in omr_obj[namn]:
                        if k in geom and _skar(rekt, _utb(geom[k])) and not rp.disjoint(geom[k]):
                            kartblad_per_obj.setdefault(k, []).append((namn, nr, fil))
                logg('  %s: %d blad' % (namn, len(blad)))
                if samlad and filer:
                    _samla(filer, os.path.join(ut_mapp, 'kartblad_%s.pdf' % _ren(namn)))

        # --- schaktgrupper
        if grupper:
            logg('Schaktgrupper')
            omr_for = {}
            for n, nycklar in omr_obj.items():
                for k in nycklar:
                    omr_for.setdefault(k, set()).add(n)
            for g in sorted(grupper):
                nycklar = grupper[g]
                utb = None
                for k in nycklar:
                    utb = ke._sla_ihop(utb, _utb(geom[k]))
                sida, skala, ryms = valj_sida(sidor, utb, SKALOR_GRUPP, MARGINAL_GRUPP_M)
                stall_in(sida, grupp=g, etiketter=True)
                ke._centrera(sida.df, utb, skala, sida.ram)
                leds = [l for k in nycklar for l in per[k] if l.get('schaktgrupp') is not None and int(l['schaktgrupp']) == g]
                system = '+'.join(sorted(set(txt(l.get('system') or '?') for l in leds),
                                         key=lambda s: SYSTEM_ORDNING.find(s) % 9))
                omr = sorted(set(n for k in nycklar for n in omr_for.get(k, ())))
                ke._satt_text(sida.mxd, 'TITEL', 'Schaktgrupp %d' % g)
                ke._satt_text(sida.mxd, 'UNDERTITEL', '  ·  '.join(
                    x for x in (('Område ' + ', '.join(omr)) if omr else '', system,
                                '%d ledningar' % len(leds)) if x))
                ke._satt_text(sida.mxd, 'SKALA', 'Skala 1:%d' % skala)
                fil = os.path.join(ut_mapp, 'schaktgrupp_%03d.pdf' % g)
                _exportera_pdf(sida, fil, dpi, oppen)
                ut['schaktgrupper'].append(fil)
                gruppkarta[g] = fil
                logg('  %s  1:%d  %s%s' % (os.path.basename(fil), skala, system, '' if ryms else '  (ryms inte)'))
            if samlad and len(ut['schaktgrupper']) > 1:
                _samla(ut['schaktgrupper'], os.path.join(ut_mapp, 'kartor_schaktgrupper.pdf'))
    finally:
        for s in sidor:
            s.aterstall()
        for s in sidor:
            if s.mxd is not oppen:
                s.stang()
        sidor = None
        for namn in (TMP_GRUPPER, TMP_KARTBLAD):
            try:
                fc = os.path.join(arcpy.env.scratchGDB, namn)
                if arcpy.Exists(fc):
                    arcpy.Delete_management(fc)
            except Exception:
                pass

    # --- kartblad och sokvagar till ledningarna och till kartblad.csv
    def sokvag(fil):
        return os.path.join(txt(kartmapp), os.path.basename(fil)) if kartmapp else fil

    def bladtext(k):
        b = kartblad_per_obj.get(k) or []
        omr = []
        for n, nr, fil in b:
            if n not in omr:
                omr.append(n)
        if len(omr) <= 1:
            return ', '.join(str(nr) for n, nr, fil in b) or None
        return '; '.join('%s: %s' % (o, ', '.join(str(nr) for n, nr, fil in b if n == o)) for o in omr)

    def gruppfil(k):
        for l in per[k]:
            if l.get('schaktgrupp') is not None and int(l['schaktgrupp']) in gruppkarta:
                return sokvag(gruppkarta[int(l['schaktgrupp'])])
        return None

    if kartblad_per_obj or gruppkarta:
        rader = []
        for k in sorted(per, key=lambda k: (k[0], k[1])):
            b = kartblad_per_obj.get(k) or []
            for l in per[k]:
                rader.append([l.get('id'), l.get('omrade') or (b[0][0] if b else ''), bladtext(k),
                              sokvag(b[0][2]) if b else '', l.get('schaktgrupp'), gruppfil(k)])
        csv = os.path.join(os.path.dirname(os.path.abspath(txt(resultat_json))), KARTBLAD_CSV)
        try:
            _skriv_csv(csv, rader)
            logg('Skrev %s (ledning_analys.py lagger kartbladen i Excel nar den kors igen)' % csv)
        except Exception as e:
            logg('  kunde inte skriva %s: %s' % (csv, txt(e)))

    if skriv_falt and (kartblad_per_obj or gruppkarta):
        for namn in lagernamn:
            mina = [k for k in per if k[0] == namn]
            try:
                src = kalla(hitta_lager(namn))[0]
                namnmap = il.lagg_till_falt(src, FALT_KARTOR)
            except Exception as e:
                logg('  OBS: kunde inte lagga till falten %s i %s: %s' % (', '.join(f[0] for f in FALT_KARTOR),
                                                                          namn, txt(e)))
                continue
            kol = ['OID@'] + [namnmap[f[0]] for f in FALT_KARTOR]
            varden = {}
            for k in mina:
                b = kartblad_per_obj.get(k) or []
                varden[k[1]] = [(bladtext(k) or '')[:100] or None, (sokvag(b[0][2]) if b else '')[:254] or None,
                                (gruppfil(k) or '')[:254] or None]

            def skriv(src=src, kol=kol, varden=varden):
                n = 0
                with arcpy.da.UpdateCursor(src, kol) as cur:
                    for rad in cur:
                        v = varden.get(rad[0])
                        if v is None:
                            v = [None, None, None]
                        else:
                            n += 1
                        if kartblad:
                            rad[1], rad[2] = v[0], v[1]
                        if schaktgrupper:
                            rad[3] = v[2]
                        cur.updateRow(rad)
                return n
            try:
                n = il.med_redigering(src, skriv, namn)
                logg('  %s: KARTBLAD/KARTA/GRUPPKARTA skrivna for %d objekt (hyperlank: Display > Support '
                     'Hyperlinks using field)' % (namn, n))
            except Exception as e:
                logg('  OBS: kunde inte skriva kartfalten i %s: %s' % (namn, txt(e)))
    logg('KLART')
    return ut
