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

Tva mallar (.mxd) med liggande respektive staende layout kan anges; da valjs per stracka den
mall som ger minsta skala (vid lika skala den dar strackan fyller sidan bast). Mallarna maste ha
samma lager (namn) som den oppna kartan. Utan mallar anvands den oppna kartans layout.

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
    try:
        df.extent = arcpy.Extent(cx - b / 2.0, cy - h / 2.0, cx + b / 2.0, cy + h / 2.0)
        df.scale = skala
    except Exception as e:
        raise RuntimeError('Kan inte flytta dataramen (%s) - dataramens Extent ska vara "Automatic", '
                           'inte Fixed Scale/Fixed Extent (Data Frame Properties > Data Frame)' % txt(e))


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
        try:
            return 'etapp_%02d.pdf' % int(post.get('ETAPP') or 0)
        except (TypeError, ValueError):
            return 'etapp_%s.pdf' % ren(post.get('ETAPP'))
    fil = os.path.splitext(os.path.basename(txt(post.get('TV3_FIL') or '')))[0]
    return '%s_%s_%s_%s-%s.pdf' % (ren(post.get('BEDOMNING') or 'X'), ren(fil), txt(post.get('NR') or ''),
                                   ren(post.get('FRAN_BRUNN')), ren(post.get('TILL_BRUNN')))


def _titel(post, per_etapp, antal=1):
    if per_etapp:
        return 'Etapp %s (%d sträckor)' % (txt(post.get('ETAPP') or '?'), antal)
    nr = post.get('NR')
    return 'Sträcka %s%s → %s' % (('%s: ' % txt(nr)) if nr is not None else '',
                                            txt(post.get('FRAN_BRUNN') or '?'), txt(post.get('TILL_BRUNN') or '?'))


def _undertitel(post, skala, per_etapp, langd=None, metoder=None):
    delar = []
    if not per_etapp and post.get('BEDOMNING'):
        delar.append('Klass %s' % txt(post['BEDOMNING']))
    if post.get('ETAPP') is not None and not per_etapp:
        delar.append('Etapp %s' % txt(post['ETAPP']))
    if metoder:
        delar.append(' / '.join(sorted(set(txt(m) for m in metoder if m))))
    elif post.get('METOD'):
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


def _lager_i_doc(mxd, namn):
    """Lagret i kartdokumentet med namnet namn: forst pa langt namn (Grupp\\Lager), sedan pa
    kortnamnet (lagret kan ligga i en annan grupp eller lost i mallen). Skiftlage spelar ingen roll."""
    sokt = txt(namn).strip().strip("'").lower()
    kort = sokt.split('\\')[-1]
    lager = [l for l in arcpy.mapping.ListLayers(mxd) if not getattr(l, 'isGroupLayer', False)]
    for l in lager:
        try:
            if txt(l.longName).strip().lower() == sokt or txt(l.name).strip().lower() == sokt:
                return l
        except Exception:
            continue
    for l in lager:
        try:
            if txt(l.name).strip().lower() == kort:
                return l
        except Exception:
            continue
    return None


def _huvudram(mxd):
    """Dataramen att exportera ur: den storsta pa sidan (mallar med oversiktskarta har flera),
    annars den forsta."""
    ramar = arcpy.mapping.ListDataFrames(mxd)
    if not ramar:
        raise RuntimeError('Kartdokumentet har ingen dataram')
    try:
        return max(ramar, key=lambda d: float(d.elementWidth) * float(d.elementHeight))
    except Exception:
        return ramar[0]


KOPIANAMN = 'Aktuell stracka (export)'   # tillfalligt lager som visar den aktuella strackan
MARKERING_LYR = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'markering.lyr')
MARKERING_TRANSPARENS = 50               # procent genomskinlighet pa det tillfalliga lagret
MARKERING_TIPS = (
    'Ingen markeringssymbol (.lyr) - den aktuella strackan visas med ArcMaps urvalsfarg (och med '
    'strackslagrets egen symbologi, klassfargerna, nar lagret slacks). Skapa symbolen en gang i '
    'ArcMap: hogerklicka strackslagret > Properties > '
    'Symbology > Features > Single symbol, valj en gul linje med bredd ca 8 pt > OK; hogerklicka '
    'lagret igen > Save As Layer File... och spara som %s. Filen ar forvald i dialogen nar den finns; '
    'genomskinligheten satts av verktyget.' % txt(MARKERING_LYR))


class Layout(object):
    """Ett kartdokument att exportera ur: det oppna (CURRENT) eller en mall (.mxd). Haller
    dataramen, dess matt per skalenhet, lagret med strackorna och ev. markeringslagret."""

    def __init__(self, mxd, namn, lagernamn, markeringsnamn=None, lyr=None, kopiera_fran=None,
                 kopiera_synliga=None, dolj_strackor=False, markering_lyr=None, transparens=None):
        """kopiera_fran: (lager, markeringslager) ur den oppna kartan - saknas de i mallen laggs de
        in i mallens dataram for exporten (mallen sparas inte). kopiera_synliga: den oppna kartans
        kartdokument - alla dess synliga toppnivalager (inkl. grupplager) laggs in i mallen i samma
        ordning, sa att kartan ser ut som pa skarmen. dolj_strackor: slack lagret med alla strackor
        under exporten och visa bara den aktuella (markeringslagret, eller en tillfallig kopia av
        lagret med definitionsfraga nar inget markeringslager finns). Ar lagret redan slackt i
        kartan gors detsamma automatiskt. markering_lyr: .lyr-fil med symbolen for den aktuella
        strackan (t.ex. bred gul linje) - utan markeringslager laggs da ett tillfalligt lager med den
        symbologin overst, aven nar strackslagret ar tant; transparens: procent genomskinlighet pa
        det tillfalliga lagret."""
        self.mxd = mxd
        self.namn = namn
        self._kopia = None
        self._gammal_synlig = None
        self._transparens = transparens
        try:
            if txt(mxd.activeView).upper() != 'PAGE_LAYOUT':
                mxd.activeView = 'PAGE_LAYOUT'      # dataramens utbredning galler layouten, inte datavyn
                logg('  %s: vaxlade till layoutvyn' % namn)
        except Exception:
            pass
        self.df = _huvudram(mxd)
        self.ram = dataram_matt(self.df)
        self.liggande = self.ram[0] >= self.ram[1]
        try:
            if abs(float(self.df.rotation or 0)) > 0.01:
                logg('  OBS: dataramen i %s ar roterad %.0f grader - skalvalet raknar pa oroterad ram, '
                     'satt rotationen till 0' % (namn, float(self.df.rotation)))
        except Exception:
            pass
        kop_lyr, kop_mark = (kopiera_fran or (None, None))
        if kopiera_synliga is not None:
            self._kopiera_synliga(kopiera_synliga)
        self.lyr = lyr if lyr is not None else _lager_i_doc(mxd, lagernamn)
        if self.lyr is None and kop_lyr is not None:
            self.lyr = self._lagg_in(kop_lyr, lagernamn)
        if self.lyr is None:
            raise RuntimeError('Mallen %s saknar lagret "%s" - lagg in det (samma namn som i kartan)'
                               % (namn, txt(lagernamn)))
        self.mark_lyr = _lager_i_doc(mxd, markeringsnamn) if markeringsnamn else None
        if markeringsnamn and self.mark_lyr is None and kop_mark is not None:
            self.mark_lyr = self._lagg_in(kop_mark, markeringsnamn)
        if markeringsnamn and self.mark_lyr is None:
            logg('  mallen %s saknar markeringslagret "%s" - %s anvands i stallet'
                 % (namn, txt(markeringsnamn), 'markeringssymbolen' if markering_lyr else 'urval'))
        if self.mark_lyr is not None and self.mark_lyr is self.lyr:
            logg('  markeringslagret ar samma lager som strackorna i %s - en definitionsfraga skulle dolja '
                 'alla andra strackor; urval anvands i stallet' % namn)
            self.mark_lyr = None
        self._gammal_dq = getattr(self.mark_lyr, 'definitionQuery', None) if self.mark_lyr is not None else None
        # Anvandarens eget urval i lagret aterstalls efterat (markeringen gors med urval)
        self._gammalt_urval = None
        try:
            self._gammalt_urval = list(self.lyr.getSelectionSet() or [])
        except Exception:
            pass
        if self.mark_lyr is None and markering_lyr:
            self.mark_lyr = self._skapa_kopia(lagernamn, txt(markering_lyr))
            if self.mark_lyr is not None:
                logg('  %s: den aktuella strackan visas med symbolen i %s (tillfalligt lager "%s")'
                     % (self.namn, os.path.basename(txt(markering_lyr)), KOPIANAMN))
        try:
            synligt = bool(self.lyr.visible)
        except Exception:
            synligt = True
        if dolj_strackor or not synligt:
            self._dolj(lagernamn, dolj_strackor)

    def _skapa_kopia(self, lagernamn, lyr_fil=None):
        """Lagger ett tillfalligt lager overst i dataramen som visar den aktuella strackan
        (definitionsfraga satts i markera) och returnerar det, eller None. Med lyr_fil skapas
        lagret fran strackslagrets datakalla och far symbologin ur filen; annars en kopia av
        strackslagret (med dess symbologi). Genomskinligheten (self._transparens) satts bara
        pa symbolen ur filen - en genomskinlig klassfargad linje syns daligt."""
        try:
            if lyr_fil:
                src = kalla(self.lyr)[0]
                ny = arcpy.mapping.Layer(src)
                ny.name = KOPIANAMN
                # Symbologin satts innan lagret laggs i kartan (AddLayer lagger in en kopia) ...
                arcpy.ApplySymbologyFromLayer_management(ny, lyr_fil)
                arcpy.mapping.AddLayer(self.df, ny, 'TOP')
                vantat = KOPIANAMN
            else:
                arcpy.mapping.AddLayer(self.df, self.lyr, 'TOP')
                vantat = txt(getattr(self.lyr, 'name', lagernamn))
        except Exception as e:
            logg('  kunde inte lagga in ett tillfalligt lager for den aktuella strackan i %s: %s'
                 % (self.namn, txt(e)))
            return None
        # Det nya lagret ligger overst i dataramen, dvs. forst i listan
        kopia = None
        try:
            lager = arcpy.mapping.ListLayers(self.mxd, '', self.df)
            if lager and txt(lager[0].name).strip().lower() == vantat.strip().lower():
                kopia = lager[0]
            else:
                for l in lager:
                    if txt(l.name).strip().lower() == vantat.strip().lower() \
                            and not getattr(l, 'isGroupLayer', False):
                        kopia = l
                        break
        except Exception as e:
            logg('  kunde inte lasa lagren i %s: %s' % (self.namn, txt(e)))
        if kopia is None:
            logg('  OBS: det tillfalliga lagret "%s" hittades inte i %s efter att det lagts in - '
                 'tas det inte bort automatiskt, ta bort det for hand' % (vantat, self.namn))
            return None
        try:
            kopia.name = KOPIANAMN
            kopia.definitionQuery = ''
            kopia.visible = True
        except Exception:
            pass
        if lyr_fil:
            # ... och for sakerhets skull aven pa lagret i kartan (symbology_only)
            try:
                arcpy.mapping.UpdateLayer(self.df, kopia, arcpy.mapping.Layer(lyr_fil), True)
            except Exception:
                pass
            if self._transparens is not None:
                try:
                    kopia.transparency = int(self._transparens)
                except Exception as e:
                    logg('  kunde inte satta genomskinlighet pa "%s": %s' % (KOPIANAMN, txt(e)))
        self._kopia = kopia
        self._gammal_dq = ''
        return kopia

    def _dolj(self, lagernamn, begart):
        """Slacker lagret med alla strackor under exporten. Finns inget markeringslager laggs ett
        tillfalligt lager (KOPIANAMN) overst och anvands som markeringslager - dess
        definitionsfraga satts till den aktuella strackan. Lagret tas bort i aterstall()."""
        try:
            self._gammal_synlig = bool(self.lyr.visible)
        except Exception:
            self._gammal_synlig = None
        if self.mark_lyr is None:
            self.mark_lyr = self._skapa_kopia(lagernamn)
            if self.mark_lyr is None:
                logg('  OBS: %s - strackslagret kan inte slackas utan markeringslager (ingen kopia kunde '
                     'laggas in); lagret visas tant' % self.namn)
                try:
                    self.lyr.visible = True
                except Exception:
                    pass
                return
        try:
            self.lyr.visible = False
        except Exception as e:
            logg('  kunde inte slacka strackslagret i %s: %s' % (self.namn, txt(e)))
            return
        logg('  %s: strackslagret ar slackt%s - bara den aktuella strackan visas%s'
             % (self.namn, '' if begart else ' i kartan',
                '' if self._kopia is None else ' (tillfalligt lager "%s")' % KOPIANAMN))

    def _kopiera_synliga(self, oppen):
        """Lagger in den oppna kartans synliga toppnivalager i mallens dataram, i samma ordning
        (nederst forst med 'TOP' sa att det oversta hamnar overst). Lager som redan finns i mallen
        (samma namn) hoppas over."""
        try:
            odf = _huvudram(oppen)
            alla = arcpy.mapping.ListLayers(oppen, '', odf)
        except Exception as e:
            logg('  kunde inte lasa den oppna kartans lager: %s' % txt(e))
            return
        finns = set()
        for l in arcpy.mapping.ListLayers(self.mxd):
            try:
                finns.add(txt(l.name).strip().lower())
            except Exception:
                pass
        topp = []
        for l in alla:
            try:
                if txt(l.longName) != txt(l.name):
                    continue                  # ligger i ett grupplager - foljer med gruppen
                if not l.visible:
                    continue
                if txt(l.name).strip().lower() in finns:
                    continue
                topp.append(l)
            except Exception:
                continue
        n = 0
        for l in reversed(topp):
            try:
                arcpy.mapping.AddLayer(self.df, l, 'TOP')
                n += 1
            except Exception as e:
                logg('  kunde inte lagga in lagret "%s" i %s: %s' % (txt(getattr(l, 'name', '?')), self.namn, txt(e)))
        if n:
            logg('  %d synliga lager ur den oppna kartan inlagda i %s (%s)'
                 % (n, self.namn, ', '.join(txt(l.name) for l in topp[:6]) + (' ...' if len(topp) > 6 else '')))

    def _lagg_in(self, lager, namn):
        """Lagger in en kopia av ett lager ur den oppna kartan overst i mallens dataram (med dess
        symbologi) och returnerar mallens lagerobjekt. Mallen sparas inte, sa den paverkas inte."""
        try:
            arcpy.mapping.AddLayer(self.df, lager, 'TOP')
            ny = _lager_i_doc(self.mxd, txt(getattr(lager, 'name', namn))) or _lager_i_doc(self.mxd, namn)
            if ny is not None:
                logg('  lagret "%s" saknades i %s - kopierades in fran den oppna kartan for exporten'
                     % (txt(getattr(lager, 'name', namn)), self.namn))
            return ny
        except Exception as e:
            logg('  kunde inte lagga in lagret "%s" i %s: %s' % (txt(namn), self.namn, txt(e)))
            return None

    def kontrollera_kalla(self, src):
        """Varnar om mallens lager pekar pa en annan featureklass an den oppna kartans."""
        for l, vad in ((self.lyr, 'lagret'), (self.mark_lyr, 'markeringslagret')):
            if l is None or l is self._kopia:
                continue
            try:
                k = kalla(l)[0]
            except Exception:
                continue
            if os.path.normcase(txt(k)) != os.path.normcase(txt(src)):
                logg('  OBS: %s i %s pekar pa %s, inte pa %s - markeringen kan hamna pa fel objekt'
                     % (vad, self.namn, txt(k), txt(src)))

    def stang(self):
        """Slapper referenserna till kartdokumentet (mallar laser annars datakallan)."""
        self.df = self.lyr = self.mark_lyr = self.mxd = None

    def passning(self, utb, skalor, marginal):
        """(skala, ryms, fyllnad) - fyllnad = hur stor del av ramen strackan tar i den skalan."""
        skala, ryms = valj_skala(utb, self.ram, skalor, marginal)
        b = (utb[2] - utb[0]) + 2 * marginal
        h = (utb[3] - utb[1]) + 2 * marginal
        fyll = max(b / (self.ram[0] * skala), h / (self.ram[1] * skala))
        return skala, ryms, fyll

    def markera(self, oids, src, oidfalt):
        where = '%s IN (%s)' % (arcpy.AddFieldDelimiters(src, oidfalt), ','.join(str(o) for o in oids))
        if self.mark_lyr is not None:
            self.mark_lyr.definitionQuery = where
            try:
                self.mark_lyr.visible = True
            except Exception:
                pass
            return
        try:
            self.lyr.setSelectionSet('NEW', list(oids))
        except Exception:
            arcpy.SelectLayerByAttribute_management(self.lyr, 'NEW_SELECTION', where)

    def avmarkera(self):
        """Tar bort markeringen av den aktuella strackan (definitionsfragan respektive urvalet)
        men behaller det tillfalliga lagret och slackningen - anropas mellan strackorna."""
        if self.lyr is None:
            return
        try:
            if self.mark_lyr is not None:
                self.mark_lyr.definitionQuery = self._gammal_dq or ''
            elif self._gammalt_urval:
                try:
                    self.lyr.setSelectionSet('NEW', list(self._gammalt_urval))
                except Exception:
                    arcpy.SelectLayerByAttribute_management(self.lyr, 'CLEAR_SELECTION')
            else:
                try:
                    self.lyr.setSelectionSet('NEW', [])
                except Exception:
                    arcpy.SelectLayerByAttribute_management(self.lyr, 'CLEAR_SELECTION')
        except Exception:
            pass

    def aterstall(self):
        """Slutstadning: avmarkerar, tander strackslagret igen och tar bort det tillfalliga
        lagret."""
        if self.lyr is None:
            return
        self.avmarkera()
        if self._gammal_synlig is not None:
            try:
                self.lyr.visible = self._gammal_synlig
            except Exception:
                pass
            self._gammal_synlig = None
        if self._kopia is not None:
            try:
                arcpy.mapping.RemoveLayer(self.df, self._kopia)
            except Exception as e:
                logg('  kunde inte ta bort det tillfalliga lagret "%s" ur %s: %s - ta bort det for hand'
                     % (KOPIANAMN, self.namn, txt(e)))
            self._kopia = None
            self.mark_lyr = None


def valj_layout(layouter, utb, skalor, marginal):
    """Layouten som ger minsta skala; vid lika skala den dar strackan fyller ramen bast.
    Returnerar (layout, skala, ryms)."""
    bast = None
    for lo in layouter:
        skala, ryms, fyll = lo.passning(utb, skalor, marginal)
        # Ryms strackan: storst fyllnad vinner vid lika skala. Ryms den inte i nagon layout:
        # minst overskjutning vinner (annars valdes den layout dar strackan stack ut mest).
        nyckel = (not ryms, skala, fyll if not ryms else -fyll)
        if bast is None or nyckel < bast[0]:
            bast = (nyckel, lo, skala, ryms)
    return bast[1], bast[2], bast[3]


def exportera(bedomda, ut_mapp, urval='atgard', skalor=SKALOR, marginal=MARGINAL_M, dpi=200,
              markeringslager=None, samlad=True, per_etapp=False, skriv_falt=True,
              kartmapp=None, bara_valda_klasser=KLASSER, mall_liggande=None, mall_staende=None,
              kopiera_synliga=True, dolj_strackor=False, markering_lyr=None,
              transparens=MARKERING_TRANSPARENS):
    """Exporterar en PDF per stracka (eller per etapp) till ut_mapp. Returnerar lista med
    (filnamn, skala, ryms, layoutnamn). urval: 'atgard' (METOD ifyllt), 'AB' (BEDOMNING i A/B),
    'valda' (markerade i kartan), 'alla'. markeringslager: lager som pekar pa samma featureklass
    och vars definitionsfraga satts till den aktuella strackan (tydligare an urvalsfargen).
    mall_liggande/mall_staende: .mxd-filer med liggande resp. staende layout; anges bada valjs
    per stracka den som ger minsta skala (vid lika skala den dar strackan fyller sidan bast).
    Saknar mallarna lagret med strackorna (eller markeringslagret) kopieras det in fran den oppna
    kartan under korningen (mallen sparas inte). kopiera_synliga: alla synliga lager i den oppna
    kartan laggs in i mallen i samma ordning, sa att kartan ser ut som pa skarmen (annars bara
    strack- och markeringslagret). Utan mallar anvands den oppna kartans layout.
    dolj_strackor: slack lagret med alla strackor under exporten sa att bara den aktuella strackan
    syns (markeringslagret, eller en tillfallig kopia av lagret med definitionsfraga); gors ocksa
    automatiskt nar lagret ar slackt i kartan. markering_lyr: .lyr-fil med symbolen for den
    aktuella strackan (bred gul linje); utan markeringslager laggs ett tillfalligt lager med den
    symbologin overst (aven med strackslagret tant); transparens: procent genomskinlighet pa det.
    kartmapp: mapp som skrivs i faltet KARTA i stallet for ut_mapp (Citrix)."""
    mxd = _mxd()
    if mxd is None:
        raise RuntimeError('Verktyget kors i ArcMap med ett oppet kartdokument')
    if not os.path.isdir(ut_mapp):
        os.makedirs(ut_mapp)
    skalor = sorted(int(s) for s in skalor if _tal(s) and float(s) > 0) or list(SKALOR)

    lyr = hitta_lager(bedomda) if isinstance(bedomda, (TEXTTYP, bytes)) else bedomda
    if isinstance(lyr, (TEXTTYP, bytes)):
        raise RuntimeError('Ange lagret som det heter i kartan (inte en sokvag) - markering och urval kraver ett lager')
    lagernamn = txt(getattr(lyr, 'longName', None) or getattr(lyr, 'name', bedomda))
    src, dq = kalla(lyr)
    falt = _falt(src)
    for f in ('FRAN_BRUNN', 'TILL_BRUNN'):
        if f not in falt:
            raise RuntimeError('Lagret saknar faltet %s - valj lagret fran "Skapa ledningslager"' % f)
    oidfalt = arcpy.Describe(src).OIDFieldName

    # Layouter: tva mallar (liggande/staende) eller den oppna kartan
    mark_namn = None
    if markeringslager:
        try:
            mark_namn = txt(getattr(hitta_lager(markeringslager), 'longName', markeringslager))
        except Exception as e:
            logg('  markeringslagret hittades inte (%s) - urval anvands i stallet' % txt(e))
    if markering_lyr and not os.path.isfile(txt(markering_lyr)):
        logg('Markeringssymbolen finns inte: %s' % txt(markering_lyr))
        markering_lyr = None
    if not mark_namn and not markering_lyr:
        logg(MARKERING_TIPS)
    layouter = []
    mallar = [(m, n) for m, n in ((mall_liggande, 'liggande'), (mall_staende, 'staende')) if m]
    if mallar:
        try:
            for mall, namn in mallar:
                if not os.path.isfile(txt(mall)):
                    raise RuntimeError('Mallen finns inte: %s' % txt(mall))
                doc = arcpy.mapping.MapDocument(txt(mall))
                mark_lyr_oppen = None
                if mark_namn:
                    try:
                        mark_lyr_oppen = hitta_lager(markeringslager)
                    except Exception:
                        mark_lyr_oppen = None
                lo = Layout(doc, '%s (%s)' % (namn, os.path.basename(txt(mall))), lagernamn, mark_namn,
                            kopiera_fran=(lyr, mark_lyr_oppen), kopiera_synliga=mxd if kopiera_synliga else None,
                            dolj_strackor=dolj_strackor, markering_lyr=markering_lyr, transparens=transparens)
                lo.kontrollera_kalla(src)
                layouter.append(lo)
                logg('Mall %s: dataramen ar %.0f x %.0f m i skala 1:%d%s'
                     % (lo.namn, lo.ram[0] * skalor[0], lo.ram[1] * skalor[0], skalor[0],
                        '' if lo.liggande == (namn == 'liggande') else '  (OBS: ramen ser %s ut)'
                        % ('liggande' if lo.liggande else 'staende')))
        except Exception:
            for lo in layouter:
                lo.stang()
            raise
    else:
        lo = Layout(mxd, 'oppna kartan', lagernamn, mark_namn, lyr=lyr, dolj_strackor=dolj_strackor,
                    markering_lyr=markering_lyr, transparens=transparens)
        layouter.append(lo)
        logg('Dataramen ar %.0f x %.0f m i skala 1:%d' % (lo.ram[0] * skalor[0], lo.ram[1] * skalor[0], skalor[0]))

    las = ['OID@', 'SHAPE@'] + [falt[f] for f in ('FRAN_BRUNN', 'TILL_BRUNN', 'BEDOMNING', 'ETAPP', 'METOD',
                                                  'LANGD_M', 'MATERIAL', 'DIMENSION', 'NR', 'TV3_FIL') if f in falt]
    namn = ['OID@', 'SHAPE@'] + [f for f in ('FRAN_BRUNN', 'TILL_BRUNN', 'BEDOMNING', 'ETAPP', 'METOD',
                                             'LANGD_M', 'MATERIAL', 'DIMENSION', 'NR', 'TV3_FIL') if f in falt]
    valda_oid = None
    if urval == 'valda':
        try:
            valda_oid = set(lyr.getSelectionSet() or [])
        except Exception:
            raise RuntimeError('Urvalet "markerade i kartan" kraver att lagret valjs i kartan (inte en sokvag)')
        if not valda_oid:
            raise RuntimeError('Inga strackor ar markerade i kartan')

    # Geometrin lases i dataramens koordinatsystem (dataramen kan ha ett annat an lagret)
    sr_ram = None
    try:
        sr_ram = layouter[0].df.spatialReference
        for lo in layouter[1:]:
            if txt(lo.df.spatialReference.name) != txt(sr_ram.name):
                logg('  OBS: mallarna har olika koordinatsystem (%s / %s) - utbredningen raknas i %s'
                     % (txt(sr_ram.name), txt(lo.df.spatialReference.name), txt(sr_ram.name)))
    except Exception:
        sr_ram = None
    poster = []
    with arcpy.da.SearchCursor(src, las, dq, spatial_reference=sr_ram) as mark:
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

    ut = []
    filer = []
    karta_per_oid = {}
    antal_per_layout = {}
    try:
        for grupp in grupper:
            utb = None
            for p in grupp:
                utb = _sla_ihop(utb, _utbredning(p['SHAPE@']))
            lo, skala, ryms = valj_layout(layouter, utb, skalor, marginal)
            antal_per_layout[lo.namn] = antal_per_layout.get(lo.namn, 0) + 1
            _centrera(lo.df, utb, skala, lo.ram)
            for annan in layouter:
                if annan is not lo:
                    annan.avmarkera()
            lo.markera([p['OID@'] for p in grupp], src, oidfalt)
            p0 = grupp[0]
            langd = sum(_tal(p.get('LANGD_M')) or 0 for p in grupp) if per_etapp else None
            _satt_text(lo.mxd, 'TITEL', _titel(p0, per_etapp, len(grupp)))
            _satt_text(lo.mxd, 'UNDERTITEL', _undertitel(p0, skala, per_etapp, langd,
                                                         [p.get('METOD') for p in grupp] if per_etapp else None))
            _satt_text(lo.mxd, 'SKALA', 'Skala 1:%d' % skala)
            if lo.mxd is mxd:
                try:
                    arcpy.RefreshActiveView()
                except Exception:
                    pass
            fil = os.path.join(ut_mapp, _filnamn(p0, per_etapp))
            if fil in filer:                      # samma klass/fil/nr/brunnar tva ganger (t.ex. filer med samma namn)
                stam, andelse = os.path.splitext(fil)
                k = 2
                while '%s_%d%s' % (stam, k, andelse) in filer:
                    k += 1
                fil = '%s_%d%s' % (stam, k, andelse)
                logg('  filnamnet fanns redan - skriver %s' % os.path.basename(fil))
            arcpy.mapping.ExportToPDF(lo.mxd, fil, 'PAGE_LAYOUT', resolution=int(dpi),
                                      image_quality='BEST', georef_info=True)
            filer.append(fil)
            ut.append((fil, skala, ryms, lo.namn))
            for p in grupp:
                karta_per_oid[p['OID@']] = fil if not kartmapp else os.path.join(txt(kartmapp), os.path.basename(fil))
            logg('  %s  1:%d  %s%s' % (os.path.basename(fil), skala, lo.namn if len(layouter) > 1 else '',
                                        '' if ryms else '  (ryms inte - storsta skalan)'))
    finally:
        for lo in layouter:
            lo.aterstall()
        for lo in layouter:
            if lo.mxd is not mxd:
                lo.stang()              # mallen sparas inte; referenserna slapps innan faltet skrivs
        layouter = [lo for lo in layouter if lo.mxd is not None]
    if len(layouter) > 1:
        logg('  layout: ' + ', '.join('%s %d' % (n, a) for n, a in sorted(antal_per_layout.items())))

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

    ej = [f for f, s, r, l in ut if not r]
    if ej:
        logg('  %d kartor rymdes inte i storsta skalan 1:%d - korta ned sidan eller lagg till storre skalor'
             % (len(ej), max(skalor)))
    logg('KLART')
    return ut
