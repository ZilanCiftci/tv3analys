#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ledning_analys.py – inventeringsbok och kostnad lägga om mot strumpa (PLAN_ledningsinventering.md steg 3)

Läser ledningsdata.json från ArcMap-verktyget Ledningsinventering och räknar två alternativ för hela nätet:

  Alt. lägg om  – alla ledningar läggs om i helschakt; ledningar i samma grav delar graven (pris per meter
                  grav efter antal ledningar, antal större än 400 mm och schakttyp). Standard är schakttypen
                  djup: gravens medeldjup under 2 m ledningschakt, 2–3 m spontlåda (spontsläde), från 3 m
                  ledningschakt + slagen spont på båda sidor (medeldjup + 3 m, kr/m²). Räknas även för de fyra
                  fasta schakttyperna (ledningschakt, med spontsläde, bergschakt, KC-förstärkning).
  Alt. schaktfritt  – spill-, dag- och kombinerade ledningar infodras (rörspräcks om infodring inte går och
                  dimensionen tillåter, annars läggs de om). Vattenledningar rörspräcks – men bara när
                  avloppsledningarna i samma grav också klaras utan schakt; läggs någon av dem om, läggs
                  vattenledningen om i samma grav. Ledningar som läggs om i alternativet delar grav.

Serviserna förnyas i båda alternativen: i helschakt per meter servisledning (serviser till samma fastighet
i samma servisgrav), schaktfritt infodras de per meter; vid rörspräckning tillkommer en servisgrop per servis.

Skriver
  * <utdata>/ledningsinventering.xlsx – Sammanfattning, Ledningar, Kostnad per schaktgrupp, Serviser,
                                         Parallella, Korsningar, Brunnar, Datakvalitet, Kostnadsposter
  * <utdata>/ledningsresultat.json    – djup, serviser, schaktgrupp och kostnader per ledning,
                                         för att fylla fält i kartan

Användning:
    python ledning_analys.py ledningsdata.json
    python ledning_analys.py ledningsdata.json -o utdata --schakttyp bergschakt --kostnader kostnader_inventering.csv

Alla belopp står i kostnadsfilen (kostnader_inventering.csv), inga i koden. Inga rekommendationer – bara
beräkningar (användarens beslut okt 2026).
"""
from __future__ import annotations

import argparse
import json
import math
import os
import re
import sys
from collections import defaultdict
from datetime import datetime

HAR = os.path.dirname(os.path.abspath(__file__))
if HAR not in sys.path:
    sys.path.insert(0, HAR)
from tv3_analys import las_text  # noqa: E402

# ============================================================================
# KONFIG – ändra här, inte i koden
# ============================================================================
KOSTNADSFIL = os.path.join(HAR, "kostnader_inventering.csv")   # alla priser (SVOA:s kalkylverktyg okt 2026)
UTDATA_XLSX = "ledningsinventering.xlsx"
UTDATA_JSON = "ledningsresultat.json"

SCHAKTTYPER = {"ledningschakt": "Ledningschakt", "spontslade": "Med spontsläde",
               "bergschakt": "Bergschakt", "kc": "KC-förstärkning"}
SCHAKTTYP_DJUP = "djup"           # schakttyp efter ledningens djup (spontlåda, slagen spont) – se nedan
ALLA_SCHAKTTYPER = {SCHAKTTYP_DJUP: "Efter djup – spont", **SCHAKTTYPER}
SCHAKTTYP = SCHAKTTYP_DJUP        # schakttypen i Alt. lägg om och Alt. schaktfritt (övriga redovisas också)
SPONTLADA_FRAN_M = 2.0            # schakttyp djup: grav från detta medeldjup (m) -> spontsläde/spontlåda (prisraderna)
SLAGEN_SPONT_FRAN_M = 3.0         # ... från detta medeldjup -> ledningschakt + slagen spont (post slagen_spont, kr/m²)
SPONT_NEDSLAGNING_M = 3.0         # slagen spont: höjd = gravens medeldjup + detta (m)
SPONT_SIDOR = 2                   # slagen spont på båda sidor av graven
SCHAKTSATT_NAMN = {"ledningschakt": "Ledningschakt", "spontlada": "Spontlåda", "slagen_spont": "Slagen spont",
                   "djup_saknas": "Djup saknas (ledningschakt)"}
STORRE_DIM = 400                  # ledningar över detta (mm) räknas som "större" i grav med 2–3 ledningar
MAX_LEDNINGAR_I_GRAV = 3          # prislistan har 1–3 ledningar i samma grav
PARALLELL_MIN_M = 10.0            # äldre ledningsdata.json (utan grav_delar): partner som går parallellt minst så här
                                  # långt delar grav på sin andel av längden (nyare filer: Ledningsinventering delar upp)
VATTEN_OMLAGGNING_ANDEL = 0.5     # alt. schaktfritt: vattenledningen läggs om i samma grav när minst denna andel av dess
                                  # längd ligger intill en avloppsledning som läggs om (annars rörspräcks den)
STANDARDDJUP_M = 2.5              # schaktpriset gäller till detta djup (vattengång) ...
DJUPTILLAGG_STEG_M = 1.5          # ... därefter ett djuptillägg per påbörjade 1,5 m
STRUMPA_SYSTEM = ("S", "D", "K")  # system som kan infodras – vattenledningar (V) infodras inte
STRUMPA_DIM_MIN = 100             # infodring från denna dimension (mm); övre gräns = sista priset i filen
SPRACKNING_DIM_MAX = 400          # rörspräckning för dimensioner till och med detta (mm), alla system
GROP_PER_M = 50                   # rörspräckning: en införingsgrop och en maskingrop per påbörjade 50 m
SERVIS_GRAV_AVSTAND_M = 3.0       # serviser vars båda ändar ligger så här nära varandra delar servisgrav
SERVIS_LANGD_MAX_M = None         # meter servisledning som förnyas per servis; None = hela servisens längd
SERVIS_HATT_MAX_M = 6.0           # spill-, dag- och kombinerade serviser hit (m) infodras med hatt (kr/st) ...
SERVIS_HATT_SYSTEM = ("S", "D", "K")  # ... längre serviser och vattenserviser per meter servisledning
FRAMSCHAKTA_TB = True             # infodring: båda ändbrunnarna tillsyns-/rensbrunn -> framschaktning + ny brunn
BRUNNSTYP_TILLSYN = r"TB|RB"      # brunnstyp (t.ex. WellFunction STB/SRB) som räknas som tillsyns-/rensbrunn
VENTILGROP_KATEGORIER = ("ventil", "brandpost")   # rörspräckning: en ventilgrop per sådan anordning på ledningen
                                                  # (servisventiler ligger i servisgropen)
DIMENSION_UR_GRANNE = True        # saknad dimension antas lika med uppströms ledning av samma system
DIMENSION_NEDSTROMS_RESERV = True # ... och nedströms ledning när ingen uppströms har dimension
ANORDNING_NAMN = {"ventil": "Ventiler", "brandpost": "Brandposter", "servisventil": "Servisventiler",
                  "ledningsande": "Ledningsändar", "anordning": "Övriga anordningar"}

# Kolumner i fliken Ledningar som döljs (grupperade, fälls ut med plustecknet)
DOLDA_LEDNINGAR = [
    "Ledningstyp", "Renoveringsår", "VG från", "VG till", "Lutning GIS", "Djup min", "Djup max", "Djupkälla",
    "Täckning min", "Djup start", "Djup slut", "Samschakt med", "Parallella", "Korsningar",
    "Lägg om efter djup", "Lägg om ledningschakt", "Lägg om med spontsläde", "Lägg om bergschakt", "Lägg om KC-förstärkning",
    "Schakt", "Gravdjup", "Slagen spont", "Spont", "Djuptillägg", "Återställning", "Brunnar", "Serviser helschakt",
    "Infodring", "Infodring per meter", "Serviser infodring", "Hattar", "Framschaktning", "Infodring totalt",
    "Servisventiler", "Ledningsändar", "Övriga anordningar", "Schaktlängd", "Schaktsätt efter djup",
    "Slagen spont efter djup", "Spont efter djup",
    "Rörspräckning", "Rörspräckning per meter", "Gropar", "Servisgropar", "Ventilgropar", "Serviser rörspräckning",
    "Rörspräckning totalt", "Ledningar i graven alt. schaktfritt", "Område", "Lager", "OID", "Del",
    "Klippt vid områdesgränsen",
]

# Investeringskalkyl (fliken Investeringskalkyl) – SVOA:s mall "Kalkyl inriktningsbeslut LEDN", flikarna Alternativ,
# Entreprenadkalkyl och Investeringsärende BP2. Procentsatserna skrivs i gula celler och kan ändras i Excel.
BYGGHERRE_ANDEL = [                # moment: (planerafasen, genomförafasen) i andel av entreprenadkostnaden
    ("Projekt- och byggledning", 0.01, 0.05),
    ("Projektering", 0.05, 0.01),
    ("Geo, mark och miljö", 0.03, 0.02),
    ("Övriga byggherrekostnader", 0.01, 0.02),
]
KANDA_RISKER_ANDEL = 0.20          # kända risker i andel av byggherrekostnader (+ entreprenad i totalbudgeten)
OFORUTSETT_ANDEL = 0.20            # oförutsett i andel av summan ovanför (mallen: 20 %)
AFC_MATERIAL_ANDEL = 0.03          # Entreprenadkalkyl: material som SVOA tillhandahåller (AFC.1522) – rörspräckning, helschakt
AFC_SVOA_ARBETEN_ANDEL = 0.03      # ... arbeten som SVOA utför och som undantas från entreprenören (AFC.114) – samma tabeller
VINTERTILLAGG = 0.0                # Tabell E: 0,05 om projektet kan utföras vintertid
TRAFIK_FALT = "TRAFIKPASLAG"       # procentpåslag för trafiktillägg (Tabell E: 10/30/60 %) – fält på områdespolygonerna,
                                   # som Ledningsinventering lägger på ledningarna i området (eller ett fält på ledningen)
PROVISORIER_FALT = "PROVISORIER"   # procentpåslag för provisorier – på samma sätt

SYSTEM_NAMN = {"S": "Spillvatten", "D": "Dagvatten", "K": "Kombinerat", "V": "Vatten", "": "Okänt"}
METODER = (("omläggning", "_omlagg"), ("strumpa", "_strumpa"), ("rörspräckning", "_sprack"))
METOD_NYCKEL = dict(METODER)


# ============================================================================
# Kostnader
# ============================================================================

KOLUMNER = ["post", "schakttyp", "ledningar", "storre", "system", "dimension_fran", "dimension_till",
            "enhet", "kr", "kommentar"]


def _tal(v: str | None) -> float | None:
    v = re.sub(r"[^\d,.\-]", "", v or "")
    if "," in v:
        v = v.replace(".", "").replace(",", ".")
    try:
        return float(v) if v else None
    except ValueError:
        return None


def las_kostnader(path: str) -> list[dict]:
    """Läser kostnadsfilen. Rubrikraden (som börjar med 'post') styr kolumnordningen; utan rubrikrad
    gäller KOLUMNER. Returnerar [{post, schakttyp, ledningar, storre, system (mängd), dim_fran, dim_till,
    enhet, kr, kommentar}]."""
    poster, kolumner = [], None
    for rad in las_text(path).splitlines():
        rad = rad.split("#", 1)[0].strip()
        if not rad:
            continue
        d = [x.strip() for x in re.split(r"[;\t]", rad)]
        if d[0].lower() == "post":
            kolumner = [k.lower() for k in d]
            continue
        kol = kolumner or KOLUMNER
        v = dict(zip(kol, d + [""] * (len(kol) - len(d))))
        if not v.get("post"):
            continue
        kr = _tal(v.get("kr"))
        if kr is None:
            print(f"  VARNING {os.path.basename(path)}: kan inte tolka beloppet '{v.get('kr')}' för {v['post']} – hoppas över")
            continue
        n, k = _tal(v.get("ledningar")), _tal(v.get("storre"))
        system = {s.strip().upper() for s in re.split(r"[,+/ ]", v.get("system") or "") if s.strip()}
        poster.append({"post": v["post"].lower(), "schakttyp": (v.get("schakttyp") or "").lower() or None,
                       "ledningar": int(n) if n is not None else None, "storre": int(k) if k is not None else None,
                       "system": system or None, "dim_fran": _tal(v.get("dimension_fran")),
                       "dim_till": _tal(v.get("dimension_till")), "enhet": v.get("enhet", ""), "kr": kr,
                       "kommentar": v.get("kommentar", "")})
    return poster


def _dim_ok(k: dict, dim: float | None) -> bool:
    if k["dim_fran"] is None and k["dim_till"] is None:
        return True
    if dim is None:
        return False
    return (k["dim_fran"] is None or dim >= k["dim_fran"]) and (k["dim_till"] is None or dim <= k["dim_till"])


def hitta(kost: list[dict], post: str, dim: float | None = None, schakttyp: str | None = None,
          system: str | None = None, ledningar: int | None = None, storre: int | None = None) -> dict | None:
    """Första raden för posten som passar. En tom cell i raden gäller alla värden."""
    for k in kost:
        if k["post"] != post or not _dim_ok(k, dim):
            continue
        if k["schakttyp"] is not None and k["schakttyp"] != schakttyp:
            continue
        if k["system"] is not None and (system or "").upper() not in k["system"]:
            continue
        if k["ledningar"] is not None and k["ledningar"] != ledningar:
            continue
        if k["storre"] is not None and k["storre"] != storre:
            continue
        return k
    return None


def kr(kost: list[dict], post: str, **kw) -> float | None:
    r = hitta(kost, post, **kw)
    return r["kr"] if r else None


def gravpris(kost, schakttyp: str, n: int, storre: int, dmax: float | None) -> float | None:
    """kr per meter grav med n ledningar varav `storre` över STORRE_DIM, största dimension dmax."""
    if n <= 1:
        return kr(kost, "schakt", dim=dmax, schakttyp=schakttyp, ledningar=1)
    return kr(kost, "schakt", dim=dmax, schakttyp=schakttyp, ledningar=n, storre=storre)


# ============================================================================
# Beräkning per ledning
# ============================================================================

def _djup_punkter(led: dict) -> list[float]:
    """Djupet i profilens provpunkter (mark − vattengång)."""
    ut = []
    for rad in led.get("profil") or []:
        if len(rad) >= 3 and rad[1] is not None and rad[2] is not None:
            ut.append(rad[1] - rad[2])
    return ut


def grav_delar(led: dict) -> list[dict]:
    """Ledningens delsträckor med de andra ledningarna i samma grav: [{langd_m, ledningar: [id]}]. Går två
    ledningar parallellt 50 m och den ena fortsätter ensam 20 m blir det två delar. Ur Ledningsinventering
    (grav_delar); i äldre ledningsdata.json uppskattas delarna ur parallella: partner som går parallellt
    minst PARALLELL_MIN_M ligger intill på sin andel av längden, de med störst andel överlappar varandra."""
    if led.get("grav_delar") is not None:
        return led["grav_delar"]
    L = led.get("langd_m") or 0.0
    par = sorted(((p.get("ledning"), min(1.0, p.get("andel") or 0.0)) for p in led.get("parallella") or []
                  if (p.get("gemensam_m") or 0.0) >= PARALLELL_MIN_M - 1e-9), key=lambda x: -x[1])
    a = [x for _, x in par] + [0.0]
    delar = [{"langd_m": L * (a[k - 1] - a[k]), "ledningar": [q for q, _ in par[:k]]}
             for k in range(len(par), 0, -1) if a[k - 1] - a[k] > 1e-9]
    if 1.0 - a[0] > 1e-9 or not delar:
        delar.append({"langd_m": L * (1.0 - a[0]), "ledningar": []})
    return delar


def _gravpartners(led: dict, per_id: dict, medlagda: set | None = None) -> list[tuple[dict, float]]:
    """Ledningar som ligger i samma grav på någon del av ledningen och också läggs om: (ledning, andel av den
    här ledningens längd), störst andel först. medlagda = id:n som läggs om i alternativet (None = alla)."""
    L = led.get("langd_m") or 0.0
    delad = defaultdict(float)
    for d in grav_delar(led):
        for x in d.get("ledningar") or []:
            delad[x] += d.get("langd_m") or 0.0
    ut = [(per_id[x], (m / L) if L > 0 else 0.0) for x, m in delad.items()
          if x in per_id and x != led.get("id") and (medlagda is None or x in medlagda)]
    return sorted(ut, key=lambda x: (-x[1], str(x[0].get("id"))))


def _i_grav(led: dict, del_: dict, per_id: dict, medlagda: set | None, rang: dict) -> list[dict]:
    """De andra ledningarna i graven på delsträckan som också läggs om (högst MAX_LEDNINGAR_I_GRAV − 1,
    de med längst gemensam sträcka)."""
    med = [per_id[x] for x in del_.get("ledningar") or []
           if x in per_id and x != led.get("id") and (medlagda is None or x in medlagda)]
    return sorted(med, key=lambda q: rang.get(q.get("id"), 99))[:MAX_LEDNINGAR_I_GRAV - 1]


def ledningar_i_grav(led: dict, per_id: dict | None = None, medlagda: set | None = None) -> int:
    """Flest ledningar i graven på någon delsträcka (1 = ensam hela vägen, högst MAX_LEDNINGAR_I_GRAV)."""
    n = 1
    for d in grav_delar(led):
        if (d.get("langd_m") or 0) <= 0:
            continue
        if per_id is None:
            k = len(d.get("ledningar") or [])
        else:
            k = sum(1 for x in d.get("ledningar") or []
                    if x in per_id and x != led.get("id") and (medlagda is None or x in medlagda))
        n = max(n, 1 + k)
    return min(MAX_LEDNINGAR_I_GRAV, n)


def _antal_djuptillagg(djup: float) -> int:
    if djup is None or djup <= STANDARDDJUP_M:
        return 0
    return int(math.ceil((djup - STANDARDDJUP_M) / DJUPTILLAGG_STEG_M - 1e-9))


def schaktsatt(djup: float | None) -> str:
    """Schakttyp djup: ledningschakt, spontlåda eller slagen spont efter gravens medeldjup."""
    if djup is None:
        return "djup_saknas"
    if djup >= SLAGEN_SPONT_FRAN_M:
        return "slagen_spont"
    if djup >= SPONTLADA_FRAN_M:
        return "spontlada"
    return "ledningschakt"


def _grav(led, kost, per_id, schakttyp, medlagda=None):
    """Huvudledningens del av graven: ({schakt, djuptillagg, aterstallning, brunnar, spont, schaktsatt,
    pris_typ, gravdjup, delad_m}, brist, n, beskrivning). Räknas delsträcka för delsträcka (grav_delar): där
    ledningen ligger i samma grav som andra ledningar som läggs om gäller priset för 2–3 ledningar delat på
    antalet, där den går ensam priset för en ledning. Schakttyp djup: gravens medeldjup (djupaste ledningen i
    graven på delsträckan) avgör schaktsättet; slagen spont = SPONT_SIDOR × (medeldjup + SPONT_NEDSLAGNING_M)
    × längd m², delad på antalet ledningar. n = flest ledningar i graven på någon delsträcka."""
    L = led.get("langd_m") or 0.0
    system = led.get("system") or ""
    rang = {q.get("id"): k for k, (q, _) in enumerate(_gravpartners(led, per_id, medlagda))}
    delar = [(d, _i_grav(led, d, per_id, medlagda, rang)) for d in grav_delar(led) if (d.get("langd_m") or 0) > 0]
    summa_delar = sum(d["langd_m"] for d, _ in delar)
    skala = L / summa_delar if summa_delar > 0 else 0.0       # delarnas avrundade längder summerar till L
    n_max = max([1 + len(med) for _, med in delar] or [1])
    # Djuptillägg per påbörjade DJUPTILLAGG_STEG_M under STANDARDDJUP_M, medel över profilens punkter
    djupen = _djup_punkter(led) if system != "V" else []
    if not djupen and led.get("djup_medel") is not None:
        djupen = [led["djup_medel"]]
    steg = (sum(_antal_djuptillagg(d) for d in djupen) / len(djupen)) if djupen else 0.0
    o = {"schakt": 0.0, "spont": 0.0, "spont_m2": 0.0, "djuptillagg": 0.0, "aterstallning": 0.0,
         "schaktsatt": None, "pris_typ": schakttyp, "gravdjup": None, "delad_m": 0.0, "poster": []}
    beskr_delar, noter, utan_pris = [], [], defaultdict(float)
    if led.get("dimension") is None:
        return None, "dimension saknas", n_max, ""
    for d, med in delar:
        lm = d["langd_m"] * skala
        utan = [q for q in med if q.get("dimension") is None]
        if utan:
            noter.append("dimension saknas på " + ", ".join(str(q.get("id")) for q in utan) +
                         " i samma grav – graven räknad utan den")
            med = [q for q in med if q.get("dimension") is not None]
        while True:
            n = 1 + len(med)
            i_grav = [led] + med
            dims = [q.get("dimension") for q in i_grav]
            dmax = max(dims)
            storre = sum(1 for x in dims if x > STORRE_DIM)
            satt, gdjup, pris_typ = None, None, schakttyp
            if schakttyp == SCHAKTTYP_DJUP:
                djupen_g = [q.get("djup_medel") for q in i_grav if q.get("djup_medel") is not None]
                gdjup = max(djupen_g) if djupen_g else None
                satt = schaktsatt(gdjup)
                pris_typ = "spontslade" if satt == "spontlada" else "ledningschakt"
            pris_m = gravpris(kost, pris_typ, n, storre, dmax)
            beskr = f"{n} ledning{'ar' if n > 1 else ''}" + (f", {storre} större" if n > 1 else "") + f", Ø {dmax:g}"
            if pris_m is not None or not med:
                break
            # prislistan saknar kombinationen: den största av de andra ledningarna räknas i egen grav här
            storst = max(med, key=lambda q: q.get("dimension"))
            utan_pris[(beskr, storst.get("id"))] += lm
            med = [q for q in med if q is not storst]
        if pris_m is None:
            return None, f"inget schaktpris ({beskr}, {SCHAKTTYPER.get(pris_typ, pris_typ)})", n_max, beskr
        o["schakt"] += pris_m / n * lm
        o["poster"].append({"post": "schakt", "typ": pris_typ, "n": n, "storre": storre, "dmax": dmax,
                            "pris": pris_m, "m": lm / n})        # ledningens del av gravmetrarna (kalkylmallen)
        if satt == "slagen_spont":
            pris_m2 = kr(kost, "slagen_spont")
            if pris_m2 is None:
                return None, "pris för slagen spont saknas i kostnadsfilen", n_max, beskr
            m2 = SPONT_SIDOR * (gdjup + SPONT_NEDSLAGNING_M) * lm / n
            o["spont_m2"] += m2
            o["spont"] += m2 * pris_m2
        o["djuptillagg"] += steg * (kr(kost, "djuptillagg", dim=dmax) or 0.0) / n * lm
        if steg > 0:
            o["poster"].append({"post": "djuptillagg", "dmax": dmax, "pris": kr(kost, "djuptillagg", dim=dmax) or 0.0,
                                "m": steg * lm / n})
        o["aterstallning"] += (kr(kost, "aterstallning") or 0.0) / n * lm
        if n > 1:
            o["delad_m"] += lm
        if o["schaktsatt"] is None or (gdjup is not None and (o["gravdjup"] is None or gdjup > o["gravdjup"])):
            o["schaktsatt"], o["pris_typ"], o["gravdjup"] = satt, pris_typ, gdjup   # djupaste delsträckan
        beskr_delar.append(f"{beskr} {lm:.0f} m".replace(".", ","))
    andar = sum(0.5 for b in (led.get("fran"), led.get("till")) if b) if system != "V" else 0.0
    o["brunnar"] = andar * (kr(kost, "brunn_ny") or 0.0)
    brist = "" if djupen or system == "V" else "djup saknas – inget djuptillägg"
    if o["schaktsatt"] == "djup_saknas":
        brist = "djup saknas – ingen spont, inget djuptillägg"
    noter += [f"inget schaktpris för {b} på {_dk(round(m, 1))} m – {q} räknad i egen grav där"
              for (b, q), m in utan_pris.items()]
    brist = "; ".join(x for x in [brist] + list(dict.fromkeys(noter)) if x)
    return o, brist, n_max, "; ".join(beskr_delar)


def _servis_m(s: dict) -> float:
    L = s.get("langd_m") or 0.0
    return min(L, SERVIS_LANGD_MAX_M) if SERVIS_LANGD_MAX_M is not None else L


def servisgravar(serviser: list[dict]) -> dict[int, int]:
    """{index i serviser: gruppnummer}: serviser vars båda ändar ligger inom SERVIS_GRAV_AVSTAND_M från
    varandras ändar (i någon riktning) delar servisgrav – t.ex. vatten-, spill- och dagservis till samma hus."""
    def avst(a, b):
        return math.hypot(a[0] - b[0], a[1] - b[1])
    foralder = list(range(len(serviser)))

    def rot(i):
        while foralder[i] != i:
            foralder[i] = foralder[foralder[i]]
            i = foralder[i]
        return i
    andar = [(s.get("fran_xy"), s.get("till_xy")) for s in serviser]
    t = SERVIS_GRAV_AVSTAND_M
    cell = defaultdict(list)
    for i, (a, b) in enumerate(andar):
        if a and b:
            cell[(int(a[0] // t), int(a[1] // t))].append(i)
            cell[(int(b[0] // t), int(b[1] // t))].append(i)
    for i, (a, b) in enumerate(andar):
        if not (a and b):
            continue
        kand = set()
        for p in (a, b):
            cx, cy = int(p[0] // t), int(p[1] // t)
            for dx in (-1, 0, 1):
                for dy in (-1, 0, 1):
                    kand.update(cell.get((cx + dx, cy + dy), ()))
        for j in kand:
            if j <= i:
                continue
            c, d = andar[j]
            if max(avst(a, c), avst(b, d)) <= t or max(avst(a, d), avst(b, c)) <= t:
                ri, rj = rot(i), rot(j)
                if ri != rj:
                    foralder[rj] = ri
    return {i: rot(i) for i in range(len(serviser))}


def _servis_helschakt(egna, alla, grupp, metod_for, kost, schakttyp, metod="omläggning", poster=None):
    """Serviskostnad i helschakt för en lednings serviser. En servis delar servisgrav med de serviser i
    samma grupp vars huvudledning också läggs om i alternativet; gravens pris delas lika mellan dem.
    grupp = (index -> gruppnummer, gruppnummer -> index)."""
    gmap, medl = grupp
    summa = 0.0
    tillagg = kr(kost, "servis_tillagg", schakttyp=schakttyp) or 0.0
    for i in egna:
        s = alla[i]
        med = [j for j in medl[gmap[i]] if metod_for(alla[j]) == metod]
        m = max(1, min(MAX_LEDNINGAR_I_GRAV, len(med)))
        L = max(_servis_m(alla[j]) for j in med) if med else _servis_m(s)
        pris_m = kr(kost, "servis_helschakt", ledningar=m, system=s.get("system")) or 0.0
        summa += (pris_m + tillagg) * L / max(1, len(med))
        if poster is not None:
            poster.append({"pris": pris_m, "system": s.get("system"), "m": L / max(1, len(med)),
                           "tillagg": tillagg, "tillagg_typ": schakttyp if tillagg else None})
    return summa


def servis_schaktfritt(serv, kost) -> float:
    """Serviserna när huvudledningen infodras eller rörspräcks: spill-, dag- och kombinerade serviser
    upp till SERVIS_HATT_MAX_M infodras med hatt (post servis_hatt, kr/st), längre serviser och
    vattenserviser per meter (servis_infodring)."""
    summa = 0.0
    for x in serv:
        hatt = kr(kost, "servis_hatt", system=x.get("system")) if _hatt(x) else None
        if hatt is not None:
            summa += hatt
        else:
            summa += (kr(kost, "servis_infodring", system=x.get("system")) or 0.0) * _servis_m(x)
    return summa


def _hatt(x: dict) -> bool:
    return (x.get("system") or "") in SERVIS_HATT_SYSTEM and (x.get("langd_m") or 0.0) <= SERVIS_HATT_MAX_M


def antal_hattar(serv) -> int:
    return sum(1 for x in serv if _hatt(x))


def _infodring(led, kost, brunnar, serv):
    """Infodring av huvudledning + serviser schaktfritt (+ framschaktning)."""
    L = led.get("langd_m") or 0.0
    dim = led.get("dimension")
    if (led.get("system") or "") not in STRUMPA_SYSTEM:
        return None, "vattenledning – infodras inte" if led.get("system") == "V" else "okänt system"
    if dim is None:
        return None, "dimension saknas"
    if dim < STRUMPA_DIM_MIN:
        return None, f"dimension {dim:g} mm under {STRUMPA_DIM_MIN}"
    kr_m = kr(kost, "infodring", dim=dim)
    if kr_m is None:
        return None, f"inget infodringspris för Ø {dim:g}"
    s = {"per_meter": kr_m, "infodring": kr_m * L,
         "serviser": servis_schaktfritt(serv, kost)}
    typer = [(brunnar.get(b) or {}).get("typ") or "" for b in (led.get("fran"), led.get("till"))]
    if FRAMSCHAKTA_TB and all(typer) and all(re.search(BRUNNSTYP_TILLSYN, t.upper()) for t in typer):
        s["framschaktning"] = (kr(kost, "framschaktning") or 0.0) + (kr(kost, "ny_brunn") or 0.0)
    s["summa"] = s["infodring"] + s["serviser"] + s.get("framschaktning", 0.0)
    return s, ""


def _sprackning(led, kost, serv):
    """Rörspräckning av huvudledning + gropar + servisgrop och infodring per servis + ventilgrop per
    ventil och brandpost på ledningen (VENTILGROP_KATEGORIER, ur inventeringens anordningar)."""
    L = led.get("langd_m") or 0.0
    dim = led.get("dimension")
    if dim is None:
        return None, "dimension saknas"
    if dim > SPRACKNING_DIM_MAX:
        return None, f"dimension {dim:g} mm > {SPRACKNING_DIM_MAX}"
    kr_m = kr(kost, "rorsprackning", dim=dim)
    if kr_m is None:
        return None, f"inget pris för rörspräckning Ø {dim:g}"
    gropar = max(1, int(math.ceil(L / GROP_PER_M - 1e-9)))
    r = {"per_meter": kr_m, "sprackning": kr_m * L, "antal_gropar": gropar,
         "gropar": gropar * ((kr(kost, "inforingsgrop") or 0.0) + (kr(kost, "maskingrop") or 0.0)),
         "servisgropar": len(serv) * (kr(kost, "servisgrop") or 0.0),
         "antal_ventilgropar": antal_ventilgropar(led),
         "serviser": servis_schaktfritt(serv, kost)}
    r["ventilgropar"] = r["antal_ventilgropar"] * (kr(kost, "ventilgrop") or kr(kost, "servisgrop") or 0.0)
    r["summa"] = r["sprackning"] + r["gropar"] + r["servisgropar"] + r["ventilgropar"] + r["serviser"]
    return r, ""


def antal_ventilgropar(led: dict) -> int:
    antal = led.get("anordningar_antal") or {}
    return sum(antal.get(k, 0) for k in VENTILGROP_KATEGORIER)


def _procent(v) -> float | None:
    """Procentpåslag: 10, "10 %" eller 0,10 -> 0.10; tomt -> None."""
    if v is None or v == "":
        return None
    if isinstance(v, (int, float)):
        x = float(v)
    else:
        m = re.search(r"-?\d+(?:[.,]\d+)?", str(v))
        if not m:
            return None
        x = float(m.group(0).replace(",", "."))
    return x / 100.0 if x > 1 else x


def _paslag(l: dict, nyckel: str, falt: str) -> float | None:
    """Ledningens procentpåslag: från områdets fält (Ledningsinventering satte l[nyckel]), annars ett fält med
    samma namn på ledningen."""
    if l.get(nyckel) is not None:
        return _procent(l[nyckel])
    attr = l.get("attribut") or {}
    return _procent(next((v for k, v in attr.items() if falt and str(k).upper() == falt.upper()), None))


def trafikpaslag(l: dict) -> float | None:
    return _paslag(l, "trafikpaslag", TRAFIK_FALT)


def provisorier(l: dict) -> float | None:
    return _paslag(l, "provisorier", PROVISORIER_FALT)


def _anordningar(l: dict, kategori: str | tuple) -> int | None:
    """Antal anordningar av kategorin på ledningen; None när filen saknar anordningar."""
    if "anordningar_antal" not in l:
        return None
    k = (kategori,) if isinstance(kategori, str) else kategori
    return sum((l["anordningar_antal"] or {}).get(x, 0) for x in k)


def _omlaggning(led, kost, per_id, schakttyp, medlagda, serv_idx, alla_serv, grupp, metod_for):
    o, brist, n, beskr = _grav(led, kost, per_id, schakttyp, medlagda)
    if o is None:
        return None, brist, n, beskr
    o["servis_poster"] = []
    o["serviser"] = _servis_helschakt(serv_idx, alla_serv, grupp, metod_for, kost, o["pris_typ"],
                                      poster=o["servis_poster"])
    o["summa"] = o["schakt"] + o["spont"] + o["djuptillagg"] + o["aterstallning"] + o["brunnar"] + o["serviser"]
    return o, brist, n, beskr


def _nod(led: dict, ande: str):
    """Nyckel för ledningens ände: brunnens littera, annars den fria ändens koordinat."""
    namn = led.get(ande)
    if namn:
        return namn
    xy = led.get(ande + "_xy")
    return ("xy", round(xy[0], 1), round(xy[1], 1)) if xy and xy[0] is not None else None


def _upp_ned(led: dict):
    """(uppströms ände, nedströms ände) – ur vattengången, annars ritad riktning."""
    a, b = _nod(led, "fran"), _nod(led, "till")
    vf, vt = led.get("vg_fran"), led.get("vg_till")
    if vf is not None and vt is not None and vt > vf:
        return b, a
    return a, b


def anta_dimensioner(ledningar: list[dict]) -> int:
    """Ledningar utan dimension får uppströms ledningens (samma system, ansluter i den uppströms änden;
    flera -> den största). Kedjor fylls steg för steg uppifrån. Utan uppströms dimension tas nedströms
    ledningens (DIMENSION_NEDSTROMS_RESERV). Sätter dimension_antagen = {fran, riktning, dimension}."""
    for l in ledningar:
        l.setdefault("dimension_gis", l.get("dimension"))
        if not l.get("dimension"):                  # 0 = okänd i GIS (platshållare), räknas som saknad
            l["dimension"] = None
    saknas = [l for l in ledningar if l.get("dimension") is None]
    if not DIMENSION_UR_GRANNE or not saknas:
        return 0
    slut_i = defaultdict(list)     # nod -> ledningar vars nedströms ände ligger där
    start_i = defaultdict(list)    # nod -> ledningar vars uppströms ände ligger där
    riktn = {}
    for l in ledningar:
        u, n = _upp_ned(l)
        riktn[id(l)] = (u, n)
        if n is not None:
            slut_i[(l.get("system"), n)].append(l)
        if u is not None:
            start_i[(l.get("system"), u)].append(l)
    steg = [("uppströms", lambda l: slut_i.get((l.get("system"), riktn[id(l)][0]), []))]
    if DIMENSION_NEDSTROMS_RESERV:
        steg.append(("nedströms", lambda l: start_i.get((l.get("system"), riktn[id(l)][1]), [])))

    def fyll(riktning, grannar) -> int:
        n = 0
        for l in saknas:
            if l.get("dimension") is not None:
                continue
            kand = [q for q in grannar(l) if q is not l and q.get("dimension") is not None]
            if not kand:
                continue
            q = max(kand, key=lambda x: x["dimension"])
            l["dimension"] = q["dimension"]
            kalla = (q.get("dimension_antagen") or {}).get("ursprung") or q.get("id")
            l["dimension_antagen"] = {"fran": q.get("id"), "riktning": riktning, "dimension": q["dimension"],
                                      "ursprung": kalla}
            n += 1
        return n
    # uppströms först; nedströms bara när uppströms inte ger något mer – sedan uppströms igen, eftersom
    # en ledning som fått nedströms dimension kan vara uppströms granne till en annan
    antal = 0
    while True:
        n = 0
        for riktning, grannar in steg:
            n = fyll(riktning, grannar)
            if n:
                break
        if not n:
            return antal
        antal += n


def _dim_text(l: dict) -> str | None:
    a = l.get("dimension_antagen")
    if not a:
        return None
    t = f"dimension saknas i GIS – Ø {a['dimension']:g} antagen från {a['riktning']} ledning {a['fran']}"
    if a.get("ursprung") and a["ursprung"] != a["fran"]:
        t += f" (ursprungligen {a['ursprung']})"
    return t


def _tal_text(v: float) -> str:
    """Som inventering._heltal_text: 2 -> '2', 2.5 -> '2,5'."""
    return f"{v:.0f}" if float(v).is_integer() else f"{v:.1f}".replace(".", ",")


def _djupklass(djup: float, granser) -> str:
    """Samma klassnamn som inventering._djupklass ('0–2', '2,5–3', '> 4')."""
    forra = 0.0
    for g in granser:
        if djup < g:
            return f"{_tal_text(forra)}–{_tal_text(g)}"
        forra = g
    return f"> {_tal_text(forra)}"


def vatten_djup_till_botten(data: dict) -> int:
    """Vattenledningarnas schablondjup räknas till rörets botten = djupet till centrum (vatten_djup_m) +
    dimensionen/2, som avloppets djup till vattengången. Görs efter anta_dimensioner så att även antagna
    dimensioner räknas med, och rättar äldre ledningsdata.json där djupet gällde centrum. Antal ledningar."""
    p = data.get("parametrar") or {}
    centrum = p.get("vatten_djup_m")
    if centrum is None:
        return 0
    granser = p.get("djupklasser")
    if not granser:
        granser = (2.0, 3.0, 4.0)
        print("  OBS: ledningsdata.json saknar djupklassgränserna – 2, 3 och 4 m antas")
    n = 0
    for led in data.get("ledningar", []):
        if led.get("system") == "V" and led.get("djup_kalla") == "schablon":
            d = round(float(centrum) + (led.get("dimension") or 0) / 2000.0, 3)
            led["djup_medel"] = led["djup_min"] = led["djup_max"] = d
            led["djupklass"] = _djupklass(d, granser)
            n += 1
    p["vatten_djup_till"] = "botten"
    data["parametrar"] = p
    return n


def berakna(data: dict, kost: list[dict], schakttyp: str = SCHAKTTYP) -> list[dict]:
    """Lägger till kostnaderna på varje ledning (i data['ledningar'])."""
    ledningar = data.get("ledningar", [])
    for led in ledningar:
        if not led.get("system"):            # gisdata.json från Exportera GIS-data saknar system
            led["system"] = (led.get("ledningstyp") or "").strip()[:1].upper()
        if not led.get("id"):
            led["id"] = f"{(led.get('lager') or '').split(chr(92))[-1]}/{led.get('oid')}"
    data["_antagna_dim"] = anta_dimensioner(ledningar)
    vatten_djup_till_botten(data)
    per_id = {l.get("id"): l for l in ledningar}
    brunnar = {b["littera"]: b for b in data.get("brunnar", [])}
    alla_serv = [s for s in data.get("serviser", []) if s.get("ledning") in per_id]
    for s in alla_serv:
        s.setdefault("system", per_id[s["ledning"]].get("system"))
    gmap = servisgravar(alla_serv)
    medl = defaultdict(list)
    for i, g in gmap.items():
        medl[g].append(i)
    grupp = (gmap, medl)
    for i, s in enumerate(alla_serv):
        s["_servisgrav"] = len(medl[gmap[i]])
    serv_idx = defaultdict(list)
    for i, s in enumerate(alla_serv):
        serv_idx[s["ledning"]].append(i)

    def metod_a(s):                           # Alt. lägg om: alla läggs om
        return "omläggning"
    for led in ledningar:
        egna = serv_idx.get(led.get("id"), [])
        serv = [alla_serv[i] for i in egna]
        led["_typer"] = {}
        for t in ALLA_SCHAKTTYPER:
            led["_typer"][t] = _omlaggning(led, kost, per_id, t, None, egna, alla_serv, grupp, metod_a)[0]
        led["_omlagg"], led["_omlagg_brist"], led["_grav"], led["_grav_hur"] = \
            _omlaggning(led, kost, per_id, schakttyp, None, egna, alla_serv, grupp, metod_a)
        led["_strumpa"], led["_strumpa_brist"] = _infodring(led, kost, brunnar, serv)
        led["_sprack"], led["_sprack_brist"] = _sprackning(led, kost, serv)
        led["_anm"] = "; ".join(f"{namn}: {led[k + '_brist']}" for namn, k in METODER
                                if led.get(k + "_brist") and not (k == "_strumpa" and led.get("system") == "V"))
    alternativ_strumpa(ledningar, per_id, kost, schakttyp, alla_serv, grupp, serv_idx)
    data["_serviser_kopplade"] = alla_serv
    return ledningar


def alternativ_strumpa(ledningar, per_id, kost, schakttyp, alla_serv, grupp, serv_idx) -> None:
    """Metod och kostnad per ledning i alternativet schaktfritt (_alt_metod, _alt, _alt_grav).
    Avlopp: infodring, annars rörspräckning, annars omläggning. Vatten: läggs om tillsammans med avloppet när
    minst VATTEN_OMLAGGNING_ANDEL av dess längd ligger i samma grav som en avloppsledning som läggs om, annars
    rörspräckning."""
    for led in ledningar:
        if led.get("system") == "V":
            continue
        led["_alt_metod"] = "strumpa" if led.get("_strumpa") else ("rörspräckning" if led.get("_sprack") else "omläggning")
    omlagda_avlopp = {l.get("id") for l in ledningar if l.get("system") != "V" and l["_alt_metod"] == "omläggning"}
    for led in ledningar:
        if led.get("system") != "V":
            continue
        schakt = sorted({x for d in grav_delar(led) for x in d.get("ledningar") or [] if x in omlagda_avlopp})
        L = led.get("langd_m") or 0.0
        intill = sum(d.get("langd_m") or 0.0 for d in grav_delar(led)
                     if any(x in omlagda_avlopp for x in d.get("ledningar") or []))
        if schakt and L > 0 and intill >= VATTEN_OMLAGGNING_ANDEL * L - 1e-6:
            led["_alt_metod"] = "omläggning"
            led["_alt_motiv"] = "läggs om i samma grav som " + ", ".join(str(x) for x in schakt)
        else:
            led["_alt_metod"] = "rörspräckning" if led.get("_sprack") else "omläggning"
    medlagda = {l.get("id") for l in ledningar if l["_alt_metod"] == "omläggning"}

    def metod_b(s):
        return per_id[s["ledning"]]["_alt_metod"]
    for led in ledningar:
        metod = led["_alt_metod"]
        if metod == "omläggning":
            o, brist, n, _ = _omlaggning(led, kost, per_id, schakttyp, medlagda, serv_idx.get(led.get("id"), []),
                                         alla_serv, grupp, metod_b)
            led["_alt"], led["_alt_brist"], led["_alt_grav"] = o, (brist if o is None else ""), n
        else:
            led["_alt"], led["_alt_brist"], led["_alt_grav"] = led[METOD_NYCKEL[metod]], "", None


# ============================================================================
# Excel
# ============================================================================

def _dela_rubrik(k: str) -> tuple[str, str]:
    m = re.match(r"^(.*?)\s*\((.+)\)$", k)
    return (m.group(1), m.group(2)) if m else (k, "")


def _formel(v) -> bool:
    return isinstance(v, str) and v.startswith("=")


class Kolumner:
    """Områden i en flik för formler: k("Längd") -> 'Ledningar'!$E$3:$E$187 (rubrik utan enhet)."""

    def __init__(self, blad: str, kolumner: list[str], forsta: int, antal: int):
        from openpyxl.utils import get_column_letter, quote_sheetname
        self.blad, self.forsta, self.sista = quote_sheetname(blad), forsta, forsta + max(antal, 1) - 1
        self.bokstav = {}
        for i, k in enumerate(kolumner, 1):
            self.bokstav.setdefault(_dela_rubrik(k)[0], get_column_letter(i))

    def __call__(self, namn: str) -> str:
        b = self.bokstav[namn]
        return f"{self.blad}!${b}${self.forsta}:${b}${self.sista}"

    def cell(self, namn: str, rad: int) -> str:
        """Cell på samma rad i samma flik (radformler)."""
        return f"{self.bokstav[namn]}{rad}"


EJ_TOM = '"<>"'


def _txt(v) -> str:
    """Värde i en formel: tal som tal, tomt som "=" (villkoret tom cell i SUMIFS/COUNTIFS), annars text."""
    if v is None or v == "":
        return '"="'
    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"
    if isinstance(v, (int, float)):
        return repr(v)
    return '"' + str(v).replace('"', '""') + '"'


def _lika(omrade: str, v) -> str:
    """SUMPRODUCT-villkor: cellen är exakt v (tomt = tom cell)."""
    return f'--({omrade}={_txt(v) if v not in (None, "") else chr(34) * 2})'


def _sumifs(summa: str, krit: list) -> str:
    """=SUMIFS över summaområdet med villkoren [(område, värde i formeln)], utan villkor SUM."""
    if not krit:
        return f"SUM({summa})"
    return "SUMIFS(" + summa + "," + ",".join(f"{r},{v}" for r, v in krit) + ")"


def _countifs(alla: str, krit: list) -> str:
    """Antal rader med villkoren; utan villkor COUNTA över en kolumn som alltid är ifylld."""
    if not krit:
        return f"COUNTA({alla})"
    return "COUNTIFS(" + ",".join(f"{r},{v}" for r, v in krit) + ")"


def tabell(ws, kolumner: list[str], rader: list[list], bredder: dict | None = None, dolda=None,
           format: dict | None = None, rad_start: int = 1) -> int:
    """Rubrikrad + enhetsrad ("Längd (m)" -> rubrik Längd, enhet m) + data, som i tv3_analys.
    format: {rubrik utan enhet: number_format}. Returnerar första dataraden."""
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter
    tunn = Side(style="thin", color="D9D9D9")
    kant = Border(top=tunn, bottom=tunn, left=tunn, right=tunn)
    namn = [_dela_rubrik(k)[0] for k in kolumner]
    enheter = [_dela_rubrik(k)[1] for k in kolumner]
    r = rad_start
    for i, n in enumerate(namn, 1):
        c = ws.cell(r, i, n)
        c.font, c.fill, c.border = Font(bold=True, color="FFFFFF"), PatternFill("solid", fgColor="1F3864"), kant
        c.alignment = Alignment(vertical="center", wrap_text=True)
    har_enhet = any(enheter)
    if har_enhet:
        r += 1
        for i, e in enumerate(enheter, 1):
            c = ws.cell(r, i, e)
            c.font, c.fill, c.border = Font(italic=True, color="595959", size=9), PatternFill("solid", fgColor="EDEDED"), kant
            c.alignment = Alignment(horizontal="center")
    start = r + 1
    fmt = {i: (format or {}).get(namn[i - 1]) for i in range(1, len(namn) + 1)}
    for ri, rad in enumerate(rader):
        for ci, v in enumerate(rad, 1):
            c = ws.cell(start + ri, ci, v)
            c.border = kant
            if fmt.get(ci) and (isinstance(v, (int, float)) or _formel(v)):
                c.number_format = fmt[ci]
    if rad_start == 1:
        ws.freeze_panes = f"A{start}"
        ws.auto_filter.ref = f"A{start - 1}:{get_column_letter(len(kolumner))}{max(start + len(rader) - 1, start)}"
    for i, kol in enumerate(kolumner, 1):
        b = (bredder or {}).get(namn[i - 1])
        if b is None:
            langsta = max((len(str(rd[i - 1])) for rd in rader[:300]
                           if i - 1 < len(rd) and rd[i - 1] is not None and not _formel(rd[i - 1])), default=0)
            b = min(40, max(8, langsta + 2, len(namn[i - 1]) + 2, len(enheter[i - 1]) + 2))
        bokstav = get_column_letter(i)
        if (ws.column_dimensions[bokstav].width or 0) < b:
            ws.column_dimensions[bokstav].width = b
    dold = set(dolda or [])
    for i, n in enumerate(namn, 1):
        if n in dold:
            cd = ws.column_dimensions[get_column_letter(i)]
            cd.hidden = True
            cd.outlineLevel = 1
    if dold:
        ws.sheet_properties.outlinePr.summaryRight = False
    return start


def _dk(v) -> str:
    """Tal som text med decimalkomma (1.7 -> 1,7)."""
    return ("%g" % v).replace(".", ",") if isinstance(v, (int, float)) else ("" if v is None else str(v))


def _r(v, n=1):
    return round(v, n) if isinstance(v, (int, float)) else v


def _kr(d: dict | None, k: str = "summa"):
    return round(d[k]) if d and d.get(k) is not None else None


def _grav_text(l: dict) -> str | None:
    """Delsträckorna med ledningarna i samma grav: "0–94,5 m: DSL1, SSL1; 94,5–200 m: SSL2"."""
    delar = grav_delar(l)
    if not any(d.get("ledningar") for d in delar):
        return None
    ut = []
    for d in delar:
        vem = ", ".join(d.get("ledningar") or []) or "ensam"
        if d.get("fran_m") is not None:
            ut.append(f"{_dk(round(d['fran_m'], 1))}–{_dk(round(d['till_m'], 1))} m: {vem}")
        else:
            ut.append(f"{_dk(round(d.get('langd_m') or 0, 1))} m: {vem}")
    return "; ".join(ut)


def _anmarkning(l: dict) -> str | None:
    delar = [x for x in (_dim_text(l), l.get("_alt_motiv"), l.get("_anm"),
                         ("alt. schaktfritt: " + l["_alt_brist"]) if l.get("_alt_brist") else None) if x]
    return "; ".join(delar) or None


def _summa(lista, f):
    v = [f(l) for l in lista]
    v = [x for x in v if x is not None]
    return round(sum(v)) if v else None


def schaktlangd(lista: list[dict]) -> float:
    """Effektiv längd på schaktet för ledningarna i listan (en schaktgrupp): delar av en ledning som går i
    samschakt med andra ledningar i listan räknas en gång. Ledningsinventering räknar ledningens andel
    (schakt_langd_m) punkt för punkt längs ledningen; i äldre ledningsdata.json uppskattas den ur andelarna: andelen s = min(1, Σ andel) av
    längden ligger i samschakt, i genomsnitt med m = 1 + Σ andel / s ledningar; bidraget blir
    L × (1 − s) + L × s / m. Två lika långa parallella ledningar ger en längd, en 100 m och en 50 m
    längs halva den ger 100 m."""
    return sum(schaktandelar(lista))


def schaktandelar(lista: list[dict]) -> list[float]:
    """Varje lednings del av schaktlängden (se schaktlangd), i listans ordning."""
    if lista and all(l.get("schakt_langd_m") is not None for l in lista):
        return [l["schakt_langd_m"] for l in lista]       # räknat punkt för punkt i Ledningsinventering
    ids = {l.get("id") for l in lista}
    ut = []
    for l in lista:
        L = l.get("langd_m") or 0.0
        a = sum(p.get("andel") or 0.0 for p in l.get("parallella") or []
                if p.get("samschakt") and p.get("ledning") in ids)
        if a <= 0:
            ut.append(L)
            continue
        andel = min(1.0, a)
        m = 1.0 + a / andel
        ut.append(L * (1.0 - andel) + L * andel / m)
    return ut


def _jamforbar(l: dict) -> bool:
    """Ledningen har kostnad i båda alternativen – bara sådana summeras, så att summorna går att jämföra."""
    return bool(l.get("_omlagg")) and bool(l.get("_alt"))


def _altsummor(lista):
    """[alt. lägg om (kr), alt. schaktfritt (kr), strumpa (st), rörspräckning (st), omläggning (st), utan kostnad (st)]."""
    jfr = [l for l in lista if _jamforbar(l)]
    return [_summa(jfr, lambda l: l["_omlagg"]["summa"]),
            _summa(jfr, lambda l: l["_alt"]["summa"]),
            sum(1 for l in lista if l.get("_alt_metod") == "strumpa"),
            sum(1 for l in lista if l.get("_alt_metod") == "rörspräckning"),
            sum(1 for l in lista if l.get("_alt_metod") == "omläggning"),
            sum(1 for l in lista if not _jamforbar(l))]


def _radformler(rad: list, kol: list[str], K: Kolumner, nr: int, l: dict) -> None:
    """Radtotalerna i fliken Ledningar som formler av delposterna på samma rad – ändras ett pris eller en
    delpost i Excel räknas totalen och summorna i Sammanfattning om. Tomma celler (ingen kostnad) lämnas tomma."""
    plats = {_dela_rubrik(k)[0]: i for i, k in enumerate(kol)}

    def c(namn):
        return K.cell(namn, nr)

    def satt(namn, formel):
        if rad[plats[namn]] is not None:
            rad[plats[namn]] = "=" + formel
    satt("Infodring", f"{c('Infodring per meter')}*{c('Längd')}")
    satt("Infodring totalt", "+".join(c(x) for x in ("Infodring", "Serviser infodring", "Framschaktning")))
    satt("Rörspräckning", f"{c('Rörspräckning per meter')}*{c('Längd')}")
    satt("Rörspräckning totalt", "+".join(c(x) for x in ("Rörspräckning", "Gropar", "Servisgropar", "Ventilgropar",
                                                         "Serviser rörspräckning")))
    satt("Alt. lägg om", "+".join(c(x) for x in ("Schakt", "Spont", "Djuptillägg", "Återställning", "Brunnar",
                                                 "Serviser helschakt")))
    alt = l.get("_alt")
    if alt is not None and alt is l.get("_strumpa"):
        satt("Alt. schaktfritt", c("Infodring totalt"))
    elif alt is not None and alt is l.get("_sprack"):
        satt("Alt. schaktfritt", c("Rörspräckning totalt"))


def _jamforbar_krit(K: Kolumner) -> list:
    """Villkoren för _jamforbar: kostnad i båda alternativen."""
    return [(K("Alt. lägg om"), EJ_TOM), (K("Alt. schaktfritt"), EJ_TOM)]


def _gruppformler(K: Kolumner, krit: list) -> list[str]:
    """[ledningar (st), längd, schaktlängd, alt. lägg om, alt. schaktfritt, strumpas, rörspräcks, läggs om,
    utan kostnad] som formler över raderna i Ledningar som uppfyller villkoren."""
    jfr = krit + _jamforbar_krit(K)
    antal = _countifs(K("ID"), krit)
    return ["=" + antal, "=" + _sumifs(K("Längd"), krit), "=" + _sumifs(K("Schaktlängd"), krit),
            "=" + _sumifs(K("Alt. lägg om"), jfr), "=" + _sumifs(K("Alt. schaktfritt"), jfr)] + \
           ["=" + _countifs(K("ID"), krit + [(K("Metod alt. schaktfritt"), _txt(m))])
            for m in ("strumpa", "rörspräckning", "omläggning")] + \
           [f"={antal}-{_countifs(K('ID'), jfr)}"]


def skriv_excel(data: dict, ledningar: list[dict], path: str, kost: list[dict], kallor: dict) -> None:
    from openpyxl import Workbook
    from openpyxl.styles import Font

    wb = Workbook()
    ws_s = wb.active
    ws_s.title = "Sammanfattning"
    KR, M, M2, P = "#,##0", "0.0", "0.00", "0.0"

    # --- Ledningar ------------------------------------------------------------------------------
    attrnamn: list[str] = []
    for l in ledningar:
        for k in (l.get("attribut") or {}):
            if k not in attrnamn:
                attrnamn.append(k)
    ws = wb.create_sheet("Ledningar")
    kol = ["ID", "System", "Från brunn", "Till brunn", "Längd (m)", "Dimension (mm)", "Dimension antagen från",
           "Material", "Ledningstyp", "Anläggningsår", "Renoveringsår", "VG från (m)", "VG till (m)", "Lutning (‰)", "Lutning GIS (‰)",
           "Djup medel (m)", "Djup min (m)", "Djup max (m)", "Djupklass (m)", "Djupkälla", "Täckning min (m)",
           "Djup start (m)", "Djup slut (m)", "Serviser (st)", "Servislängd (m)", "Ventiler (st)",
           "Brandposter (st)", "Servisventiler (st)", "Ledningsändar (st)", "Övriga anordningar (st)", "Schaktgrupp",
           "Ledningar i graven (st)", "Delad grav (m)", "Grav längs ledningen", "Schaktlängd (m)", "Samschakt med", "Parallella", "Korsningar",
           "Alt. lägg om (kr)", "Schaktsätt alt. lägg om", "Metod alt. schaktfritt", "Alt. schaktfritt (kr)", "Anmärkning",
           "Trafiktillägg (%)", "Provisorier (%)",
           "Lägg om efter djup (kr)", "Lägg om ledningschakt (kr)", "Lägg om med spontsläde (kr)", "Lägg om bergschakt (kr)",
           "Lägg om KC-förstärkning (kr)", "Schaktsätt efter djup", "Slagen spont efter djup (m²)", "Spont efter djup (kr)",
           "Schakt (kr)", "Gravdjup (m)", "Slagen spont (m²)", "Spont (kr)", "Djuptillägg (kr)", "Återställning (kr)", "Brunnar (kr)", "Serviser helschakt (kr)",
           "Infodring per meter (kr/m)", "Infodring (kr)", "Serviser infodring (kr)", "Hattar (st)", "Framschaktning (kr)",
           "Infodring totalt (kr)",
           "Rörspräckning per meter (kr/m)", "Rörspräckning (kr)", "Gropar (kr)", "Servisgropar (kr)",
           "Ventilgropar (kr)", "Serviser rörspräckning (kr)", "Rörspräckning totalt (kr)", "Ledningar i graven alt. schaktfritt (st)",
           "Flaggor", "Område", "Lager", "OID", "Del", "Klippt vid områdesgränsen (m)"]
    egna = {_dela_rubrik(k)[0] for k in kol}
    attrkol = [a if a not in egna else a + " (GIS)" for a in attrnamn]
    rader = []
    hattar = defaultdict(int)
    for x in data.get("_serviser_kopplade") or []:
        if _hatt(x):
            hattar[x.get("ledning")] += 1
    grupper = defaultdict(list)
    for l in ledningar:
        grupper[l.get("schaktgrupp")].append(l)
    schaktdel = {}
    for lista in grupper.values():
        for l, v in zip(lista, schaktandelar(lista)):
            schaktdel[id(l)] = v
    ordn = sorted(ledningar, key=lambda l: ("SDKV".find(l.get("system") or "?") % 9, str(l.get("id"))))
    K = Kolumner("Ledningar", kol, 3, len(ordn))
    for nr, l in enumerate(ordn, 3):
        o, s, r = l.get("_omlagg"), l.get("_strumpa"), l.get("_sprack")
        dj = (l.get("_typer") or {}).get(SCHAKTTYP_DJUP)
        rad = [l.get("id"), l.get("system"), l.get("fran") or "(fri ände)", l.get("till") or "(fri ände)",
               _r(l.get("langd_m"), 3), l.get("dimension"),
               (f"{l['dimension_antagen']['riktning']} {l['dimension_antagen']['fran']}"
                if l.get("dimension_antagen") else None),
               l.get("material"), l.get("ledningstyp"),
               l.get("anlaggningsar"), l.get("renoveringsar"), l.get("vg_fran"), l.get("vg_till"),
               l.get("lutning_promille"), l.get("lutning_gis"), l.get("djup_medel"), l.get("djup_min"),
               l.get("djup_max"), l.get("djupklass") or None, l.get("djup_kalla") or None, l.get("tackning_min"),
               l.get("djup_start"), l.get("djup_slut"), l.get("antal_serviser"), _r(l.get("serviser_m")),
               _anordningar(l, "ventil"), _anordningar(l, "brandpost"), _anordningar(l, "servisventil"),
               _anordningar(l, "ledningsande"), _anordningar(l, "anordning"),
               l.get("schaktgrupp"), l.get("_grav"), _r((o or {}).get("delad_m")) if o else None, _grav_text(l),
               _r(schaktdel.get(id(l)), 2), ", ".join(l.get("schakt_ledningar") or []) or None,
               ", ".join(f"{p['ledning']} ({p['system']}, {round(100 * (p.get('andel') or 0))} %)"
                         for p in l.get("parallella") or []) or None,
               ", ".join(f"{k['ledning']} vid {k['station_m']:.1f} m".replace(".", ",")
                         for k in l.get("korsningar") or []) or None,
               _kr(o), SCHAKTSATT_NAMN.get((o or {}).get("schaktsatt")), l.get("_alt_metod"), _kr(l.get("_alt")),
               _anmarkning(l), trafikpaslag(l), provisorier(l)] + \
              [_kr((l.get("_typer") or {}).get(t)) for t in ALLA_SCHAKTTYPER] + \
              [SCHAKTSATT_NAMN.get((dj or {}).get("schaktsatt")), _r((dj or {}).get("spont_m2")) or None,
               _kr(dj, "spont") or None] + \
              [_kr(o, "schakt"), _r((o or {}).get("gravdjup"), 2), _r((o or {}).get("spont_m2")) or None,
               _kr(o, "spont") or None, _kr(o, "djuptillagg"), _kr(o, "aterstallning"), _kr(o, "brunnar"),
               _kr(o, "serviser"),
               _kr(s, "per_meter"), _kr(s, "infodring"), _kr(s, "serviser"),
               hattar.get(l.get("id")) or None, _kr(s, "framschaktning"), _kr(s),
               _kr(r, "per_meter"), _kr(r, "sprackning"), _kr(r, "gropar"), _kr(r, "servisgropar"),
               _kr(r, "ventilgropar"), _kr(r, "serviser"), _kr(r), l.get("_alt_grav"),
               ", ".join(l.get("flaggor") or []) or None, l.get("omrade") or None,
               l.get("lager"), l.get("oid"), l.get("del"),
               l.get("klippt_m") or (l.get("attribut") or {}).get("KLIPPT_M") or None]
        rad += [(l.get("attribut") or {}).get(a) for a in attrnamn]
        _radformler(rad, kol, K, nr, l)
        rader.append(rad)
    format_ = {"Längd": M, "VG från": "0.000", "VG till": "0.000", "Lutning": P, "Lutning GIS": P, "Djup medel": M2,
               "Djup min": M2, "Djup max": M2, "Täckning min": M2, "Djup start": M2, "Djup slut": M2,
               "Servislängd": M}
    for k in kol:
        if _dela_rubrik(k)[1] in ("kr", "kr/m"):
            format_[_dela_rubrik(k)[0]] = KR
    format_["Schaktlängd"] = M
    format_["Delad grav"] = M
    format_["Trafiktillägg"] = format_["Provisorier"] = "0%"
    format_["Slagen spont efter djup"] = M
    tabell(ws, kol + attrkol, rader, bredder={"Anmärkning": 40, "Flaggor": 24, "ID": 14, "Grav längs ledningen": 40},
           dolda=DOLDA_LEDNINGAR + attrkol, format=format_)

    # --- Kostnad per schaktgrupp (formler mot fliken Ledningar) ----------------------------------
    ws = wb.create_sheet("Kostnad per schaktgrupp")
    kol = ["Schaktgrupp", "System", "Ledningar (st)", "Ledningar", "Längd (m)", "Schaktlängd (m)", "Alt. lägg om (kr)",
           "Alt. schaktfritt (kr)", "Strumpas (st)", "Rörspräcks (st)", "Läggs om (st)", "Utan kostnad (st)"]
    rader = []

    def grupprad(namn_, lista, krit, ids):
        f = _gruppformler(K, krit)
        rader.append([namn_, "+".join(sorted({l.get("system") or "?" for l in lista})), f[0], ids] + f[1:])
    for g in sorted(k for k in grupper if k is not None):
        grupprad(g, grupper[g], [(K("Schaktgrupp"), _txt(g))], ", ".join(str(l.get("id")) for l in grupper[g]))
    if grupper.get(None):
        grupprad("utan samschakt", grupper[None], [(K("Schaktgrupp"), _txt(None))], None)
    if ledningar:
        grupprad("totalt", ledningar, [], None)
    start = tabell(ws, kol, rader, bredder={"Ledningar": 40},
                   format={"Längd": M, "Schaktlängd": M, "Alt. lägg om": KR, "Alt. schaktfritt": KR})
    ws.cell(start + len(rader) + 1, 1, "Schaktlängd: schaktets längd när alla ledningar i gruppen läggs om – delar "
            "som går i samschakt räknas en gång.").font = Font(italic=True, color="595959")
    for ri in range(start, start + len(rader)):
        if isinstance(ws.cell(ri, 1).value, str):
            for c in ws[ri]:
                c.font = Font(bold=True)

    # --- Serviser, Parallella, Korsningar, Brunnar, Datakvalitet ----------------------------------
    ws = wb.create_sheet("Serviser")
    rader = [[s.get("id"), s.get("ledning") or "(ingen huvudledning)", s.get("system"), s.get("position_m"),
              _r(s.get("langd_m")), s.get("_servisgrav"), s.get("dimension"), s.get("material") or None,
              s.get("ledningstyp") or None, s.get("vg_vid_anslutning"), s.get("djup_vid_anslutning"), s.get("lager"),
              s.get("oid")]
             for s in data.get("serviser", [])]
    tabell(ws, ["ID", "Huvudledning", "System", "Läge (m)", "Längd (m)", "Serviser i samma grav (st)", "Dimension (mm)",
                "Material", "Ledningstyp", "VG vid anslutning (m)", "Djup vid anslutning (m)", "Lager", "OID"], rader,
           format={"Läge": M, "Längd": M, "Djup vid anslutning": M2})

    if data.get("anordningar"):
        ws = wb.create_sheet("Anordningar")
        rader = [[a.get("id"), ANORDNING_NAMN.get(a.get("kategori"), a.get("kategori")), a.get("typ") or None,
                  a.get("ledning") or "(ingen ledning)", a.get("system"), a.get("position_m"), a.get("avstand_m"),
                  a.get("via_servis"), a.get("dimension"), a.get("anlaggningsar"), a.get("x"), a.get("y"),
                  a.get("lager"), a.get("oid")]
                 for a in sorted(data["anordningar"], key=lambda a: (str(a.get("ledning")), a.get("position_m") or 0))]
        tabell(ws, ["ID", "Kategori", "Typ", "Ledning", "System", "Läge (m)", "Avstånd till ledning (m)", "Via servis",
                    "Dimension (mm)", "Anläggningsår", "X", "Y", "Lager", "OID"], rader,
               dolda=["X", "Y", "OID"], format={"Läge": M, "Avstånd till ledning": M2})

    ws = wb.create_sheet("Parallella")
    rader = [[p.get("ledning_a"), p.get("ledning_b"), p.get("system_a"), p.get("system_b"),
              _r(100 * (p.get("andel_a") or 0), 0), _r(100 * (p.get("andel_b") or 0), 0), p.get("medelavstand_m"),
              p.get("gemensam_m"), "ja" if p.get("samschakt") else "nej", p.get("schaktgrupp")]
             for p in data.get("parallella", [])]
    tabell(ws, ["Ledning A", "Ledning B", "System A", "System B", "Andel av A (%)", "Andel av B (%)",
                "Medelavstånd c–c (m)", "Gemensam längd (m)", "Samschakt", "Schaktgrupp"], rader,
           format={"Medelavstånd c–c": M2, "Gemensam längd": M})

    ws = wb.create_sheet("Korsningar")
    rader = [[k.get("ledning_a"), k.get("ledning_b"), k.get("system_a"), k.get("system_b"), k.get("station_a"),
              k.get("station_b"), k.get("niva_a"), k.get("niva_b"), k.get("hojdskillnad_m"),
              ", ".join(x for x in (k.get("niva_kalla_a"), k.get("niva_kalla_b")) if x), k.get("x"), k.get("y")]
             for k in data.get("korsningar", [])]
    tabell(ws, ["Ledning A", "Ledning B", "System A", "System B", "Läge längs A (m)", "Läge längs B (m)",
                "Nivå A (m)", "Nivå B (m)", "Höjdskillnad A − B (m)", "Nivåkälla", "X", "Y"], rader,
           format={"Läge längs A": M, "Läge längs B": M, "Höjdskillnad A − B": M2})

    ws = wb.create_sheet("Brunnar")
    battr: list[str] = []
    for b in data.get("brunnar", []):
        for k in (b.get("attribut") or {}):
            if k not in battr:
                battr.append(k)
    rader = [[b.get("littera"), b.get("typ") or None, b.get("system") or None, b.get("lockniva"), b.get("bottenniva"),
              b.get("vg_min"), b.get("djup_m"), b.get("antal_ledningar"), b.get("omrade") or None, b.get("x"),
              b.get("y"), b.get("lager")] + [(b.get("attribut") or {}).get(a) for a in battr]
             for b in data.get("brunnar", [])]
    bkol = ["Littera", "Typ", "System", "Locknivå (m)", "Bottennivå (m)", "Lägsta VG (m)", "Djup (m)",
            "Ledningar (st)", "Område", "X", "Y", "Lager"]
    tabell(ws, bkol + [a if a not in bkol else a + " (GIS)" for a in battr], rader, dolda=battr + ["X", "Y"],
           format={"Djup": M2})

    ws = wb.create_sheet("Datakvalitet")
    rader = [[q.get("typ"), q.get("id"), q.get("text")] for q in
             sorted(data.get("datakvalitet", []), key=lambda q: (q.get("typ") or "", str(q.get("id"))))]
    for l in ledningar:
        if l.get("_omlagg_brist"):
            rader.append(["ingen omläggningskostnad" if not l.get("_omlagg") else "omläggning", l.get("id"),
                          l["_omlagg_brist"]])
        if l.get("dimension") is None:
            rader.append(["dimension saknas", l.get("id"), "ingen ledning uppströms eller nedströms med dimension"
                          if DIMENSION_UR_GRANNE else ""])
        elif l.get("dimension_antagen"):
            rader.append(["dimension antagen", l.get("id"), _dim_text(l)])
        if l.get("anlaggningsar") is None:
            rader.append(["anläggningsår saknas", l.get("id"), ""])
        if not l.get("fran") or not l.get("till"):
            if l.get("system") != "V":
                rader.append(["ledning utan brunn i någon ände", l.get("id"), "fri ände (servis, T-rör eller fel i GIS)"])
    dq_start = tabell(ws, ["Typ", "ID", "Beskrivning"], rader, bredder={"Beskrivning": 70, "Typ": 32})
    DQ = Kolumner("Datakvalitet", ["Typ", "ID", "Beskrivning"], dq_start, len(rader))
    dq_typer = sorted({r[0] for r in rader if r[0]})

    ws = wb.create_sheet("Kostnadsposter")
    rader = [[k["post"], k["schakttyp"], k["ledningar"], k["storre"], ", ".join(sorted(k["system"] or [])) or None,
              k["dim_fran"], k["dim_till"], k["enhet"], k["kr"], k["kommentar"]] for k in kost]
    tabell(ws, ["Post", "Schakttyp", "Ledningar i graven", "Större än 400 mm", "System", "Dimension från (mm)",
                "Dimension till (mm)", "Enhet", "Belopp (kr)", "Kommentar"], rader,
           bredder={"Kommentar": 70}, format={"Belopp": "#,##0.##"})

    skriv_sammanfattning(ws_s, data, ledningar, kost, kallor, K, DQ, dq_typer)
    skriv_investeringskalkyl(wb.create_sheet("Investeringskalkyl", 1), K)
    wb.calculation.fullCalcOnLoad = True           # summorna är formler – räkna när filen öppnas
    wb.save(path)


def skriv_investeringskalkyl(ws, K: Kolumner) -> None:
    """Investeringskalkyl enligt SVOA:s mall (Tabell A och B, inriktningsbeslut BP2) för båda alternativen.
    Entreprenaden = de jämförbara ledningarnas kostnad i fliken Ledningar (som Sammanfattning), med påslagen för
    material och SVOA-arbeten på rörspräckning och helschakt (inte infodring) och innerstads-/vintertillägg på allt.
    Byggherrekostnaderna är procent av entreprenaden per fas, kända risker och oförutsett procent av summan ovanför.
    Alla tal är formler; procentsatserna står i gula celler överst."""
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    KR, PROC = '#,##0 "kr"', "0%"
    tunn = Side(style="thin", color="BFBFBF")
    kant = Border(top=tunn, bottom=tunn, left=tunn, right=tunn)
    rubrikfarg = PatternFill("solid", fgColor="1F3864")
    gul = PatternFill("solid", fgColor="FFF2CC")
    ljus = PatternFill("solid", fgColor="DDEBF7")
    for kol, b in zip("ABCDEFG", (44, 22, 22, 22, 3, 22, 22)):
        ws.column_dimensions[kol].width = b

    def cell(rad, kol, v, fmt=None, fet=False, fyll=None, kursiv=False, ram=True):
        c = ws.cell(rad, kol, v)
        if fmt:
            c.number_format = fmt
        if fet or kursiv:
            c.font = Font(bold=fet, italic=kursiv, color="595959" if kursiv else None)
        if fyll:
            c.fill = fyll
        if ram:
            c.border = kant
        return c

    def rubrik(rad, texter):
        for i, t in enumerate(texter, 1):
            if t is None:
                continue
            c = cell(rad, i, t)
            c.font, c.fill = Font(bold=True, color="FFFFFF"), rubrikfarg
            c.alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[rad].height = 45

    ws.cell(1, 1, "Investeringskalkyl – inriktningsbeslut (BP2)").font = Font(bold=True, size=14)
    ws.cell(2, 1, "Enligt SVOA:s mall Kalkyl inriktningsbeslut LEDN (flikarna Alternativ, Entreprenadkalkyl och "
                  "Investeringsärende BP2). Procentsatserna i de gula cellerna kan ändras – allt räknas om.").font = \
        Font(italic=True, color="595959")

    # --- procentsatser --------------------------------------------------------------------------
    rad = 4
    rubrik(rad, ["Påslag i % av entreprenadkostnaden", "Planerafasen", "Genomförafasen"])
    andel = {}
    for namn, a_pl, a_ge in BYGGHERRE_ANDEL:
        rad += 1
        cell(rad, 1, namn)
        cell(rad, 2, a_pl, PROC, fyll=gul)
        cell(rad, 3, a_ge, PROC, fyll=gul)
        andel[namn] = (f"$B${rad}", f"$C${rad}")
    rad += 2
    rubrik(rad, ["Övriga påslag", "Andel", None])
    param = {}
    for nyckel, namn, v, kommentar in (
            ("risk", "Kända risker", KANDA_RISKER_ANDEL, "av byggherrekostnaderna (och entreprenaden i totalbudgeten)"),
            ("oforutsett", "Oförutsett", OFORUTSETT_ANDEL, "av summan ovanför i tabellen"),
            ("material", "Material som SVOA tillhandahåller (AFC.1522)", AFC_MATERIAL_ANDEL,
             "på rörspräckning och helschakt (inte infodring)"),
            ("arbeten", "Arbeten som SVOA utför (AFC.114)", AFC_SVOA_ARBETEN_ANDEL,
             "på rörspräckning och helschakt (inte infodring)"),
            ("vinter", "Vintertillägg", VINTERTILLAGG, "på hela entreprenaden: 5 % om arbetet kan ske vintertid")):
        rad += 1
        cell(rad, 1, namn)
        cell(rad, 2, v, PROC, fyll=gul)
        cell(rad, 3, kommentar, kursiv=True, ram=False)
        param[nyckel] = f"$B${rad}"

    # --- sammanställning ------------------------------------------------------------------------
    rad += 2
    rubrik(rad, ["Total projektbudget utan prisindexrisk", "Alt. lägg om", "Alt. schaktfritt"])
    sammanst = rad + 1
    cell(sammanst, 1, "Avrundat till närmaste 100 000-tal", fet=True)
    cell(sammanst + 1, 1, "Varav entreprenad inklusive material")
    rad = sammanst + 3
    ws.cell(rad, 1, "Bara ledningar med kostnad i båda alternativen ingår (som i Sammanfattning) – övriga står under "
                    "Utan kostnad där.").font = Font(italic=True, color="595959")

    JFR = _jamforbar_krit(K)
    METOD = K("Metod alt. schaktfritt")
    TRAFIK, PROV = K("Trafiktillägg"), K("Provisorier")
    for kol_samman, titel, kostnad, delar in (
            (2, "Alt. lägg om", K("Alt. lägg om"), [("Helschakt (omläggning)", None, True)]),
            (3, "Alt. schaktfritt", K("Alt. schaktfritt"), [("Infodring", "strumpa", False),
                                                            ("Rörspräckning", "rörspräckning", True),
                                                            ("Helschakt (omläggning)", "omläggning", True)])):
        rad += 3
        ws.cell(rad, 1, titel).font = Font(bold=True, size=13)
        # Entreprenadkalkyl
        rad += 1
        rubrik(rad, ["Entreprenadkalkyl", "Entreprenadkostnad", None])
        rad_delar = []
        for namn, metod, med_afc in delar:
            rad += 1
            krit = JFR + ([(METOD, _txt(metod))] if metod else [])
            cell(rad, 1, namn)
            cell(rad, 2, "=" + _sumifs(kostnad, krit), KR)
            rad_delar.append((rad, med_afc))
        afc = [f"B{r}" for r, m in rad_delar if m]
        rad += 1
        cell(rad, 1, "Material och SVOA-arbeten (AFC.1522, AFC.114)")
        cell(rad, 2, f"=({'+'.join(afc)})*({param['material']}+{param['arbeten']})" if afc else 0, KR)
        r_afc = rad
        # trafiktillägg och provisorier per ledning (procent i kolumnerna i Ledningar, från områdets fält),
        # på ledningens kostnad inklusive AFC-påslagen
        for kol_pct, namn_rad in ((TRAFIK, "Trafiktillägg (per område)"), (PROV, "Provisorier (per område)")):
            termer = []
            for namn, metod, med_afc in delar:
                villkor = [f'--({K("Alt. lägg om")}<>"")', f'--({K("Alt. schaktfritt")}<>"")'] + \
                    ([_lika(METOD, metod)] if metod else [])
                faktor = f"*(1+{param['material']}+{param['arbeten']})" if med_afc else ""
                termer.append(f"SUMPRODUCT({','.join(villkor + [kostnad, kol_pct])}){faktor}")
            rad += 1
            cell(rad, 1, namn_rad)
            cell(rad, 2, "=" + "+".join(termer), KR)
        rad += 1
        bas = "+".join([f"B{r}" for r, _ in rad_delar] + [f"B{r_afc}"])
        cell(rad, 1, "Vintertillägg")
        cell(rad, 2, f"=({bas})*{param['vinter']}", KR)
        rad += 1
        cell(rad, 1, "Entreprenad inklusive material", fet=True)
        cell(rad, 2, f"=SUM(B{rad_delar[0][0]}:B{rad - 1})", KR, fet=True, fyll=ljus)
        E = f"$B${rad}"

        # Tabell A – planeringsbudget
        rad += 2
        ws.cell(rad, 1, "TABELL A – INRIKTNINGSBESLUT (BP2)").font = Font(bold=True)
        rad += 1
        rubrik(rad, ["Moment", "Beräknad planeringsbudget",
                     "Beräknad planeringsbudget avrundad till närmsta 100 000-tal", "Kommentar"])
        a0 = rad + 1
        for namn, _, _ in BYGGHERRE_ANDEL:
            rad += 1
            cell(rad, 1, namn)
            cell(rad, 2, f"={E}*{andel[namn][0]}", KR)
            cell(rad, 3, f"=ROUND(B{rad},-5)", KR)
            cell(rad, 4, None, fyll=gul)
        rad += 1
        cell(rad, 1, "Kända risker")
        cell(rad, 2, f"=SUM(B{a0}:B{rad - 1})*{param['risk']}", KR)
        cell(rad, 3, f"=ROUND(B{rad},-5)", KR)
        cell(rad, 4, None, fyll=gul)
        rad += 1
        cell(rad, 1, "Oförutsett")
        cell(rad, 2, f"=SUM(B{a0}:B{rad - 1})*{param['oforutsett']}", KR)
        cell(rad, 3, f"=ROUND(B{rad},-5)", KR)
        cell(rad, 4, None, fyll=gul)
        rad += 1
        cell(rad, 1, "Summa", fet=True)
        cell(rad, 2, f"=SUM(B{a0}:B{rad - 1})", KR, fet=True)
        cell(rad, 3, f"=SUM(C{a0}:C{rad - 1})", KR, fet=True)

        # Tabell B – indikativ totalbudget
        rad += 2
        rubrik(rad, ["Moment", "Indikativ totalbudget inklusive planeringsbudget",
                     "Indikativ totalbudget inklusive planeringsbudget avrundad till närmsta 100 000-tal", "Kommentar"])
        b0 = rad + 1
        for namn, _, _ in BYGGHERRE_ANDEL:
            rad += 1
            cell(rad, 1, namn)
            cell(rad, 2, f"={E}*({andel[namn][0]}+{andel[namn][1]})", KR)
            cell(rad, 3, f"=ROUND(B{rad},-5)", KR)
            cell(rad, 4, None, fyll=gul)
        rad += 1
        cell(rad, 1, "Entreprenad inklusive material")
        cell(rad, 2, f"={E}", KR)
        cell(rad, 3, f"=ROUND(B{rad},-5)", KR)
        cell(rad, 4, None, fyll=gul)
        rad += 1
        cell(rad, 1, "Kända risker")
        cell(rad, 2, f"=SUM(B{b0}:B{rad - 1})*{param['risk']}", KR)
        cell(rad, 3, f"=ROUND(B{rad},-5)", KR)
        cell(rad, 4, None, fyll=gul)
        rad += 1
        cell(rad, 1, "Oförutsett")
        cell(rad, 2, f"=SUM(B{b0}:B{rad - 1})*{param['oforutsett']}", KR)
        cell(rad, 3, f"=ROUND(B{rad},-5)", KR)
        cell(rad, 4, None, fyll=gul)
        rad += 1
        cell(rad, 1, "Oförutsett – avrundning av totalbelopp (vid behov)")
        cell(rad, 2, None)
        cell(rad, 3, None, KR, fyll=gul)
        cell(rad, 4, "Summeras till oförutsett.", kursiv=True)
        rad += 1
        cell(rad, 1, "Summa", fet=True)
        cell(rad, 2, f"=SUM(B{b0}:B{rad - 1})", KR, fet=True)
        cell(rad, 3, f"=SUM(C{b0}:C{rad - 1})", KR, fet=True)
        r_summa = rad
        rad += 2
        cell(rad, 1, "Total projektbudget utan prisindexrisk", fet=True)
        cell(rad, 3, f"=C{r_summa}", KR, fet=True, fyll=ljus)
        cell(sammanst, kol_samman, f"=C{rad}", KR, fet=True, fyll=ljus)
        cell(sammanst + 1, kol_samman, f"={E}", KR)
    ws.sheet_view.showGridLines = False


def skriv_sammanfattning(ws, data, ledningar, kost, kallor, K: Kolumner, DQ: Kolumner, dq_typer: list) -> None:
    """Sammanfattningen. Alla tal i blocken är formler mot flikarna Ledningar och Datakvalitet (SUMIFS/COUNTIFS),
    så att de följer med när något ändras där."""
    from openpyxl.styles import Font
    KR, M = "#,##0", "#,##0"
    rad = 1
    ws.cell(rad, 1, "Ledningsinventering").font = Font(bold=True, size=14)
    rad += 1
    p = data.get("parametrar") or {}
    info = [("Ledningsdata", kallor["ledningsdata"]), ("Genererad i ArcMap", data.get("genererad")),
            ("Koordinatsystem", data.get("koordinatsystem")), ("Höjdsystem", data.get("hojdsystem")),
            ("Parallellt", (f"inom {_dk(p.get('parallell_avstand_m'))} m c–c, minst "
                            f"{_dk(p.get('parallell_min_m', PARALLELL_MIN_M))} m i följd"
                            if p.get("parallell_andel") is None else
                            f"inom {_dk(p.get('parallell_avstand_m'))} m c–c (äldre ledningsdata: andelen av längden "
                            f"uppskattad, delar med minst {_dk(PARALLELL_MIN_M)} m gemensamt)") if p else None),
            ("Samma grav", "räknas delsträcka för delsträcka längs ledningen: där ledningar går parallellt delar de "
                           "graven (pris för 2–3 ledningar delat på antalet), där en ledning går ensam gäller priset "
                           "för en ledning – se kolumnen Grav längs ledningen"),
            ("Vatten, djup till centrum", f"{_dk(p.get('vatten_djup_m'))} m (djupet räknas till rörets botten: + dimension/2)" if p.get("vatten_djup_m") is not None else None),
            ("Markhöjd", p.get("markhojd") or "locknivåerna (ingen markhöjd vald)"),
            ("Schakttyp", f"{ALLA_SCHAKTTYPER.get(kallor['schakttyp'], kallor['schakttyp'])} (alla typer i tabellen nedan)"),
            ("Schakt efter djup", f"gravens medeldjup under {_dk(SPONTLADA_FRAN_M)} m ledningschakt, "
                                  f"{_dk(SPONTLADA_FRAN_M)}–{_dk(SLAGEN_SPONT_FRAN_M)} m spontlåda (spontsläde), från "
                                  f"{_dk(SLAGEN_SPONT_FRAN_M)} m ledningschakt + slagen spont på {SPONT_SIDOR} sidor, "
                                  f"höjd medeldjup + {_dk(SPONT_NEDSLAGNING_M)} m"
             if kallor["schakttyp"] == SCHAKTTYP_DJUP else None),
            ("Infodring", f"spill-, dag- och kombinerade ledningar från {STRUMPA_DIM_MIN} mm, så långt prislistan räcker"
                          " (inte vatten)"),
            ("Rörspräckning", f"alla ledningar till och med {SPRACKNING_DIM_MAX} mm; införings- och maskingrop per "
                              f"påbörjade {GROP_PER_M} m, servisgrop per servis, ventilgrop per "
                              f"{' och '.join(VENTILGROP_KATEGORIER)} på ledningen"),
            ("Djuptillägg", f"per påbörjade {_dk(DJUPTILLAGG_STEG_M)} m djup under {_dk(STANDARDDJUP_M)} m"),
            ("Saknad dimension", (f"{data.get('_antagna_dim', 0)} ledningar har fått uppströms ledningens dimension"
                                  + (" (nedströms när ingen finns uppströms)" if DIMENSION_NEDSTROMS_RESERV else "")
                                  + " – se Anmärkning") if DIMENSION_UR_GRANNE else None),
            ("Serviser", "förnyas i båda alternativen (helschakt eller infodring per meter servisledning)"),
            ("Hattar", f"spill-, dag- och kombinerade serviser högst {_dk(SERVIS_HATT_MAX_M)} m infodras med hatt "
                       f"({_dk(kr(kost, 'servis_hatt', system='S'))} kr/st) när huvudledningen infodras eller rörspräcks; "
                       "längre per meter"
             if kr(kost, "servis_hatt", system="S") is not None else None),
            ("Kostnadsfil", os.path.basename(kallor["kostnader"])),
            ("Analys", datetime.now().strftime("%Y-%m-%d %H:%M"))]
    antagna = sorted({k["post"] for k in kost if re.search(r"schablon|antaget", (k.get("kommentar") or "").lower())})
    if antagna:
        info.append(("OBS", "antagna belopp/schabloner i kostnadsfilen: " + ", ".join(antagna)))
    for k, v in info:
        if v is None:
            continue
        ws.cell(rad + 1, 1, k).font = Font(bold=True)
        ws.cell(rad + 1, 2, v)
        rad += 1
    rad += 2

    def block(titel, kolumner, rader, format=None):
        nonlocal rad
        ws.cell(rad, 1, titel).font = Font(bold=True, size=12)
        start = tabell(ws, kolumner, rader, format=format, rad_start=rad + 1)
        rad = start + len(rader) + 2

    system = sorted({l.get("system") or "" for l in ledningar}, key=lambda s: "SDKV".find(s) % 9)
    L, ID, SYS = K("Längd"), K("ID"), K("System")
    JFR = _jamforbar_krit(K)

    def f(x):
        return "=" + x

    def per_system(formler):
        """En rad per system + Alla; formler(villkor) får [(System-området, systemet)] eller [] för Alla."""
        ut = []
        for s_ in system + [None]:
            krit = [(SYS, _txt(s_))] if s_ is not None else []
            ut.append([SYSTEM_NAMN.get(s_, s_) if s_ is not None else "Alla"] + formler(krit))
        return ut

    def utan_kostnad(krit, villkor):
        return f(f"{_countifs(ID, krit)}-{_countifs(ID, krit + villkor)}")

    block("Lägga om eller schaktfritt", ["System", "Längd (m)", "Alt. lägg om (kr)", "Alt. schaktfritt (kr)",
                                     "Strumpas (m)", "Rörspräcks (m)", "Läggs om (m)", "Utan kostnad (st)"],
          per_system(lambda krit: [f(_sumifs(L, krit)), f(_sumifs(K("Alt. lägg om"), krit + JFR)),
                                   f(_sumifs(K("Alt. schaktfritt"), krit + JFR))] +
                     [f(_sumifs(L, krit + [(K("Metod alt. schaktfritt"), _txt(m))]))
                      for m in ("strumpa", "rörspräckning", "omläggning")] +
                     [utan_kostnad(krit, JFR)]),
          format={"Längd": M, "Alt. lägg om": KR, "Alt. schaktfritt": KR, "Strumpas": M, "Rörspräcks": M,
                  "Läggs om": M})
    for t in ("Alt. lägg om: alla ledningar läggs om; ledningar i samma grav delar graven.",
              "Alt. schaktfritt: spill/dag infodras (rörspräcks eller läggs om där infodring inte går); vatten rörspräcks, "
              "men läggs om i samma grav när en avloppsledning i graven läggs om.",
              "Summorna omfattar bara ledningar med kostnad i båda alternativen; övriga står under Utan kostnad "
              "(djup eller dimension saknas – se fliken Datakvalitet).",
              "Talen är formler mot fliken Ledningar – ändras en kostnad där räknas sammanfattningen om."):
        ws.cell(rad - 1, 1, t).font = Font(italic=True, color="595959")
        rad += 1
    rad += 1

    typkol = [_dela_rubrik(k)[0] for k in ("Lägg om efter djup (kr)", "Lägg om ledningschakt (kr)",
                                           "Lägg om med spontsläde (kr)", "Lägg om bergschakt (kr)",
                                           "Lägg om KC-förstärkning (kr)")]
    alla_typer = [(K(t), EJ_TOM) for t in typkol]
    block("Alt. lägg om per schakttyp", ["System", "Längd (m)"] + [f"{v} (kr)" for v in ALLA_SCHAKTTYPER.values()] +
          ["Utan kostnad (st)"],
          per_system(lambda krit: [f(_sumifs(L, krit))] + [f(_sumifs(K(t), krit + alla_typer)) for t in typkol] +
                     [utan_kostnad(krit, alla_typer)]),
          format=dict([("Längd", M)] + [(v, KR) for v in ALLA_SCHAKTTYPER.values()]))

    def satt_rader():
        rader_ = []
        for k, namn in SCHAKTSATT_NAMN.items():
            if not any((l.get("_typer") or {}).get(SCHAKTTYP_DJUP)
                       and l["_typer"][SCHAKTTYP_DJUP].get("schaktsatt") == k for l in ledningar):
                continue
            krit = [(K("Schaktsätt efter djup"), _txt(namn))]
            rader_.append([namn, f(_countifs(ID, krit)), f(_sumifs(L, krit)),
                           f(_sumifs(K("Slagen spont efter djup"), krit)), f(_sumifs(K("Spont efter djup"), krit)),
                           f(_sumifs(K("Lägg om efter djup"), krit))])
        return rader_
    block("Alt. lägg om efter djup – schaktsätt", ["Schaktsätt", "Ledningar (st)", "Längd (m)", "Slagen spont (m²)",
                                                   "Spont (kr)", "Alt. lägg om efter djup (kr)"], satt_rader(),
          format={"Längd": M, "Slagen spont": M, "Spont": KR, "Alt. lägg om efter djup": KR})

    def medeldjup(krit):
        D = K("Djup medel")
        villkor = [f"--({SYS}={krit[0][1] if krit[0][1] != _txt(None) else chr(34) * 2})"] if krit else []
        taljare = "SUMPRODUCT(" + ",".join(villkor + [D, L]) + ")"
        return f(f'IFERROR({taljare}/{_sumifs(L, krit + [(D, EJ_TOM)])},"")')

    block("Ledningsnätet", ["System", "Ledningar (st)", "Längd (m)", "Serviser (st)", "Servislängd (m)",
                            "Medeldjup (m)", "I samschakt (%)"],
          per_system(lambda krit: [f(_countifs(ID, krit)), f(_sumifs(L, krit)), f(_sumifs(K("Serviser"), krit)),
                                   f(_sumifs(K("Servislängd"), krit)), medeldjup(krit),
                                   f(f'IFERROR(100*{_sumifs(L, krit + [(K("Schaktgrupp"), EJ_TOM)])}'
                                     f'/{_sumifs(L, krit)},"")')]),
          format={"Längd": M, "Servislängd": M, "Medeldjup": "0.00", "Serviser": M, "I samschakt": "0"})

    if data.get("anordningar"):
        kat = [k for k in ANORDNING_NAMN if any(a.get("kategori") == k for a in data["anordningar"])]
        block("Anordningar (se fliken Anordningar)", ["System"] + [f"{ANORDNING_NAMN[k]} (st)" for k in kat],
              per_system(lambda krit: [f(_sumifs(K(ANORDNING_NAMN[k]), krit)) for k in kat]))

    def fordelning(titel, nyckel, rubrik, villkor, sortering=None):
        """Längd per system och nyckel; villkor(nyckel) ger SUMPRODUCT-villkoren på fliken Ledningar (exakt
        jämförelse, så att text som "> 4" inte tolkas som ett villkor)."""
        nycklar = {(l.get("system") or "", nyckel(l)) for l in ledningar}
        rader = [[SYSTEM_NAMN.get(s_, s_), k if k not in (None, "") else "(saknas)",
                  f("SUMPRODUCT(" + ",".join([_lika(SYS, s_)] + villkor(k) + [L]) + ")")]
                 for (s_, k) in sorted(nycklar, key=lambda x: ("SDKV".find(x[0]) % 9,
                                                              sortering(x[1]) if sortering else str(x[1])))]
        block(titel, ["System", rubrik, "Längd (m)"], rader, format={"Längd": M})

    def lika(kolumn):
        return lambda k: [_lika(K(kolumn), k)]

    def decennium(k):
        AR = K("Anläggningsår")
        if k is None:
            return [f'--({AR}="")']
        d = int(k[:4])
        return [f'--({AR}<>"")', f"--({AR}>={d})", f"--({AR}<{d + 10})"]

    fordelning("Material", lambda l: l.get("material"), "Material", lika("Material"))
    fordelning("Dimension", lambda l: l.get("dimension"), "Dimension (mm)", lika("Dimension"),
               lambda v: (v is None, v if isinstance(v, (int, float)) else 0))
    fordelning("Anläggningsdecennium", lambda l: f"{l['anlaggningsar'] // 10 * 10}-tal" if l.get("anlaggningsar") else None,
               "Decennium", decennium, lambda v: (v is None, str(v)))
    fordelning("Djupklass", lambda l: l.get("djupklass") or None, "Djupklass (m)", lika("Djupklass"),
               lambda v: (v is None, str(v).replace(">", "9")))

    rader = [[t, f(f"COUNTIF({DQ('Typ')},{_txt(t)})")] for t in dq_typer]
    if rader:
        block("Datakvalitet (se fliken Datakvalitet)", ["Typ", "Antal"], rader)
    ws.column_dimensions["A"].width = max(ws.column_dimensions["A"].width or 0, 28)
    ws.column_dimensions["B"].width = max(ws.column_dimensions["B"].width or 0, 22)


# ============================================================================
# JSON tillbaka till kartan
# ============================================================================

def skriv_resultat(ledningar: list[dict], path: str, kalla: str) -> None:
    poster = []
    for l in ledningar:
        poster.append({
            "id": l.get("id"), "lager": l.get("lager"), "oid": l.get("oid"), "del": l.get("del"),
            "fran": l.get("fran"), "till": l.get("till"), "system": l.get("system"),
            "dimension": l.get("dimension"), "dimension_gis": l.get("dimension_gis"),
            "dimension_antagen": l.get("dimension_antagen"),
            "langd_m": l.get("langd_m"), "djup_medel": l.get("djup_medel"), "djupklass": l.get("djupklass"),
            "antal_serviser": l.get("antal_serviser"), "serviser_m": l.get("serviser_m"),
            "schaktgrupp": l.get("schaktgrupp"), "ledningar_i_grav": l.get("_grav"),
            "delad_grav_m": _r((l.get("_omlagg") or {}).get("delad_m")) if l.get("_omlagg") else None,
            "grav_delar": [{k: d.get(k) for k in ("fran_m", "till_m", "langd_m", "ledningar") if k in d}
                           for d in grav_delar(l)],
            "kostnad_omlaggning": _kr(l.get("_omlagg")), "kostnad_strumpa": _kr(l.get("_strumpa")),
            "kostnad_rorsprackning": _kr(l.get("_sprack")),
            "anordningar_antal": l.get("anordningar_antal"),
            "alt_lagg_om_kr": _kr(l.get("_omlagg")),
            "alt_lagg_om_per_schakttyp": {t: _kr(v) for t, v in (l.get("_typer") or {}).items()},
            "alt_schaktfritt_metod": l.get("_alt_metod"),
            "alt_schaktfritt_kr": _kr(l.get("_alt")), "anmarkning": _anmarkning(l),
            "omrade": l.get("omrade") or None, "trafikpaslag": trafikpaslag(l), "provisorier": provisorier(l),
        })
    with open(path, "w", encoding="utf-8") as f:
        json.dump({"version": 2, "kalla": "ledning_analys.py", "ledningsdata": kalla,
                   "genererad": datetime.now().strftime("%Y-%m-%d %H:%M"), "ledningar": poster},
                  f, ensure_ascii=False, indent=1)


# ============================================================================
# main
# ============================================================================

def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="Inventeringsbok och kostnad per metod (lägg om, strumpa, rörspräckning)"
                                             " ur ledningsdata.json")
    ap.add_argument("ledningsdata", help="ledningsdata.json från ArcMap-verktyget Ledningsinventering")
    ap.add_argument("-o", "--utdata", help="utdatamapp (standard: samma mapp som ledningsdata.json)")
    ap.add_argument("--kostnader", default=KOSTNADSFIL, help="kostnadsfil (kostnader_inventering.csv)")
    ap.add_argument("--schakttyp", default=SCHAKTTYP, choices=sorted(ALLA_SCHAKTTYPER),
                    help="schakttyp i alternativen (alla redovisas i Sammanfattning)")
    a = ap.parse_args(argv)

    with open(a.ledningsdata, encoding="utf-8") as f:
        data = json.load(f)
    if data.get("kalla") != "inventering.py":
        print("OBS: filen kommer inte från Ledningsinventering – djup, serviser och parallella ledningar saknas"
              " (kör verktyget Ledningsinventering i ArcMap).")
    if not os.path.isfile(a.kostnader):
        print(f"Kostnadsfilen finns inte: {a.kostnader}")
        return 2
    kost = las_kostnader(a.kostnader)

    ledningar = berakna(data, kost, a.schakttyp)
    ut = a.utdata or os.path.dirname(os.path.abspath(a.ledningsdata))
    os.makedirs(ut, exist_ok=True)
    xlsx = os.path.join(ut, UTDATA_XLSX)
    kallor = {"ledningsdata": os.path.abspath(a.ledningsdata), "kostnader": a.kostnader, "schakttyp": a.schakttyp}
    try:
        skriv_excel(data, ledningar, xlsx, kost, kallor)
    except PermissionError:
        print(f"Kan inte skriva {xlsx} – är filen öppen i Excel? Stäng den och kör igen.")
        return 1
    skriv_resultat(ledningar, os.path.join(ut, UTDATA_JSON), os.path.abspath(a.ledningsdata))

    L = sum(l.get("langd_m") or 0 for l in ledningar)
    print(f"{len(ledningar)} ledningar, {L:,.0f} m – schakttyp {ALLA_SCHAKTTYPER[a.schakttyp]}".replace(",", " "))
    if data.get("_antagna_dim"):
        kvar = sum(1 for l in ledningar if l.get("dimension") is None)
        print(f"  {data['_antagna_dim']} ledningar utan dimension fick grannledningens dimension"
              + (f" ({kvar} saknar fortfarande)" if kvar else ""))
    jfr = [l for l in ledningar if _jamforbar(l)]
    for namn, k in (("Alt. lägg om", "_omlagg"), ("Alt. schaktfritt", "_alt")):
        belopp = sum(l[k]["summa"] for l in jfr)
        print(f"  {namn}: {belopp:,.0f} kr".replace(",", " "))
    if len(jfr) < len(ledningar):
        print(f"  ({len(ledningar) - len(jfr)} ledningar saknar kostnad i något alternativ och ingår inte – se Datakvalitet)")
    print(f"Skrev {xlsx}")
    print(f"Skrev {os.path.join(ut, UTDATA_JSON)}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
