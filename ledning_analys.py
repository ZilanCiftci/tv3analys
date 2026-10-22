#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ledning_analys.py – inventeringsbok och kostnad lägga om mot strumpa (PLAN_ledningsinventering.md steg 3)

Läser ledningsdata.json från ArcMap-verktyget Ledningsinventering och räknar två alternativ för hela nätet:

  Alt. lägg om  – alla ledningar läggs om i helschakt; ledningar i samma grav delar graven (pris per meter
                  grav efter antal ledningar, antal större än 400 mm och schakttyp). Räknas för alla fyra
                  schakttyperna (ledningschakt, med spontsläde, bergschakt, KC-förstärkning).
  Alt. strumpa  – spill-, dag- och kombinerade ledningar infodras (rörspräcks om infodring inte går och
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
SCHAKTTYP = "ledningschakt"       # schakttypen i Alt. lägg om och Alt. strumpa (övriga redovisas också)
STORRE_DIM = 400                  # ledningar över detta (mm) räknas som "större" i grav med 2–3 ledningar
MAX_LEDNINGAR_I_GRAV = 3          # prislistan har 1–3 ledningar i samma grav
STANDARDDJUP_M = 2.5              # schaktpriset gäller till detta djup (vattengång) ...
DJUPTILLAGG_STEG_M = 1.5          # ... därefter ett djuptillägg per påbörjade 1,5 m
STRUMPA_SYSTEM = ("S", "D", "K")  # system som kan infodras – vattenledningar (V) infodras inte
STRUMPA_DIM_MIN = 100             # infodring från denna dimension (mm); övre gräns = sista priset i filen
SPRACKNING_DIM_MAX = 400          # rörspräckning för dimensioner till och med detta (mm), alla system
GROP_PER_M = 50                   # rörspräckning: en införingsgrop och en maskingrop per påbörjade 50 m
SERVIS_GRAV_AVSTAND_M = 3.0       # serviser vars båda ändar ligger så här nära varandra delar servisgrav
SERVIS_LANGD_MAX_M = None         # meter servisledning som förnyas per servis; None = hela servisens längd
FRAMSCHAKTA_TB = True             # infodring: båda ändbrunnarna tillsyns-/rensbrunn -> framschaktning + ny brunn
BRUNNSTYP_TILLSYN = r"TB|RB"      # brunnstyp (t.ex. WellFunction STB/SRB) som räknas som tillsyns-/rensbrunn
VENTILGROP_KATEGORIER = ("ventil", "brandpost")   # rörspräckning: en ventilgrop per sådan anordning på ledningen
                                                  # (servisventiler ligger i servisgropen)
ANORDNING_NAMN = {"ventil": "Ventiler", "brandpost": "Brandposter", "servisventil": "Servisventiler",
                  "ledningsande": "Ledningsändar", "anordning": "Övriga anordningar"}

# Kolumner i fliken Ledningar som döljs (grupperade, fälls ut med plustecknet)
DOLDA_LEDNINGAR = [
    "Ledningstyp", "Renoveringsår", "VG från", "VG till", "Lutning GIS", "Djup min", "Djup max", "Djupkälla",
    "Täckning min", "Djup start", "Djup slut", "Samschakt med", "Parallella", "Korsningar",
    "Lägg om ledningschakt", "Lägg om med spontsläde", "Lägg om bergschakt", "Lägg om KC-förstärkning",
    "Schakt", "Djuptillägg", "Återställning", "Brunnar", "Serviser helschakt",
    "Infodring", "Infodring per meter", "Serviser infodring", "Framschaktning", "Infodring totalt",
    "Servisventiler", "Övriga anordningar",
    "Rörspräckning", "Rörspräckning per meter", "Gropar", "Servisgropar", "Ventilgropar", "Serviser rörspräckning",
    "Rörspräckning totalt", "Ledningar i graven alt. strumpa", "Område", "Lager", "OID", "Del",
    "Klippt vid områdesgränsen",
]

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


def _gravpartners(led: dict, per_id: dict, medlagda: set | None = None) -> list[tuple[dict, float]]:
    """Samschakt-partner (ledning, andel av den här ledningens längd) som också läggs om, störst andel först.
    medlagda = id:n som läggs om i alternativet (None = alla)."""
    ut = []
    for p in led.get("parallella") or []:
        q = per_id.get(p.get("ledning"))
        if p.get("samschakt") and q is not None and (medlagda is None or q.get("id") in medlagda):
            ut.append((q, p.get("andel") or 0.0))
    return sorted(ut, key=lambda x: -x[1])


def ledningar_i_grav(led: dict, per_id: dict | None = None, medlagda: set | None = None) -> int:
    """1 + summan av andelen av ledningens längd som delas med varje samschakt-partner som också läggs
    om, avrundat (högst MAX_LEDNINGAR_I_GRAV)."""
    if per_id is None:
        andel = sum(p.get("andel") or 0.0 for p in led.get("parallella") or [] if p.get("samschakt"))
    else:
        andel = sum(a for _, a in _gravpartners(led, per_id, medlagda))
    return max(1, min(MAX_LEDNINGAR_I_GRAV, 1 + int(round(andel))))


def _antal_djuptillagg(djup: float) -> int:
    if djup is None or djup <= STANDARDDJUP_M:
        return 0
    return int(math.ceil((djup - STANDARDDJUP_M) / DJUPTILLAGG_STEG_M - 1e-9))


def _grav(led, kost, per_id, schakttyp, medlagda=None):
    """Huvudledningens del av graven: ({schakt, djuptillagg, aterstallning, brunnar}, brist, n, beskrivning)."""
    L = led.get("langd_m") or 0.0
    system = led.get("system") or ""
    n = ledningar_i_grav(led, per_id, medlagda)
    i_grav = [led] + [q for q, _ in _gravpartners(led, per_id, medlagda)][:n - 1]
    dims = [q.get("dimension") for q in i_grav]
    if any(d is None for d in dims):
        return None, "dimension saknas" + ("" if led.get("dimension") is None else " på ledning i samma grav"), n, ""
    dmax = max(dims)
    storre = sum(1 for d in dims if d > STORRE_DIM)
    pris_m = gravpris(kost, schakttyp, n, storre, dmax)
    beskr = f"{n} ledning{'ar' if n > 1 else ''}" + (f", {storre} större" if n > 1 else "") + f", Ø {dmax:g}"
    if pris_m is None:
        return None, f"inget schaktpris ({beskr}, {SCHAKTTYPER.get(schakttyp, schakttyp)})", n, beskr
    o = {"schakt": pris_m / n * L}
    # Djuptillägg per påbörjade DJUPTILLAGG_STEG_M under STANDARDDJUP_M, medel över profilens punkter
    tillagg_m = kr(kost, "djuptillagg", dim=dmax) or 0.0
    djupen = _djup_punkter(led) if system != "V" else []
    if not djupen and led.get("djup_medel") is not None:
        djupen = [led["djup_medel"]]
    steg = (sum(_antal_djuptillagg(d) for d in djupen) / len(djupen)) if djupen else 0.0
    o["djuptillagg"] = steg * tillagg_m / n * L
    o["aterstallning"] = (kr(kost, "aterstallning") or 0.0) / n * L
    andar = sum(0.5 for b in (led.get("fran"), led.get("till")) if b) if system != "V" else 0.0
    o["brunnar"] = andar * (kr(kost, "brunn_ny") or 0.0)
    brist = "" if djupen or system == "V" else "djup saknas – inget djuptillägg"
    return o, brist, n, beskr


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


def _servis_helschakt(egna, alla, grupp, metod_for, kost, schakttyp, metod="omläggning"):
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
    return summa


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
         "serviser": sum((kr(kost, "servis_infodring", system=x.get("system")) or 0.0) * _servis_m(x) for x in serv)}
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
         "serviser": sum((kr(kost, "servis_infodring", system=x.get("system")) or 0.0) * _servis_m(x) for x in serv)}
    r["ventilgropar"] = r["antal_ventilgropar"] * (kr(kost, "ventilgrop") or kr(kost, "servisgrop") or 0.0)
    r["summa"] = r["sprackning"] + r["gropar"] + r["servisgropar"] + r["ventilgropar"] + r["serviser"]
    return r, ""


def antal_ventilgropar(led: dict) -> int:
    antal = led.get("anordningar_antal") or {}
    return sum(antal.get(k, 0) for k in VENTILGROP_KATEGORIER)


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
    o["serviser"] = _servis_helschakt(serv_idx, alla_serv, grupp, metod_for, kost, schakttyp)
    o["summa"] = o["schakt"] + o["djuptillagg"] + o["aterstallning"] + o["brunnar"] + o["serviser"]
    return o, brist, n, beskr


def berakna(data: dict, kost: list[dict], schakttyp: str = SCHAKTTYP) -> list[dict]:
    """Lägger till kostnaderna på varje ledning (i data['ledningar'])."""
    ledningar = data.get("ledningar", [])
    for led in ledningar:
        if not led.get("system"):            # gisdata.json från Exportera GIS-data saknar system
            led["system"] = (led.get("ledningstyp") or "").strip()[:1].upper()
        if not led.get("id"):
            led["id"] = f"{(led.get('lager') or '').split(chr(92))[-1]}/{led.get('oid')}"
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
        for t in SCHAKTTYPER:
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
    """Metod och kostnad per ledning i alternativet strumpa (_alt_metod, _alt, _alt_grav).
    Avlopp: infodring, annars rörspräckning, annars omläggning. Vatten: rörspräckning bara när ingen
    avloppsledning i samma grav läggs om – annars omläggning tillsammans med den."""
    for led in ledningar:
        if led.get("system") == "V":
            continue
        led["_alt_metod"] = "strumpa" if led.get("_strumpa") else ("rörspräckning" if led.get("_sprack") else "omläggning")
    omlagda_avlopp = {l.get("id") for l in ledningar if l.get("system") != "V" and l["_alt_metod"] == "omläggning"}
    for led in ledningar:
        if led.get("system") != "V":
            continue
        grannar = [q for q, _ in _gravpartners(led, per_id) if q.get("system") != "V"]
        schakt = [q.get("id") for q in grannar if q.get("id") in omlagda_avlopp]
        if schakt:
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
            if fmt.get(ci) and isinstance(v, (int, float)):
                c.number_format = fmt[ci]
    if rad_start == 1:
        ws.freeze_panes = f"A{start}"
        ws.auto_filter.ref = f"A{start - 1}:{get_column_letter(len(kolumner))}{max(start + len(rader) - 1, start)}"
    for i, kol in enumerate(kolumner, 1):
        b = (bredder or {}).get(namn[i - 1])
        if b is None:
            langsta = max((len(str(rd[i - 1])) for rd in rader[:300] if i - 1 < len(rd) and rd[i - 1] is not None),
                          default=0)
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


def _anmarkning(l: dict) -> str | None:
    delar = [x for x in (l.get("_alt_motiv"), l.get("_anm"),
                         ("alt. strumpa: " + l["_alt_brist"]) if l.get("_alt_brist") else None) if x]
    return "; ".join(delar) or None


def _summa(lista, f):
    v = [f(l) for l in lista]
    v = [x for x in v if x is not None]
    return round(sum(v)) if v else None


def _jamforbar(l: dict) -> bool:
    """Ledningen har kostnad i båda alternativen – bara sådana summeras, så att summorna går att jämföra."""
    return bool(l.get("_omlagg")) and bool(l.get("_alt"))


def _altsummor(lista):
    """[alt. lägg om (kr), alt. strumpa (kr), strumpa (st), rörspräckning (st), omläggning (st), utan kostnad (st)]."""
    jfr = [l for l in lista if _jamforbar(l)]
    return [_summa(jfr, lambda l: l["_omlagg"]["summa"]),
            _summa(jfr, lambda l: l["_alt"]["summa"]),
            sum(1 for l in lista if l.get("_alt_metod") == "strumpa"),
            sum(1 for l in lista if l.get("_alt_metod") == "rörspräckning"),
            sum(1 for l in lista if l.get("_alt_metod") == "omläggning"),
            sum(1 for l in lista if not _jamforbar(l))]


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
    kol = ["ID", "System", "Från brunn", "Till brunn", "Längd (m)", "Dimension (mm)", "Material", "Ledningstyp",
           "Anläggningsår", "Renoveringsår", "VG från (m)", "VG till (m)", "Lutning (‰)", "Lutning GIS (‰)",
           "Djup medel (m)", "Djup min (m)", "Djup max (m)", "Djupklass (m)", "Djupkälla", "Täckning min (m)",
           "Djup start (m)", "Djup slut (m)", "Serviser (st)", "Servislängd (m)", "Ventiler (st)",
           "Brandposter (st)", "Servisventiler (st)", "Övriga anordningar (st)", "Schaktgrupp",
           "Ledningar i graven (st)", "Samschakt med", "Parallella", "Korsningar",
           "Alt. lägg om (kr)", "Metod alt. strumpa", "Alt. strumpa (kr)", "Anmärkning",
           "Lägg om ledningschakt (kr)", "Lägg om med spontsläde (kr)", "Lägg om bergschakt (kr)",
           "Lägg om KC-förstärkning (kr)",
           "Schakt (kr)", "Djuptillägg (kr)", "Återställning (kr)", "Brunnar (kr)", "Serviser helschakt (kr)",
           "Infodring per meter (kr/m)", "Infodring (kr)", "Serviser infodring (kr)", "Framschaktning (kr)",
           "Infodring totalt (kr)",
           "Rörspräckning per meter (kr/m)", "Rörspräckning (kr)", "Gropar (kr)", "Servisgropar (kr)",
           "Ventilgropar (kr)", "Serviser rörspräckning (kr)", "Rörspräckning totalt (kr)", "Ledningar i graven alt. strumpa (st)",
           "Flaggor", "Område", "Lager", "OID", "Del", "Klippt vid områdesgränsen (m)"]
    egna = {_dela_rubrik(k)[0] for k in kol}
    attrkol = [a if a not in egna else a + " (GIS)" for a in attrnamn]
    rader = []
    ordn = sorted(ledningar, key=lambda l: ("SDKV".find(l.get("system") or "?") % 9, str(l.get("id"))))
    for l in ordn:
        o, s, r = l.get("_omlagg"), l.get("_strumpa"), l.get("_sprack")
        rad = [l.get("id"), l.get("system"), l.get("fran") or "(fri ände)", l.get("till") or "(fri ände)",
               _r(l.get("langd_m")), l.get("dimension"), l.get("material"), l.get("ledningstyp"),
               l.get("anlaggningsar"), l.get("renoveringsar"), l.get("vg_fran"), l.get("vg_till"),
               l.get("lutning_promille"), l.get("lutning_gis"), l.get("djup_medel"), l.get("djup_min"),
               l.get("djup_max"), l.get("djupklass") or None, l.get("djup_kalla") or None, l.get("tackning_min"),
               l.get("djup_start"), l.get("djup_slut"), l.get("antal_serviser"), _r(l.get("serviser_m")),
               _anordningar(l, "ventil"), _anordningar(l, "brandpost"), _anordningar(l, "servisventil"),
               _anordningar(l, ("ledningsande", "anordning")),
               l.get("schaktgrupp"), l.get("_grav"), ", ".join(l.get("schakt_ledningar") or []) or None,
               ", ".join(f"{p['ledning']} ({p['system']}, {round(100 * (p.get('andel') or 0))} %)"
                         for p in l.get("parallella") or []) or None,
               ", ".join(f"{k['ledning']} vid {k['station_m']:.1f} m".replace(".", ",")
                         for k in l.get("korsningar") or []) or None,
               _kr(o), l.get("_alt_metod"), _kr(l.get("_alt")), _anmarkning(l)] + \
              [_kr((l.get("_typer") or {}).get(t)) for t in SCHAKTTYPER] + \
              [_kr(o, "schakt"), _kr(o, "djuptillagg"), _kr(o, "aterstallning"), _kr(o, "brunnar"),
               _kr(o, "serviser"),
               _kr(s, "per_meter"), _kr(s, "infodring"), _kr(s, "serviser"), _kr(s, "framschaktning"), _kr(s),
               _kr(r, "per_meter"), _kr(r, "sprackning"), _kr(r, "gropar"), _kr(r, "servisgropar"),
               _kr(r, "ventilgropar"), _kr(r, "serviser"), _kr(r), l.get("_alt_grav"),
               ", ".join(l.get("flaggor") or []) or None, l.get("omrade") or None,
               l.get("lager"), l.get("oid"), l.get("del"), l.get("klippt_m")]
        rad += [(l.get("attribut") or {}).get(a) for a in attrnamn]
        rader.append(rad)
    format_ = {"Längd": M, "VG från": "0.000", "VG till": "0.000", "Lutning": P, "Lutning GIS": P, "Djup medel": M2,
               "Djup min": M2, "Djup max": M2, "Täckning min": M2, "Djup start": M2, "Djup slut": M2,
               "Servislängd": M}
    for k in kol:
        if _dela_rubrik(k)[1] in ("kr", "kr/m"):
            format_[_dela_rubrik(k)[0]] = KR
    tabell(ws, kol + attrkol, rader, bredder={"Anmärkning": 40, "Flaggor": 24, "ID": 14},
           dolda=DOLDA_LEDNINGAR + attrkol, format=format_)

    # --- Kostnad per schaktgrupp ----------------------------------------------------------------
    grupper = defaultdict(list)
    for l in ledningar:
        grupper[l.get("schaktgrupp")].append(l)
    ws = wb.create_sheet("Kostnad per schaktgrupp")
    kol = ["Schaktgrupp", "System", "Ledningar (st)", "Ledningar", "Längd (m)", "Alt. lägg om (kr)",
           "Alt. strumpa (kr)", "Strumpas (st)", "Rörspräcks (st)", "Läggs om (st)", "Utan kostnad (st)"]
    rader = []
    for g in sorted(k for k in grupper if k is not None):
        lista = grupper[g]
        rader.append([g, "+".join(sorted({l.get("system") or "?" for l in lista})), len(lista),
                      ", ".join(str(l.get("id")) for l in lista), _r(sum(l.get("langd_m") or 0 for l in lista))]
                     + _altsummor(lista))
    for namn_, lista in (("utan samschakt", grupper.get(None, [])), ("totalt", ledningar)):
        if lista:
            rader.append([namn_, "+".join(sorted({l.get("system") or "?" for l in lista})), len(lista), None,
                          _r(sum(l.get("langd_m") or 0 for l in lista))] + _altsummor(lista))
    start = tabell(ws, kol, rader, bredder={"Ledningar": 40},
                   format={"Längd": M, "Alt. lägg om": KR, "Alt. strumpa": KR})
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
            rader.append(["dimension saknas", l.get("id"), ""])
        if l.get("anlaggningsar") is None:
            rader.append(["anläggningsår saknas", l.get("id"), ""])
        if not l.get("fran") or not l.get("till"):
            if l.get("system") != "V":
                rader.append(["ledning utan brunn i någon ände", l.get("id"), "fri ände (servis, T-rör eller fel i GIS)"])
    tabell(ws, ["Typ", "ID", "Beskrivning"], rader, bredder={"Beskrivning": 70, "Typ": 32})

    ws = wb.create_sheet("Kostnadsposter")
    rader = [[k["post"], k["schakttyp"], k["ledningar"], k["storre"], ", ".join(sorted(k["system"] or [])) or None,
              k["dim_fran"], k["dim_till"], k["enhet"], k["kr"], k["kommentar"]] for k in kost]
    tabell(ws, ["Post", "Schakttyp", "Ledningar i graven", "Större än 400 mm", "System", "Dimension från (mm)",
                "Dimension till (mm)", "Enhet", "Belopp (kr)", "Kommentar"], rader,
           bredder={"Kommentar": 70}, format={"Belopp": "#,##0.##"})

    skriv_sammanfattning(ws_s, data, ledningar, kost, kallor)
    wb.save(path)


def skriv_sammanfattning(ws, data, ledningar, kost, kallor) -> None:
    from openpyxl.styles import Font
    KR, M = "#,##0", "#,##0"
    rad = 1
    ws.cell(rad, 1, "Ledningsinventering").font = Font(bold=True, size=14)
    rad += 1
    p = data.get("parametrar") or {}
    info = [("Ledningsdata", kallor["ledningsdata"]), ("Genererad i ArcMap", data.get("genererad")),
            ("Koordinatsystem", data.get("koordinatsystem")), ("Höjdsystem", data.get("hojdsystem")),
            ("Parallellt (c–c, andel)", f"{_dk(p.get('parallell_avstand_m'))} m, {round(100 * (p.get('parallell_andel') or 0))} %"
             if p else None),
            ("Vatten, djup till centrum", f"{_dk(p.get('vatten_djup_m'))} m" if p.get("vatten_djup_m") is not None else None),
            ("Markhöjd", p.get("markhojd") or "locknivåerna (ingen markhöjd vald)"),
            ("Schakttyp", f"{SCHAKTTYPER.get(kallor['schakttyp'], kallor['schakttyp'])} (alla typer i tabellen nedan)"),
            ("Infodring", f"spill-, dag- och kombinerade ledningar från {STRUMPA_DIM_MIN} mm, så långt prislistan räcker"
                          " (inte vatten)"),
            ("Rörspräckning", f"alla ledningar till och med {SPRACKNING_DIM_MAX} mm; införings- och maskingrop per "
                              f"påbörjade {GROP_PER_M} m, servisgrop per servis, ventilgrop per "
                              f"{' och '.join(VENTILGROP_KATEGORIER)} på ledningen"),
            ("Djuptillägg", f"per påbörjade {_dk(DJUPTILLAGG_STEG_M)} m djup under {_dk(STANDARDDJUP_M)} m"),
            ("Serviser", "förnyas i båda alternativen (helschakt eller infodring per meter servisledning)"),
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

    def langd(lista):
        return round(sum(l.get("langd_m") or 0 for l in lista))

    def medeldjup(lista):
        v = [(l["djup_medel"], l.get("langd_m") or 0) for l in lista if l.get("djup_medel") is not None]
        tot = sum(w for _, w in v)
        return round(sum(d * w for d, w in v) / tot, 2) if tot else None

    def per_system(f):
        ut = []
        for s in system + ["alla"]:
            lista = [l for l in ledningar if s == "alla" or (l.get("system") or "") == s]
            ut.append([SYSTEM_NAMN.get(s, s) if s != "alla" else "Alla"] + f(lista))
        return ut

    block("Lägga om eller strumpa", ["System", "Längd (m)", "Alt. lägg om (kr)", "Alt. strumpa (kr)",
                                     "Strumpas (m)", "Rörspräcks (m)", "Läggs om (m)", "Utan kostnad (st)"],
          per_system(lambda lista: [langd(lista),
                                    _summa([l for l in lista if _jamforbar(l)], lambda l: l["_omlagg"]["summa"]),
                                    _summa([l for l in lista if _jamforbar(l)], lambda l: l["_alt"]["summa"]),
                                    langd([l for l in lista if l.get("_alt_metod") == "strumpa"]),
                                    langd([l for l in lista if l.get("_alt_metod") == "rörspräckning"]),
                                    langd([l for l in lista if l.get("_alt_metod") == "omläggning"]),
                                    sum(1 for l in lista if not _jamforbar(l))]),
          format={"Längd": M, "Alt. lägg om": KR, "Alt. strumpa": KR, "Strumpas": M, "Rörspräcks": M,
                  "Läggs om": M})
    for t in ("Alt. lägg om: alla ledningar läggs om; ledningar i samma grav delar graven.",
              "Alt. strumpa: spill/dag infodras (rörspräcks eller läggs om där infodring inte går); vatten rörspräcks, "
              "men läggs om i samma grav när en avloppsledning i graven läggs om.",
              "Summorna omfattar bara ledningar med kostnad i båda alternativen; övriga står under Utan kostnad "
              "(djup eller dimension saknas – se fliken Datakvalitet)."):
        ws.cell(rad - 1, 1, t).font = Font(italic=True, color="595959")
        rad += 1
    rad += 1

    block("Alt. lägg om per schakttyp", ["System", "Längd (m)"] + [f"{v} (kr)" for v in SCHAKTTYPER.values()] +
          ["Utan kostnad (st)"],
          per_system(lambda lista: [langd(lista)] +
                     [_summa([l for l in lista if all((l.get("_typer") or {}).get(x) for x in SCHAKTTYPER)],
                             lambda l, t=t: l["_typer"][t]["summa"]) for t in SCHAKTTYPER] +
                     [sum(1 for l in lista if not all((l.get("_typer") or {}).get(x) for x in SCHAKTTYPER))]),
          format=dict([("Längd", M)] + [(v, KR) for v in SCHAKTTYPER.values()]))

    block("Ledningsnätet", ["System", "Ledningar (st)", "Längd (m)", "Serviser (st)", "Servislängd (m)",
                            "Medeldjup (m)", "I samschakt (%)"],
          per_system(lambda lista: [len(lista), langd(lista), sum(l.get("antal_serviser") or 0 for l in lista),
                                    round(sum(l.get("serviser_m") or 0 for l in lista)), medeldjup(lista),
                                    round(100 * langd([l for l in lista if l.get("schaktgrupp")]) / langd(lista))
                                    if langd(lista) else None]),
          format={"Längd": M, "Servislängd": M, "Medeldjup": "0.00"})

    if data.get("anordningar"):
        kat = [k for k in ANORDNING_NAMN if any(a.get("kategori") == k for a in data["anordningar"])]
        block("Anordningar (se fliken Anordningar)", ["System"] + [f"{ANORDNING_NAMN[k]} (st)" for k in kat],
              per_system(lambda lista: [sum(_anordningar(l, k) or 0 for l in lista) for k in kat]))

    def fordelning(titel, nyckel, rubrik, sortering=None):
        tab = defaultdict(float)
        for l in ledningar:
            tab[(l.get("system") or "", nyckel(l))] += l.get("langd_m") or 0
        rader = [[SYSTEM_NAMN.get(s, s), k if k not in (None, "") else "(saknas)", round(v)]
                 for (s, k), v in sorted(tab.items(), key=lambda x: ("SDKV".find(x[0][0]) % 9,
                                                                       sortering(x[0][1]) if sortering else str(x[0][1])))]
        block(titel, ["System", rubrik, "Längd (m)"], rader, format={"Längd": M})

    fordelning("Material", lambda l: l.get("material"), "Material")
    fordelning("Dimension", lambda l: l.get("dimension"), "Dimension (mm)",
               lambda v: (v is None, v if isinstance(v, (int, float)) else 0))
    fordelning("Anläggningsdecennium", lambda l: f"{l['anlaggningsar'] // 10 * 10}-tal" if l.get("anlaggningsar") else None,
               "Decennium", lambda v: (v is None, str(v)))
    fordelning("Djupklass", lambda l: l.get("djupklass") or None, "Djupklass (m)",
               lambda v: (v is None, str(v).replace(">", "9")))

    tab = defaultdict(int)
    for q in data.get("datakvalitet", []):
        tab[q.get("typ")] += 1
    rader = [[k, v] for k, v in sorted(tab.items())]
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
            "langd_m": l.get("langd_m"), "djup_medel": l.get("djup_medel"), "djupklass": l.get("djupklass"),
            "antal_serviser": l.get("antal_serviser"), "serviser_m": l.get("serviser_m"),
            "schaktgrupp": l.get("schaktgrupp"), "ledningar_i_grav": l.get("_grav"),
            "kostnad_omlaggning": _kr(l.get("_omlagg")), "kostnad_strumpa": _kr(l.get("_strumpa")),
            "kostnad_rorsprackning": _kr(l.get("_sprack")),
            "anordningar_antal": l.get("anordningar_antal"),
            "alt_lagg_om_kr": _kr(l.get("_omlagg")),
            "alt_lagg_om_per_schakttyp": {t: _kr(v) for t, v in (l.get("_typer") or {}).items()},
            "alt_strumpa_metod": l.get("_alt_metod"),
            "alt_strumpa_kr": _kr(l.get("_alt")), "anmarkning": _anmarkning(l),
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
    ap.add_argument("--schakttyp", default=SCHAKTTYP, choices=sorted(SCHAKTTYPER),
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
    print(f"{len(ledningar)} ledningar, {L:,.0f} m – schakttyp {SCHAKTTYPER[a.schakttyp]}".replace(",", " "))
    jfr = [l for l in ledningar if _jamforbar(l)]
    for namn, k in (("Alt. lägg om", "_omlagg"), ("Alt. strumpa", "_alt")):
        belopp = sum(l[k]["summa"] for l in jfr)
        print(f"  {namn}: {belopp:,.0f} kr".replace(",", " "))
    if len(jfr) < len(ledningar):
        print(f"  ({len(ledningar) - len(jfr)} ledningar saknar kostnad i något alternativ och ingår inte – se Datakvalitet)")
    print(f"Skrev {xlsx}")
    print(f"Skrev {os.path.join(ut, UTDATA_JSON)}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
