#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ledning_analys.py – inventeringsbok och kostnad strumpa mot omläggning (PLAN_ledningsinventering.md steg 3)

Läser ledningsdata.json från ArcMap-verktyget Ledningsinventering och skriver

  * <utdata>/ledningsinventering.xlsx – Sammanfattning, Ledningar, Kostnad per schaktgrupp, Serviser,
                                         Parallella, Korsningar, Brunnar, Datakvalitet, Kostnadsposter
  * <utdata>/ledningsresultat.json    – djup, serviser, schaktgrupp, kostnader och rekommendation per
                                         ledning, för att fylla fält i kartan

Användning:
    python ledning_analys.py ledningsdata.json
    python ledning_analys.py ledningsdata.json -o utdata --kostnader kostnader_omlaggning.csv
    python ledning_analys.py ledningsdata.json --karta tv3_resultat/kartunderlag.json   # TV-klass per ledning

Kostnadsmodell (alla belopp i kostnadsfilerna, inga i koden):
  Strumpa     = strumpa kr/m efter dimension × längd + hatt per servis (kostnader.csv, samma som
                åtgärdspaketet i tv3_analys) + framschaktning och ny brunn när båda ändbrunnarna är
                tillsyns-/rensbrunnar. Vatten: strumpa_vatten kr/m + omkoppling per servis.
  Omläggning  = graven (schakt kr/m efter största dimension i graven, djup och antal ledningar i graven,
                delad lika mellan ledningarna) + återställning (delad på samma sätt) + serviser (omkoppling
                per st + ny servisledning per m) + ny brunn per brunnsände (halv per ände, så att en brunn
                mellan två omlagda sträckor räknas en gång).
  Kvot        = omläggning / strumpa. Rekommendationen är en sorteringshjälp, inte ett beslut.
"""
from __future__ import annotations

import argparse
import json
import os
import re
import sys
from collections import defaultdict
from datetime import datetime

HAR = os.path.dirname(os.path.abspath(__file__))
if HAR not in sys.path:
    sys.path.insert(0, HAR)
from tv3_analys import KLASS_FARG, _normlittera, las_kostnader, las_text, pris  # noqa: E402

# ============================================================================
# KONFIG – ändra här, inte i koden
# ============================================================================
KOSTNADSFIL_OMLAGGNING = os.path.join(HAR, "kostnader_omlaggning.csv")   # schakt, återställning, serviser, brunnar
KOSTNADSFIL_STRUMPA = os.path.join(HAR, "kostnader.csv")                 # strumpa, hatt, framschaktning, ny_brunn
UTDATA_XLSX = "ledningsinventering.xlsx"
UTDATA_JSON = "ledningsresultat.json"

STRUMPA_KVOT = 1.5            # strumpa när omläggning kostar minst så här många gånger mer
OMLAGG_KVOT = 1.0             # omläggning när den är billigare än strumpa (kvot under detta)
GRUND_DJUP_M = 2.0            # grund ledning: medeldjup högst så här ...
FA_SERVISER_PER_100M = 2.0    # ... och högst så här många serviser per 100 m -> omläggning
STRUMPA_DIM_MIN = 100         # strumpbara dimensioner (mm); utanför -> ej strumpbar
STRUMPA_DIM_MAX = 1200
MAX_LEDNINGAR_I_GRAV = 3      # antal ledningar i graven räknas upp till detta (prisdatabasen har 1–3)
SERVIS_OMLAGG_MAX_M = None    # meter servisledning som läggs om per servis; None = hela servisens längd
FRAMSCHAKTA_TB = True         # strumpa: båda ändbrunnarna tillsyns-/rensbrunn -> framschaktning + ny brunn
BRUNNSTYP_TILLSYN = r"TB|RB"  # brunnstyp (t.ex. WellFunction STB/SRB) som räknas som tillsyns-/rensbrunn

# Kolumner i fliken Ledningar som döljs (grupperade, fälls ut med plustecknet)
DOLDA_LEDNINGAR = [
    "Ledningstyp", "Renoveringsår", "VG från", "VG till", "Lutning GIS", "Djup min", "Djup max", "Djupkälla",
    "Täckning min", "Djup start", "Djup slut", "Samschakt med", "Parallella", "Korsningar",
    "Schakt", "Återställning", "Serviser omläggning", "Brunnar", "Hattar", "Framschaktning",
    "K-index", "TV-flagga", "Område", "Lager", "OID", "Del",
]

SYSTEM_NAMN = {"S": "Spillvatten", "D": "Dagvatten", "K": "Kombinerat", "V": "Vatten", "": "Okänt"}
REK_FARG = {"strumpa": "C6E0B4", "omläggning": "F8CBAD", "bedöm": "FFE699"}


# ============================================================================
# Kostnader
# ============================================================================

KOLUMNER_OMLAGG = ["post", "dimension_fran", "dimension_till", "djup_fran", "djup_till", "ledningar",
                   "enhet", "kr", "kommentar"]


def _tal(v: str | None) -> float | None:
    v = re.sub(r"[^\d,.\-]", "", v or "")
    if "," in v:
        v = v.replace(".", "").replace(",", ".")
    try:
        return float(v) if v else None
    except ValueError:
        return None


def las_kostnader_omlaggning(path: str) -> list[dict]:
    """Läser kostnader_omlaggning.csv. Rubrikraden (som börjar med 'post') styr kolumnordningen;
    utan rubrikrad gäller KOLUMNER_OMLAGG. Returnerar [{post, dim_fran, dim_till, djup_fran, djup_till,
    ledningar, enhet, kr, kommentar}]."""
    poster, kolumner = [], None
    for rad in las_text(path).splitlines():
        rad = rad.split("#", 1)[0].strip()
        if not rad:
            continue
        d = [x.strip() for x in re.split(r"[;\t]", rad)]
        if d[0].lower() == "post":
            kolumner = [k.lower() for k in d]
            continue
        kol = kolumner or KOLUMNER_OMLAGG
        v = dict(zip(kol, d + [""] * (len(kol) - len(d))))
        if not v.get("post"):
            continue
        kr = _tal(v.get("kr"))
        if kr is None:
            print(f"  VARNING {os.path.basename(path)}: kan inte tolka beloppet '{v.get('kr')}' för {v['post']} – hoppas över")
            continue
        n = _tal(v.get("ledningar"))
        poster.append({"post": v["post"].lower(), "dim_fran": _tal(v.get("dimension_fran")),
                       "dim_till": _tal(v.get("dimension_till")), "djup_fran": _tal(v.get("djup_fran")),
                       "djup_till": _tal(v.get("djup_till")), "ledningar": int(n) if n else None,
                       "enhet": v.get("enhet", ""), "kr": kr, "kommentar": v.get("kommentar", "")})
    return poster


def _dim_ok(k: dict, dim: float | None) -> bool:
    if k["dim_fran"] is None and k["dim_till"] is None:
        return True
    if dim is None:
        return False
    return (k["dim_fran"] is None or dim >= k["dim_fran"]) and (k["dim_till"] is None or dim <= k["dim_till"])


def _djuprad(rader: list[dict], djup: float | None) -> dict | None:
    """Raden vars djupintervall gäller: den med största djup_fran ≤ djup (gränsen 2,0 m hör till 2–3 m,
    djup över sista intervallet får det sista). Rader utan djup gäller alla djup."""
    utan = [k for k in rader if k["djup_fran"] is None and k["djup_till"] is None]
    med = [k for k in rader if k not in utan]
    if djup is None or not med:
        return utan[0] if utan else None
    under = [k for k in med if (k["djup_fran"] or 0.0) <= djup]
    if under:
        return max(under, key=lambda k: k["djup_fran"] or 0.0)
    return min(med, key=lambda k: k["djup_fran"] or 0.0)        # negativt djup: grundaste raden


def enkelpris(kost: list[dict], post: str, dim: float | None = None) -> float | None:
    rader = [k for k in kost if k["post"] == post and _dim_ok(k, dim)]
    return rader[0]["kr"] if rader else None


def gravpris(kost: list[dict], dim: float | None, djup: float | None, n: int) -> tuple[float | None, str]:
    """kr per meter grav med n ledningar. Rader med ledningar = n används direkt; annars raden för en
    ledning × (1 + samschakt_faktor × (n − 1)). Returnerar (kr, hur)."""
    rader = [k for k in kost if k["post"] == "schakt" and _dim_ok(k, dim)]
    if not rader:
        return None, "inget schaktpris för dimensionen"
    med_n = [k for k in rader if k["ledningar"] == n]
    if med_n:
        r = _djuprad(med_n, djup)
        if r is not None:
            return r["kr"], f"{n} ledningar i graven"
    en = [k for k in rader if k["ledningar"] in (None, 1)]
    r = _djuprad(en, djup)
    if r is None:
        return None, "inget schaktpris för djupet"
    if n <= 1:
        return r["kr"], "1 ledning"
    f = enkelpris(kost, "samschakt_faktor") or 0.0
    return r["kr"] * (1 + f * (n - 1)), f"{n} ledningar (faktor {f:g})"


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


def ledningar_i_grav(led: dict) -> int:
    """1 + summan av andelen av ledningens längd som delas med varje samschakt-partner, avrundat."""
    andel = sum(p.get("andel") or 0.0 for p in led.get("parallella") or [] if p.get("samschakt"))
    return max(1, min(MAX_LEDNINGAR_I_GRAV, 1 + int(round(andel))))


def tv_index(kartfiler: list[str]) -> dict:
    """{frozenset(brunnspar): sämsta TV-post} ur kartunderlag.json från tv3_analys."""
    ut: dict = {}
    for path in kartfiler:
        with open(path, encoding="utf-8") as f:
            data = json.load(f)
        for p in data.get("strackor", []):
            if p.get("filmstatus") and ("ersatt" in p["filmstatus"] or "ingår" in p["filmstatus"]):
                continue
            par = frozenset((_normlittera(p.get("startbrunn") or ""), _normlittera(p.get("slutbrunn") or "")))
            klass = p.get("gallande_bedomning") or p.get("maskinell_bedomning") or "E"
            gammal = ut.get(par)
            if gammal is None or klass < (gammal.get("gallande_bedomning") or gammal.get("maskinell_bedomning") or "E"):
                ut[par] = p
    return ut


def berakna(data: dict, kost_o: list[dict], kost_s: list[dict], tv: dict) -> list[dict]:
    """Lägger till kostnader, rekommendation och TV-uppgifter på varje ledning (i data['ledningar'])."""
    ledningar = data.get("ledningar", [])
    per_id = {l.get("id"): l for l in ledningar}
    brunnar = {b["littera"]: b for b in data.get("brunnar", [])}
    serviser = defaultdict(list)
    for s in data.get("serviser", []):
        if s.get("ledning"):
            serviser[s["ledning"]].append(s)
    f_sam = enkelpris(kost_o, "samschakt_faktor") or 0.0

    for led in ledningar:
        L = led.get("langd_m") or 0.0
        dim = led.get("dimension")
        system = led.get("system") or ""
        n = ledningar_i_grav(led)
        partners = [per_id.get(p["ledning"]) for p in led.get("parallella") or [] if p.get("samschakt")]
        dimensioner = [d for d in [dim] + [q.get("dimension") for q in partners if q] if d]
        dim_grav = max(dimensioner) if dimensioner else None
        antal_serv = led.get("antal_serviser") or 0
        serv_m = sum(min(s.get("langd_m") or 0.0, SERVIS_OMLAGG_MAX_M) if SERVIS_OMLAGG_MAX_M is not None
                     else (s.get("langd_m") or 0.0) for s in serviser.get(led.get("id"), []))
        if not serviser.get(led.get("id")):
            serv_m = led.get("serviser_m") or 0.0

        # --- TV-klass ur kartunderlaget ---------------------------------------------------------
        par = frozenset((_normlittera(led.get("fran") or ""), _normlittera(led.get("till") or "")))
        tvp = tv.get(par) if all(par) and len(par) == 2 else None
        led["_tv"] = tvp

        # --- omläggning -----------------------------------------------------------------------
        o: dict = {}
        o_brist = ""
        djupen = _djup_punkter(led) if system != "V" else []
        if djupen:
            priser = [gravpris(kost_o, dim_grav, d, n)[0] for d in djupen]
            priser = [p for p in priser if p is not None]
            grav_m = sum(priser) / len(priser) if priser else None
            hur = gravpris(kost_o, dim_grav, led.get("djup_medel"), n)[1]
        else:
            grav_m, hur = gravpris(kost_o, dim_grav, led.get("djup_medel"), n)
        if led.get("djup_medel") is None and not djupen:
            o_brist = "djup saknas"
        elif grav_m is None:
            o_brist = hur
        if not o_brist:
            o["schakt"] = grav_m / n * L
            aterst = enkelpris(kost_o, "aterstallning") or 0.0
            o["aterstallning"] = aterst * (1 + f_sam * (n - 1)) / n * L
            o["serviser"] = antal_serv * (enkelpris(kost_o, "servis_omkoppling") or 0.0) + \
                serv_m * (enkelpris(kost_o, "servis_ledning") or 0.0)
            andar = sum(0.5 for b in (led.get("fran"), led.get("till")) if b) if system != "V" else 0.0
            o["brunnar"] = andar * (enkelpris(kost_o, "brunn_ny") or 0.0)
            o["summa"] = sum(o.values())
        led["_omlagg"] = o or None
        led["_omlagg_brist"] = o_brist
        led["_grav"] = n
        led["_grav_hur"] = hur

        # --- strumpa ----------------------------------------------------------------------------
        s: dict = {}
        s_brist = ""
        if dim is None:
            s_brist = "dimension saknas"
        elif not (STRUMPA_DIM_MIN <= dim <= STRUMPA_DIM_MAX):
            s_brist = f"dimension {dim} mm utanför {STRUMPA_DIM_MIN}–{STRUMPA_DIM_MAX}"
        elif system == "V":
            kr_m = enkelpris(kost_o, "strumpa_vatten", dim)
            if kr_m is None:
                s_brist = "inget pris för strumpa_vatten"
            else:
                s["strumpa"] = kr_m * L
                s["hattar"] = antal_serv * (enkelpris(kost_o, "servis_omkoppling") or 0.0)
        else:
            kr_m = pris(kost_s, "strumpa", int(dim))
            if kr_m is None:
                s_brist = "inget strumpapris för dimensionen"
            else:
                s["strumpa"] = kr_m * L
                s["hattar"] = antal_serv * (pris(kost_s, "hatt") or 0.0)
                typer = [(brunnar.get(b) or {}).get("typ") or "" for b in (led.get("fran"), led.get("till"))]
                if FRAMSCHAKTA_TB and all(typer) and all(re.search(BRUNNSTYP_TILLSYN, t.upper()) for t in typer):
                    s["framschaktning"] = (pris(kost_s, "framschaktning") or 0.0) + (pris(kost_s, "ny_brunn") or 0.0)
        if s:
            s["summa"] = sum(s.values())
        led["_strumpa"] = s or None
        led["_strumpa_brist"] = s_brist

        so = led["_omlagg"]["summa"] if led["_omlagg"] else None
        ss = led["_strumpa"]["summa"] if led["_strumpa"] else None
        led["_kvot"] = so / ss if (so is not None and ss) else None
        led["_rek"], led["_motiv"] = rekommendation(led)
    return ledningar


def rekommendation(led: dict) -> tuple[str, str]:
    """(strumpa | omläggning | bedöm, motiv). Metod om ledningen ska åtgärdas – inte om den ska det."""
    tvp = led.get("_tv") or {}
    o, s, kvot = led.get("_omlagg"), led.get("_strumpa"), led.get("_kvot")
    if tvp.get("relinad"):
        return "bedöm", "redan relinad enligt TV-inspektionen"
    if s is None:
        if o is not None:
            return "omläggning", "ej strumpbar: " + led.get("_strumpa_brist", "")
        return "bedöm", "ej strumpbar (" + led.get("_strumpa_brist", "") + ") och " + (led.get("_omlagg_brist") or "")
    if o is None:
        return "bedöm", "omläggningskostnad saknas: " + (led.get("_omlagg_brist") or "")
    if "grad 4" in (tvp.get("atgardsflagga") or ""):
        return "bedöm", "grad 4 RBR/DEF enligt TV – går strumpa?"
    if kvot >= STRUMPA_KVOT:
        return "strumpa", f"omläggning {kvot:.1f} × strumpa".replace(".", ",")
    if kvot < OMLAGG_KVOT:
        return "omläggning", "billigare än strumpa"
    djup = led.get("djup_medel")
    L = led.get("langd_m") or 0.0
    per100 = (led.get("antal_serviser") or 0) / max(L, 1.0) * 100
    if djup is not None and djup <= GRUND_DJUP_M and per100 <= FA_SERVISER_PER_100M:
        return "omläggning", f"grund ({djup:.1f} m) och få serviser".replace(".", ",")
    if led.get("_grav", 1) >= 2:
        return "omläggning", f"samschakt – {led['_grav']} ledningar i graven"
    return "bedöm", f"kvot {kvot:.2f}".replace(".", ",")


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


def skriv_excel(data: dict, ledningar: list[dict], path: str, kost_o: list[dict], kost_s: list[dict],
                kallor: dict) -> None:
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill

    wb = Workbook()
    ws_s = wb.active
    ws_s.title = "Sammanfattning"
    KR, M, M2, P, K = "#,##0", "0.0", "0.00", "0.0", "0.00"

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
           "Djup start (m)", "Djup slut (m)", "Serviser (st)", "Servislängd (m)", "Schaktgrupp",
           "Ledningar i graven (st)", "Samschakt med", "Parallella", "Korsningar",
           "Strumpa (kr)", "Omläggning (kr)", "Kvot", "Rekommendation", "Motiv",
           "Schakt (kr)", "Återställning (kr)", "Serviser omläggning (kr)", "Brunnar (kr)", "Hattar (kr)",
           "Framschaktning (kr)", "TV-klass", "K-index", "TV-flagga", "Flaggor", "Område", "Lager", "OID", "Del"]
    egna = {_dela_rubrik(k)[0] for k in kol}
    attrkol = [a if a not in egna else a + " (GIS)" for a in attrnamn]
    rader = []
    ordn = sorted(ledningar, key=lambda l: ("SDKV".find(l.get("system") or "?") % 9, str(l.get("id"))))
    for l in ordn:
        o, s, tvp = l.get("_omlagg"), l.get("_strumpa"), l.get("_tv") or {}
        rad = [l.get("id"), l.get("system"), l.get("fran") or "(fri ände)", l.get("till") or "(fri ände)",
               _r(l.get("langd_m")), l.get("dimension"), l.get("material"), l.get("ledningstyp"),
               l.get("anlaggningsar"), l.get("renoveringsar"), l.get("vg_fran"), l.get("vg_till"),
               l.get("lutning_promille"), l.get("lutning_gis"), l.get("djup_medel"), l.get("djup_min"),
               l.get("djup_max"), l.get("djupklass") or None, l.get("djup_kalla") or None, l.get("tackning_min"),
               l.get("djup_start"), l.get("djup_slut"), l.get("antal_serviser"), _r(l.get("serviser_m")),
               l.get("schaktgrupp"), l.get("_grav"), ", ".join(l.get("schakt_ledningar") or []) or None,
               ", ".join(f"{p['ledning']} ({p['system']}, {round(100 * (p.get('andel') or 0))} %)"
                         for p in l.get("parallella") or []) or None,
               ", ".join(f"{k['ledning']} vid {k['station_m']:.1f} m".replace(".", ",")
                         for k in l.get("korsningar") or []) or None,
               _kr(s), _kr(o), _r(l.get("_kvot"), 2), l.get("_rek"), l.get("_motiv"),
               _kr(o, "schakt"), _kr(o, "aterstallning"), _kr(o, "serviser"), _kr(o, "brunnar"),
               _kr(s, "hattar"), _kr(s, "framschaktning"),
               tvp.get("gallande_bedomning") or tvp.get("maskinell_bedomning"), tvp.get("konstruktionsindex"),
               tvp.get("atgardsflagga") or None, ", ".join(l.get("flaggor") or []) or None, l.get("omrade") or None,
               l.get("lager"), l.get("oid"), l.get("del")]
        rad += [(l.get("attribut") or {}).get(a) for a in attrnamn]
        rader.append(rad)
    format_ = {"Längd": M, "VG från": "0.000", "VG till": "0.000", "Lutning": P, "Lutning GIS": P, "Djup medel": M2,
               "Djup min": M2, "Djup max": M2, "Täckning min": M2, "Djup start": M2, "Djup slut": M2,
               "Servislängd": M, "Strumpa": KR, "Omläggning": KR, "Kvot": K, "Schakt": KR, "Återställning": KR,
               "Serviser omläggning": KR, "Brunnar": KR, "Hattar": KR, "Framschaktning": KR}
    start = tabell(ws, kol + attrkol, rader, bredder={"Motiv": 34, "Flaggor": 30, "ID": 14},
                   dolda=DOLDA_LEDNINGAR + attrkol, format=format_)
    ci_rek = kol.index("Rekommendation") + 1
    ci_tv = kol.index("TV-klass") + 1
    for ri in range(start, start + len(rader)):
        c = ws.cell(ri, ci_rek)
        if c.value in REK_FARG:
            c.fill = PatternFill("solid", fgColor=REK_FARG[c.value])
        c = ws.cell(ri, ci_tv)
        if c.value and c.value[0] in KLASS_FARG:
            c.fill = PatternFill("solid", fgColor=KLASS_FARG[c.value[0]])
            c.font = Font(bold=True, color="FFFFFF" if c.value[0] in "AB" else "000000")

    # --- Kostnad per schaktgrupp ----------------------------------------------------------------
    grupper = defaultdict(list)
    for l in ledningar:
        grupper[l.get("schaktgrupp")].append(l)

    def blandat(l):
        o, s = l.get("_omlagg"), l.get("_strumpa")
        if l.get("_rek") == "strumpa" and s:
            return s["summa"]
        if l.get("_rek") == "omläggning" and o:
            return o["summa"]
        v = [d["summa"] for d in (o, s) if d]
        return min(v) if v else None

    def summa(lista, f):
        v = [f(l) for l in lista]
        return round(sum(x for x in v if x is not None)) if any(x is not None for x in v) else None

    def saknas(lista, nyckel):
        return sum(1 for l in lista if not l.get(nyckel))

    ws = wb.create_sheet("Kostnad per schaktgrupp")
    kol = ["Schaktgrupp", "System", "Ledningar (st)", "Ledningar", "Längd (m)", "Strumpa alla (kr)",
           "Utan strumppris (st)", "Lägg om alla (kr)", "Utan omläggningspris (st)", "Enligt rekommendation (kr)",
           "Strumpa (st)", "Omläggning (st)", "Bedöm (st)"]
    rader = []
    for g in sorted(k for k in grupper if k is not None):
        lista = grupper[g]
        rader.append([g, "+".join(sorted({l.get("system") or "?" for l in lista})), len(lista),
                      ", ".join(str(l.get("id")) for l in lista), _r(sum(l.get("langd_m") or 0 for l in lista)),
                      summa(lista, lambda l: (l.get("_strumpa") or {}).get("summa")), saknas(lista, "_strumpa"),
                      summa(lista, lambda l: (l.get("_omlagg") or {}).get("summa")), saknas(lista, "_omlagg"),
                      summa(lista, blandat), sum(1 for l in lista if l.get("_rek") == "strumpa"),
                      sum(1 for l in lista if l.get("_rek") == "omläggning"),
                      sum(1 for l in lista if l.get("_rek") == "bedöm")])
    for namn_, lista in (("utan samschakt", grupper.get(None, [])), ("totalt", ledningar)):
        if not lista:
            continue
        rader.append([namn_, "+".join(sorted({l.get("system") or "?" for l in lista})), len(lista), None,
                      _r(sum(l.get("langd_m") or 0 for l in lista)),
                      summa(lista, lambda l: (l.get("_strumpa") or {}).get("summa")), saknas(lista, "_strumpa"),
                      summa(lista, lambda l: (l.get("_omlagg") or {}).get("summa")), saknas(lista, "_omlagg"),
                      summa(lista, blandat), sum(1 for l in lista if l.get("_rek") == "strumpa"),
                      sum(1 for l in lista if l.get("_rek") == "omläggning"),
                      sum(1 for l in lista if l.get("_rek") == "bedöm")])
    start = tabell(ws, kol, rader, bredder={"Ledningar": 40},
                   format={"Längd": M, "Strumpa alla": KR, "Lägg om alla": KR, "Enligt rekommendation": KR})
    for ri in range(start, start + len(rader)):
        if isinstance(ws.cell(ri, 1).value, str):
            for c in ws[ri]:
                c.font = Font(bold=True)

    # --- Serviser, Parallella, Korsningar, Brunnar, Datakvalitet ----------------------------------
    ws = wb.create_sheet("Serviser")
    rader = [[s.get("id"), s.get("ledning") or "(ingen huvudledning)", s.get("system"), s.get("position_m"),
              _r(s.get("langd_m")), s.get("dimension"), s.get("material") or None, s.get("ledningstyp") or None,
              s.get("vg_vid_anslutning"), s.get("djup_vid_anslutning"), s.get("lager"), s.get("oid")]
             for s in data.get("serviser", [])]
    tabell(ws, ["ID", "Huvudledning", "System", "Läge (m)", "Längd (m)", "Dimension (mm)", "Material", "Ledningstyp",
                "VG vid anslutning (m)", "Djup vid anslutning (m)", "Lager", "OID"], rader,
           format={"Läge": M, "Längd": M, "Djup vid anslutning": M2})

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
        if l.get("_strumpa_brist") and l.get("_strumpa_brist") != "inget pris för strumpa_vatten":
            rader.append(["ingen strumpkostnad", l.get("id"), l["_strumpa_brist"]])
        if l.get("_omlagg_brist"):
            rader.append(["ingen omläggningskostnad", l.get("id"), l["_omlagg_brist"]])
        if l.get("dimension") is None:
            rader.append(["dimension saknas", l.get("id"), ""])
        if l.get("anlaggningsar") is None:
            rader.append(["anläggningsår saknas", l.get("id"), ""])
        if not l.get("fran") or not l.get("till"):
            rader.append(["ledning utan brunn i någon ände", l.get("id"), "fri ände (servis, T-rör eller fel i GIS)"])
    tabell(ws, ["Typ", "ID", "Beskrivning"], rader, bredder={"Beskrivning": 70, "Typ": 32})

    ws = wb.create_sheet("Kostnadsposter")
    rader = [[k["post"], k["dim_fran"], k["dim_till"], k["djup_fran"], k["djup_till"], k["ledningar"], k["enhet"],
              k["kr"], k["kommentar"], os.path.basename(kallor["omlaggning"])] for k in kost_o]
    rader += [[k["post"], k["fran"], k["till"], None, None, None, k["enhet"], k["kr"], k["kommentar"],
               os.path.basename(kallor["strumpa"])] for k in kost_s]
    tabell(ws, ["Post", "Dimension från (mm)", "Dimension till (mm)", "Djup från (m)", "Djup till (m)",
                "Ledningar i graven", "Enhet", "Belopp (kr)", "Kommentar", "Fil"], rader,
           bredder={"Kommentar": 60}, format={"Belopp": "#,##0.##"})

    skriv_sammanfattning(ws_s, data, ledningar, kost_o, kost_s, kallor, blandat)
    wb.save(path)


def skriv_sammanfattning(ws, data, ledningar, kost_o, kost_s, kallor, blandat) -> None:
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
            ("Kostnadsfiler", f"{os.path.basename(kallor['omlaggning'])}, {os.path.basename(kallor['strumpa'])}"),
            ("TV-inspektioner", ", ".join(os.path.basename(x) for x in kallor["karta"]) or "–"),
            ("Analys", datetime.now().strftime("%Y-%m-%d %H:%M"))]
    if any("schablon" in (k.get("kommentar") or "").lower() for k in kost_o + kost_s):
        info.append(("OBS", "kostnadsfilerna innehåller schablonbelopp – byt mot egna priser"))
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

    rader = []
    for s in system + ["alla"]:
        lista = [l for l in ledningar if s == "alla" or (l.get("system") or "") == s]
        sam = [l for l in lista if l.get("schaktgrupp")]
        rader.append([SYSTEM_NAMN.get(s, s) if s != "alla" else "Alla", len(lista), langd(lista),
                      sum(l.get("antal_serviser") or 0 for l in lista),
                      round(sum(l.get("serviser_m") or 0 for l in lista)), medeldjup(lista),
                      round(100 * langd(sam) / langd(lista)) if langd(lista) else None])
    block("Ledningsnätet", ["System", "Ledningar (st)", "Längd (m)", "Serviser (st)", "Servislängd (m)",
                            "Medeldjup (m)", "I samschakt (%)"], rader,
          format={"Längd": M, "Servislängd": M, "Medeldjup": "0.00"})

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

    rader = []
    for s in system + ["alla"]:
        lista = [l for l in ledningar if s == "alla" or (l.get("system") or "") == s]

        def summa(f):
            v = [f(l) for l in lista if f(l) is not None]
            return round(sum(v)) if v else None
        rader.append([SYSTEM_NAMN.get(s, s) if s != "alla" else "Alla", langd(lista),
                      summa(lambda l: (l.get("_strumpa") or {}).get("summa")),
                      summa(lambda l: (l.get("_omlagg") or {}).get("summa")), summa(blandat),
                      sum(1 for l in lista if l.get("_rek") == "strumpa"),
                      sum(1 for l in lista if l.get("_rek") == "omläggning"),
                      sum(1 for l in lista if l.get("_rek") == "bedöm")])
    block("Kostnad (alla ledningar, oavsett behov)", ["System", "Längd (m)", "Strumpa alla (kr)",
                                                        "Lägg om alla (kr)", "Enligt rekommendation (kr)",
                                                        "Strumpa (st)", "Omläggning (st)", "Bedöm (st)"], rader,
          format={"Längd": M, "Strumpa alla": KR, "Lägg om alla": KR, "Enligt rekommendation": KR})

    tv = [l for l in ledningar if l.get("_tv")]
    if kallor["karta"]:
        tab = defaultdict(lambda: [0, 0.0])
        for l in tv:
            k = (l["_tv"].get("gallande_bedomning") or l["_tv"].get("maskinell_bedomning") or "?", l.get("_rek"))
            tab[k][0] += 1
            tab[k][1] += l.get("langd_m") or 0
        rader = [[k, r_, n, round(m)] for (k, r_), (n, m) in sorted(tab.items())]
        rader.append(["utan TV-inspektion", None, len(ledningar) - len(tv),
                      round(sum(l.get("langd_m") or 0 for l in ledningar if not l.get("_tv")))])
        block("TV-klass och rekommenderad metod", ["TV-klass", "Rekommendation", "Ledningar (st)", "Längd (m)"],
              rader, format={"Längd": M})

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
        tvp = l.get("_tv") or {}
        poster.append({
            "id": l.get("id"), "lager": l.get("lager"), "oid": l.get("oid"), "del": l.get("del"),
            "fran": l.get("fran"), "till": l.get("till"), "system": l.get("system"),
            "langd_m": l.get("langd_m"), "djup_medel": l.get("djup_medel"), "djupklass": l.get("djupklass"),
            "antal_serviser": l.get("antal_serviser"), "serviser_m": l.get("serviser_m"),
            "schaktgrupp": l.get("schaktgrupp"), "ledningar_i_grav": l.get("_grav"),
            "kostnad_strumpa": _kr(l.get("_strumpa")), "kostnad_omlaggning": _kr(l.get("_omlagg")),
            "kvot": _r(l.get("_kvot"), 2), "rekommendation": l.get("_rek"), "motiv": l.get("_motiv"),
            "tv_klass": tvp.get("gallande_bedomning") or tvp.get("maskinell_bedomning"),
        })
    with open(path, "w", encoding="utf-8") as f:
        json.dump({"version": 1, "kalla": "ledning_analys.py", "ledningsdata": kalla,
                   "genererad": datetime.now().strftime("%Y-%m-%d %H:%M"), "ledningar": poster},
                  f, ensure_ascii=False, indent=1)


# ============================================================================
# main
# ============================================================================

def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="Inventeringsbok och kostnad strumpa mot omläggning ur ledningsdata.json")
    ap.add_argument("ledningsdata", help="ledningsdata.json från ArcMap-verktyget Ledningsinventering")
    ap.add_argument("-o", "--utdata", help="utdatamapp (standard: samma mapp som ledningsdata.json)")
    ap.add_argument("--kostnader", default=KOSTNADSFIL_OMLAGGNING, help="kostnadsfil för omläggning")
    ap.add_argument("--strumpa", default=KOSTNADSFIL_STRUMPA, help="kostnadsfil för strumpa (kostnader.csv)")
    ap.add_argument("--karta", action="append", default=[],
                    help="kartunderlag.json från tv3_analys – TV-klass per ledning (kan upprepas)")
    a = ap.parse_args(argv)

    with open(a.ledningsdata, encoding="utf-8") as f:
        data = json.load(f)
    if data.get("kalla") != "inventering.py":
        print("OBS: filen kommer inte från Ledningsinventering – djup, serviser och parallella ledningar saknas"
              " (kör verktyget Ledningsinventering i ArcMap).")
    for path, namn in ((a.kostnader, "--kostnader"), (a.strumpa, "--strumpa")):
        if not os.path.isfile(path):
            print(f"Kostnadsfilen finns inte: {path} ({namn})")
            return 2
    kost_o = las_kostnader_omlaggning(a.kostnader)
    kost_s = las_kostnader(a.strumpa)
    tv = tv_index(a.karta) if a.karta else {}

    ledningar = berakna(data, kost_o, kost_s, tv)
    ut = a.utdata or os.path.dirname(os.path.abspath(a.ledningsdata))
    os.makedirs(ut, exist_ok=True)
    xlsx = os.path.join(ut, UTDATA_XLSX)
    kallor = {"ledningsdata": os.path.abspath(a.ledningsdata), "omlaggning": a.kostnader, "strumpa": a.strumpa,
              "karta": a.karta}
    try:
        skriv_excel(data, ledningar, xlsx, kost_o, kost_s, kallor)
    except PermissionError:
        print(f"Kan inte skriva {xlsx} – är filen öppen i Excel? Stäng den och kör igen.")
        return 1
    skriv_resultat(ledningar, os.path.join(ut, UTDATA_JSON), os.path.abspath(a.ledningsdata))

    rek = defaultdict(int)
    for l in ledningar:
        rek[l.get("_rek")] += 1
    L = sum(l.get("langd_m") or 0 for l in ledningar)
    print(f"{len(ledningar)} ledningar, {L:,.0f} m".replace(",", " "))
    print("Rekommendation: " + ", ".join(f"{k} {v}" for k, v in sorted(rek.items())))
    if tv:
        print(f"TV-klass för {sum(1 for l in ledningar if l.get('_tv'))} ledningar")
    print(f"Skrev {xlsx}")
    print(f"Skrev {os.path.join(ut, UTDATA_JSON)}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
