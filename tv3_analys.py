#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
tv3_analys.py – Analys av TV-inspektioner (Svenskt Vatten TV-fil v3.0 / P93)

Läser en eller flera .TV3-filer, poängsätter varje ledningssträcka utifrån
observationernas grad (1–4) och tar fram ett presentationsunderlag:

  * <utdata>/prioritering.xlsx   – Excel med sammanfattning, prioriteringslista,
                                    alla observationer, kodstatistik
  * <utdata>/diagram/*.png        – bara med --diagram; diagrammen bäddas annars
                                    enbart in i Excel-filen
  * <utdata>/rapporter/*.pdf      – inspektionsprotokoll per sträcka (kräver reportlab)
  * <utdata>/kartunderlag.json    – underlag för ArcMap-skriptet i arcmap/

Videofiler och bilder: TV3-filen innehåller bara filnamnen. Skriptet söker
igenom TV3-filens katalog (och undermappar) samt eventuella extra kataloger
(--media) efter filerna och lägger klickbara länkar i Excel.

Användning:
    python tv3_analys.py fil1.TV3 [fil2.TV3 ...] [-o utdata] [--topp 15]
    python tv3_analys.py -l filer.txt              # textfil med en TV3-sökväg per rad
    python tv3_analys.py katalog/                  # alla .TV3-filer i katalogen (rekursivt)
    python tv3_analys.py -l filer.txt --media D:/Filmer   # extra katalog att söka videofiler i

Listfilen (-l/--lista) är en vanlig textfil. Tomma rader och rader som börjar
med # hoppas över. Relativa sökvägar tolkas relativt listfilens katalog.

    # exempel på listfil
    media: D:\\Inspektioner\\Filmer            mediakatalog som gäller alla filer
    DUF 701.TV3                                TV3-fil (media söks även i filens egen katalog)
    DUF 702.TV3 ; D:\\Filmer\\DUF702 ; E:\\Bilder  TV3-fil med egna mediakataloger
    C:\\Inspektioner\\2022\\                    katalog: alla TV3-filer i den

Alla filer analyseras tillsammans i ett gemensamt underlag; kolumnen "Fil" i
Excel visar varifrån varje sträcka kommer.

Poängmodell (kan justeras i KONFIG nedan):
    grad 1 = 1 p, grad 2 = 3 p, grad 3 = 10 p, grad 4 = 30 p
    Konstruktionsskador (SPR, RBR, DEF, YTS, FOG, FRF, DEA) räknas fullt.
    Driftskador (ROT, INL, UTF, SED, INH) räknas med faktor 0,5.
    Poängen normeras till poäng per 100 m ledning (sträckor kortare än
    MINLANGD räknas som MINLANGD, så att mycket korta sträckor inte överdrivs).
    Löpande skador (A1…B1) räknas en gång (vid startmarkeringen).

Prioritetsklass (för renoveringsbehov):
    A – Åtgärd snarast    : konstruktionsgrad 4, eller konstruktionsindex ≥ 80 p/100 m
    B – Planera renovering: konstruktionsgrad 3, eller konstruktionsindex ≥ 25 p/100 m
    C – Bevaka            : övriga sträckor med registrerade skador
    D – Inga skador       : inga skadeobservationer
    E – Ej bedömd         : ingen inspekterad längd (< 1 m)
Inom varje klass rangordnas sträckorna efter totalindex.
Driftåtgärd (spolning/rotskärning) flaggas separat när driftgrad ≥ 3.
"""

from __future__ import annotations

import argparse
import math
import os
import re
import sys
from collections import Counter, defaultdict
from dataclasses import dataclass, field
from datetime import datetime

# ----------------------------------------------------------------------------
# KONFIG – justera här om ni har en egen viktning
# ----------------------------------------------------------------------------

GRADPOANG = {1: 1, 2: 3, 3: 10, 4: 30}

# Observationskoder (P93). typ: K = konstruktion, D = drift, I = information
KODER = {
    "SPR": ("Sprickor", "K"),
    "RBR": ("Rörbrott", "K"),
    "DEF": ("Deformation", "K"),
    "YTS": ("Ytskada", "K"),
    "FOG": ("Fogförskjutning", "K"),
    "FRF": ("Fog, rörfel", "K"),
    "DEA": ("Defekt anslutning", "K"),
    "ROT": ("Rötter", "D"),
    "INL": ("Inläckage", "D"),
    "UTF": ("Utfällning", "D"),
    "SED": ("Sediment", "D"),
    "INH": ("Inträngande hinder", "D"),
    "KAM": ("Kamera/inspektion avbruten", "I"),
}
DRIFTFAKTOR = 0.5

# Attribut/typkoder (kolumn 8) – för läsbara texter
ATTRIBUT = {
    "KOMPL": "komplexa", "CIRK": "cirkulära", "LÄNGS": "längsgående",
    "TUNNA": "tunna", "GROVA": "grova", "PAKET": "rotpaket",
    "RIKTN": "riktningsavvikelse", "TVÄRS": "tvärsförskjutning", "LÄFÖR": "längsförskjutning",
    "FAFOG": "fel i fog", "UTFÄL": "utfällning", "FINT": "fint", "GROVT": "grovt",
    "HINDE": "hinder – inspektion avbruten", "ANNAN": "annan orsak", "EJÖPP": "ej öppen",
    # ej-skade-koder
    "PGREN": "påstick/grenrör", "INHUG": "inhuggen", "BORRA": "borrad", "SAHUG": "sadelhuggen",
    "LAGAT": "lagat", "BÖJD": "böj", "PLAST": "plast", "BTG": "betong", "GJUTJ": "gjutjärn",
    "RAMP": "ramp", "UTNED": "utvändig nedgång",
}

# Koder utan skadegrad (kolumn 7) – information om ledningen
INFOKODER = {
    "AS": "Anslutning", "RB": "Rensbrunn", "NB": "Nedstigningsbrunn", "TB": "Tillsynsbrunn",
    "RP": "Reparation", "BR": "Böj", "DF": "Dimensionsförändring", "MF": "Materialförändring",
    "ST": "Stalp", "PP": "Pumpstation/övrigt", "AG": "Anslutning, annan",
}

KLASS_TEXT = {
    "A": "A – Åtgärd snarast",
    "B": "B – Planera renovering",
    "C": "C – Bevaka",
    "D": "D – Inga skador",
    "E": "E – Ej bedömd",
}
KLASS_FARG = {"A": "D03B3B", "B": "EC835A", "C": "FAB219", "D": "0CA30C", "E": "BFBFBF"}
KLASS_FARG_HEX = {k: "#" + v for k, v in KLASS_FARG.items()}

TROSKEL_A = 80.0   # konstruktionsindex p/100 m
TROSKEL_B = 25.0
MINLANGD = 20.0    # m – nämnare vid normering av korta sträckor

# Relinade (infodrade) sträckor redovisas som ett eget material i fliken Material och i
# diagrammet "Prioritetsklass per material", i stället för som rörets ursprungsmaterial.
# Sätt RELINAD_SOM_MATERIAL = False för att räkna dem som betong/plast som tidigare.
RELINAD_SOM_MATERIAL = True
RELINAD_MATERIAL_NAMN = "Relinad"

# Diagrammen bäddas alltid in i Excel-fliken Sammanfattning. Sätt SPARA_DIAGRAM = True
# (eller kör med --diagram) om du dessutom vill ha dem som PNG-filer i <utdata>/diagram
# för t.ex. PowerPoint.
SPARA_DIAGRAM = False

# Kartunderlag: en JSON-fil per körning med en post per sträcka (brunnspar, klass, index,
# material m.m.) som ArcMap-skriptet i arcmap/ läser för att skapa ett ledningslager.
SKRIV_KARTUNDERLAG = True
KARTUNDERLAG_FIL = "kartunderlag.json"


# ----------------------------------------------------------------------------
# Datamodell
# ----------------------------------------------------------------------------

@dataclass
class Observation:
    fil: str
    nr: int
    lage: float            # m från start
    tid: str
    lopande: str           # A1/B1 …
    kod: str               # SPR/YTS/… (skadekod)
    grad: int | None
    infokod: str           # AS/RB/… (ej skada)
    attribut: str
    klocka_fran: str
    klocka_till: str
    vattenniva: str
    bild: str
    kommentar: str
    bild_b: str = ""                     # andra bilden (kolumn 16)
    lopande_langd: float | None = None   # sätts för A-rader
    bild_sokvag: str | None = None       # hittad sökväg till bildfilen
    bild_b_sokvag: str | None = None

    @property
    def bilder(self) -> list[tuple[str, str | None]]:
        return [(n, p) for n, p in ((self.bild, self.bild_sokvag), (self.bild_b, self.bild_b_sokvag)) if n]

    @property
    def ar_skada(self) -> bool:
        return bool(self.kod) and KODER.get(self.kod, ("", "I"))[1] in ("K", "D")

    @property
    def raknas(self) -> bool:
        """Räknas i poängen: skadekod med grad, och inte en B-slutmarkering."""
        return self.ar_skada and self.grad is not None and not self.lopande.startswith("B")

    @property
    def typ(self) -> str:
        return KODER.get(self.kod, ("", ""))[1] if self.kod else ""

    @property
    def poang(self) -> float:
        if not self.raknas:
            return 0.0
        p = GRADPOANG.get(self.grad, 0)
        return p * (DRIFTFAKTOR if self.typ == "D" else 1.0)

    def beskrivning(self) -> str:
        if self.kod:
            namn = KODER.get(self.kod, (self.kod, ""))[0]
            delar = [namn]
            if self.attribut:
                delar.append(ATTRIBUT.get(self.attribut, self.attribut))
            if self.grad:
                delar.append(f"grad {self.grad}")
            if self.lopande:
                delar.append("löpande " + ("start" if self.lopande[0] == "A" else "slut"))
            s = ", ".join(delar)
        elif self.infokod:
            namn = INFOKODER.get(self.infokod, self.infokod)
            s = namn + (f" ({ATTRIBUT.get(self.attribut, self.attribut)})" if self.attribut else "")
            if self.infokod in ("AS", "AG"):
                sida = anslutning_sida(self.klocka_fran)
                if sida:
                    s += f", {sida}"
        else:
            s = "Start/slut"
        if self.kommentar:
            s += f" – {self.kommentar}"
        return s


@dataclass
class Stracka:
    fil: str
    nr: int
    startbrunn: str
    slutbrunn: str
    utgangsbrunn: str
    agare: str
    omrade: str
    projekt: str
    riktning: str
    datum: str
    klockslag: str
    operator: str
    videofil: str
    form: str
    dimension: str
    dimension2: str
    material: str
    foder: str
    fodermaterial: str
    ledningstyp: str
    vader: str
    observationer: list[Observation] = field(default_factory=list)
    profil: list[tuple[float, float, float]] = field(default_factory=list)  # (m, rel z, z)
    profil_start_z: float | None = None
    profil_slut_z: float | None = None
    tv3_sokvag: str = ""                 # absolut sökväg till TV3-filen
    rapport_fil: str | None = None       # relativ sökväg (från utdatakatalogen) till PDF-rapporten
    media_kataloger: list[str] = field(default_factory=list)   # extra kataloger för just denna fil
    video_sokvag: str | None = None      # hittad sökväg till videofilen

    # ---- härledda värden ----
    @property
    def id(self) -> str:
        return f"{self.startbrunn} → {self.slutbrunn}"

    @property
    def langd(self) -> float:
        return max((o.lage for o in self.observationer), default=0.0)

    def skador(self, typ: str | None = None) -> list[Observation]:
        return [o for o in self.observationer if o.raknas and (typ is None or o.typ == typ)]

    def maxgrad(self, typ: str) -> int:
        return max((o.grad for o in self.skador(typ)), default=0)

    def poang(self, typ: str | None = None) -> float:
        return sum(o.poang for o in self.skador(typ))

    def index(self, typ: str | None = None) -> float:
        """Poäng per 100 m."""
        L = self.langd
        return self.poang(typ) / max(L, MINLANGD) * 100 if L > 0 else 0.0

    @property
    def avbruten(self) -> bool:
        return any(o.kod == "KAM" and o.attribut == "HINDE" for o in self.observationer)

    @property
    def relinad(self) -> bool:
        return bool(self.foder.strip()) or any(
            "relin" in o.kommentar.lower() for o in self.observationer)

    @property
    def material_grupp(self) -> str:
        """Material som det redovisas i statistiken – relinade rör får en egen grupp."""
        if RELINAD_SOM_MATERIAL and self.relinad:
            return RELINAD_MATERIAL_NAMN
        return self.material

    @property
    def klass(self) -> str:
        if self.langd < 1:
            return "E"
        kmax, kidx = self.maxgrad("K"), self.index("K")
        if kmax >= 4 or kidx >= TROSKEL_A:
            return "A"
        if kmax >= 3 or kidx >= TROSKEL_B:
            return "B"
        if self.skador():
            return "C"
        return "D"

    @property
    def driftatgard(self) -> str:
        d = self.skador("D")
        if not d:
            return ""
        vals = []
        if any(o.kod == "ROT" and o.grad >= 3 for o in d):
            vals.append("Rotskärning")
        elif any(o.kod == "ROT" and o.grad == 2 for o in d):
            vals.append("Bevaka rötter")
        if any(o.kod in ("SED", "UTF") and o.grad >= 3 for o in d):
            vals.append("Spolning")
        if any(o.kod == "INL" and o.grad >= 3 for o in d):
            vals.append("Täta inläckage")
        if any(o.kod == "INH" and o.grad >= 3 for o in d):
            vals.append("Ta bort inträngande hinder")
        return ", ".join(vals)

    flerinspekterad: bool = False   # samma brunnspar förekommer flera gånger (t.ex. från båda håll)

    def _profil_i_flodesriktning(self) -> list[tuple[float, float]] | None:
        """Profilen (position, höjd) ordnad i flödesriktningen uppströms → nedströms,
        glesad till ca 0,25 m mellan punkterna för att dämpa mätbrus."""
        if len(self.profil) < 3:
            return None
        pts = [(x, z) for x, _, z in self.profil]
        if self.fran_brunn != self.startbrunn:      # kameran gick motströms
            pts = pts[::-1]
        ut = [pts[0]]
        for x, z in pts[1:]:
            if abs(x - ut[-1][0]) >= 0.25:
                ut.append((x, z))
        return ut if len(ut) >= 3 else None

    @property
    def profil_analys(self) -> dict | None:
        """Svackor och bakfall ur inklinometerprofilen.
        svackdjup  – största vattendjup (m) som blir stående i en svacka (punkt lägre än
                     både uppströms och nedströms kant) – 'fill'-metoden i flödesriktningen
        svacklangd – total längd (m) där stående vatten > 1 cm
        svackpos   – position (m från kamerans start) för djupaste punkten
        bakfall    – total längd (m) med lutning mot flödesriktningen (> 5 ‰ över minst 1 m)
        osaker     – True om profilen verkar opålitlig (starthöjd saknas, eller inklinometerns
                     fall avviker kraftigt från brunnshöjderna i PROFILADM)"""
        pts = self._profil_i_flodesriktning()
        if pts is None:
            return None
        n = len(pts)
        z = [p[1] for p in pts]
        upp = [0.0] * n
        ned = [0.0] * n
        m = -1e9
        for i in range(n):
            m = max(m, z[i]); upp[i] = m
        m = -1e9
        for i in range(n - 1, -1, -1):
            m = max(m, z[i]); ned[i] = m
        djup = [max(0.0, min(upp[i], ned[i]) - z[i]) for i in range(n)]
        i_max = max(range(n), key=lambda i: djup[i])
        svacklangd = sum(abs(pts[i + 1][0] - pts[i][0]) for i in range(n - 1) if djup[i] > 0.01 or djup[i + 1] > 0.01)
        bakfall = 0.0
        for i in range(n - 1):
            dx = abs(pts[i + 1][0] - pts[i][0])
            if dx > 0 and (z[i + 1] - z[i]) / dx > 0.005:      # stiger i flödesriktningen
                bakfall += dx
        osaker = self.profil_start_z is None or self.profil_slut_z is None
        if not osaker:
            fall_brunnar = self.profil_start_z - self.profil_slut_z
            fall_inkl = z[0] - z[-1]
            if abs(fall_inkl - fall_brunnar) > max(0.3, 0.5 * abs(fall_brunnar)):
                osaker = True
        return {"svackdjup": djup[i_max], "svacklangd": svacklangd, "svackpos": pts[i_max][0],
                "bakfall": bakfall, "osaker": osaker}

    @property
    def svacka_andel(self) -> float | None:
        """Svackdjup i förhållande till rördiametern (0–1)."""
        a = self.profil_analys
        try:
            d = float(self.dimension) / 1000
        except ValueError:
            return None
        return round(a["svackdjup"] / d, 2) if a and d > 0 else None

    @property
    def max_svacka(self) -> float | None:
        a = self.profil_analys
        return a["svackdjup"] if a else None

    @property
    def lutning_promille(self) -> float | None:
        if self.profil_start_z is None or self.profil_slut_z is None or self.langd <= 0:
            return None
        return (self.profil_start_z - self.profil_slut_z) / self.langd * 1000

    @property
    def fran_brunn(self) -> str:
        """Brunnen där kameran startade (position 0 m)."""
        return self.utgangsbrunn or self.startbrunn

    @property
    def till_brunn(self) -> str:
        return self.startbrunn if self.fran_brunn == self.slutbrunn else self.slutbrunn

    @property
    def antal_anslutningar(self) -> int:
        """Antal registrerade anslutningar (AS/AG) på sträckan."""
        return sum(1 for o in self.observationer if o.infokod in ("AS", "AG"))

    def sammanfattning_skador(self) -> str:
        c = Counter()
        for o in self.skador():
            c[(o.kod, o.grad)] += 1
        delar = []
        for (kod, grad), n in sorted(c.items(), key=lambda kv: (-kv[0][1], kv[0][0])):
            delar.append(f"{n}×{kod}{grad}")
        return ", ".join(delar)


# ----------------------------------------------------------------------------
# Parser
# ----------------------------------------------------------------------------

def _f(s: str) -> float | None:
    s = s.strip().replace(",", ".")
    if not s:
        return None
    try:
        return float(s)
    except ValueError:
        return None


def _i(s: str) -> int | None:
    s = s.strip()
    return int(s) if s.isdigit() else None


def las_text(path: str) -> str:
    raw = open(path, "rb").read()
    for enc in ("utf-8-sig", "cp1252", "latin-1"):
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            continue
    return raw.decode("latin-1", errors="replace")


def las_tv3(path: str) -> list[Stracka]:
    text = las_text(path)
    filnamn = os.path.basename(path)
    sektioner: dict[str, list[list[str]]] = defaultdict(list)
    aktuell = None
    for rad in text.splitlines():
        if rad.startswith("#"):
            aktuell = rad[1:].strip().split("=")[0].upper()
            continue
        if aktuell and rad.strip():
            sektioner[aktuell].append(rad.split(";"))

    if "TVADM" not in sektioner:
        raise ValueError(f"{filnamn}: ingen #TVADM-sektion hittades – är det en TV3-fil?")

    def g(r, i):
        return r[i].strip() if i < len(r) else ""

    strackor: dict[int, Stracka] = {}
    for r in sektioner["TVADM"]:
        nr = _i(g(r, 0))
        if nr is None:
            continue
        strackor[nr] = Stracka(
            fil=filnamn, nr=nr,
            startbrunn=g(r, 2), slutbrunn=g(r, 3), utgangsbrunn=g(r, 4),
            agare=g(r, 7), omrade=g(r, 8), projekt=g(r, 9), riktning=g(r, 10),
            datum=g(r, 12), klockslag=g(r, 13), operator=g(r, 14), videofil=g(r, 18),
            form=g(r, 22), dimension=g(r, 23), dimension2=g(r, 24), material=g(r, 25),
            foder=g(r, 26), fodermaterial=g(r, 27), ledningstyp=g(r, 29), vader=g(r, 31),
            tv3_sokvag=os.path.abspath(path),
        )

    for r in sektioner.get("TVDAT", []):
        nr = _i(g(r, 0))
        if nr is None or nr not in strackor:
            continue
        o = Observation(
            fil=filnamn, nr=nr, lage=_f(g(r, 1)) or 0.0, tid=g(r, 2), lopande=g(r, 3),
            kod=g(r, 4).upper(), grad=_i(g(r, 5)), infokod=g(r, 6).upper(), attribut=g(r, 7).upper(),
            klocka_fran=g(r, 11), klocka_till=g(r, 12), vattenniva=g(r, 13),
            bild=g(r, 14), bild_b=g(r, 15), kommentar=g(r, 17),
        )
        strackor[nr].observationer.append(o)

    # Längd på löpande skador (A1 … B1)
    for s in strackor.values():
        start: dict[str, Observation] = {}
        for o in s.observationer:
            if o.lopande.startswith("A"):
                start[o.lopande[1:]] = o
            elif o.lopande.startswith("B"):
                a = start.pop(o.lopande[1:], None)
                if a:
                    a.lopande_langd = round(o.lage - a.lage, 2)

    for r in sektioner.get("PROFILADM", []):
        nr = _i(g(r, 0))
        if nr in strackor:
            strackor[nr].profil_start_z = _f(g(r, 16))
            strackor[nr].profil_slut_z = _f(g(r, 17))
    for r in sektioner.get("PROFILDAT", []):
        nr = _i(g(r, 0))
        if nr in strackor:
            x, rz, z = _f(g(r, 1)), _f(g(r, 2)), _f(g(r, 3))
            if x is not None and z is not None:
                strackor[nr].profil.append((x, rz or 0.0, z))

    for s in strackor.values():
        s.profil.sort(key=lambda p: p[0])

    par = Counter(frozenset((s.startbrunn, s.slutbrunn)) for s in strackor.values())
    for s in strackor.values():
        s.flerinspekterad = par[frozenset((s.startbrunn, s.slutbrunn))] > 1

    return [strackor[k] for k in sorted(strackor)]


# ----------------------------------------------------------------------------
# Media (video + bilder)
# ----------------------------------------------------------------------------

MEDIA_ANDELSER = {".mp4", ".mpg", ".mpeg", ".avi", ".wmv", ".mov", ".mkv", ".jpg", ".jpeg", ".png", ".bmp"}
_media_cache: dict[str, dict[str, str]] = {}


def indexera_katalog(katalog: str) -> dict[str, str]:
    """Filnamn (gemener) -> absolut sökväg för alla media-filer under katalogen (rekursivt)."""
    katalog = os.path.abspath(katalog)
    if katalog in _media_cache:
        return _media_cache[katalog]
    idx: dict[str, str] = {}
    if os.path.isdir(katalog):
        for rot, _, namn in os.walk(katalog):
            for n in namn:
                if os.path.splitext(n)[1].lower() in MEDIA_ANDELSER:
                    idx.setdefault(n.lower(), os.path.join(rot, n))   # första träffen vinner
    _media_cache[katalog] = idx
    return idx


def koppla_media(strackor: list[Stracka], extra_kataloger: list[str]) -> tuple[int, int, int, int]:
    """Letar upp video- och bildfiler. Söker i TV3-filens katalog först, sedan i filens egna
    mediakataloger (från listfilen) och sist i de globala (media: i listfilen eller --media).
    Returnerar (videor hittade, videor totalt, bilder hittade, bilder totalt)."""
    vh = vt = bh = bt = 0
    for s in strackor:
        kataloger = [os.path.dirname(s.tv3_sokvag)] + list(s.media_kataloger) + list(extra_kataloger)
        index = [indexera_katalog(k) for k in kataloger]

        def hitta(namn: str) -> str | None:
            n = namn.strip().lower()
            if not n:
                return None
            for idx in index:
                if n in idx:
                    return idx[n]
            return None

        if s.videofil:
            vt += 1
            s.video_sokvag = hitta(s.videofil)
            vh += s.video_sokvag is not None
        for o in s.observationer:
            if o.bild:
                bt += 1
                o.bild_sokvag = hitta(o.bild)
                bh += o.bild_sokvag is not None
            if o.bild_b:
                bt += 1
                o.bild_b_sokvag = hitta(o.bild_b)
                bh += o.bild_b_sokvag is not None
    return vh, vt, bh, bt


def fil_url(sokvag: str) -> str:
    """Absolut sökväg -> file:///-URL som Excel kan öppna."""
    from urllib.parse import quote
    p = os.path.abspath(sokvag).replace("\\", "/")
    if not p.startswith("/"):
        p = "/" + p                     # C:/... -> /C:/...
    return "file://" + quote(p, safe="/:")


# ----------------------------------------------------------------------------
# Excel
# ----------------------------------------------------------------------------

def skriv_excel(strackor: list[Stracka], path: str, diagram: dict[str, str], topp: int):
    from openpyxl import Workbook
    from openpyxl.drawing.image import Image as XLImage
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter

    wb = Workbook()
    rubrik = Font(bold=True, color="FFFFFF")
    rubrikfyll = PatternFill("solid", fgColor="1F3864")
    tunn = Side(style="thin", color="D9D9D9")
    kant = Border(top=tunn, bottom=tunn, left=tunn, right=tunn)

    lank = Font(color="0563C1", underline="single")

    def tabell(ws, kolumner, rader, bredder=None, klasskol=None, lankar=None):
        """lankar: {kolumnindex: [url eller None per rad]} – gör cellerna klickbara."""
        ws.append(kolumner)
        for c in ws[1]:
            c.font, c.fill, c.border = rubrik, rubrikfyll, kant
            c.alignment = Alignment(vertical="center", wrap_text=True)
        for rad in rader:
            ws.append(rad)
        for ri, row in enumerate(ws.iter_rows(min_row=2, max_row=ws.max_row)):
            for c in row:
                c.border = kant
            for ci, urls in (lankar or {}).items():
                if ri < len(urls) and urls[ri] and row[ci].value:
                    row[ci].hyperlink = urls[ri]
                    row[ci].font = lank
            if klasskol is not None:
                k = row[klasskol].value
                if k and k[0] in KLASS_FARG:
                    row[klasskol].fill = PatternFill("solid", fgColor=KLASS_FARG[k[0]])
                    row[klasskol].font = Font(bold=True, color="FFFFFF" if k[0] in "AB" else "000000")
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
        for i, kol in enumerate(kolumner, 1):
            b = (bredder or {}).get(kol)
            if b is None:
                b = min(45, max(10, max(len(str(r[i - 1])) if i - 1 < len(r) and r[i - 1] is not None else 0
                                        for r in rader[:200]) + 2, len(kol) + 2))
            ws.column_dimensions[get_column_letter(i)].width = b

    # ---- Sammanfattning ----
    ws = wb.active
    ws.title = "Sammanfattning"
    tot_m = sum(s.langd for s in strackor)
    klasser = Counter(s.klass for s in strackor)
    klass_m = defaultdict(float)
    for s in strackor:
        klass_m[s.klass] += s.langd
    filer = sorted({s.fil for s in strackor})
    projekt = sorted({s.projekt for s in strackor if s.projekt})
    datum = sorted({s.datum for s in strackor if s.datum})

    rader = [
        ["Analys av TV-inspektion", ""],
        ["Filer", ", ".join(filer)],
        ["Projekt", ", ".join(projekt)],
        ["Inspektionsperiod", f"{datum[0]} – {datum[-1]}" if datum else ""],
        ["Antal sträckor", len(strackor)],
        ["Inspekterad längd (m)", round(tot_m)],
        ["Antal skadeobservationer", sum(len(s.skador()) for s in strackor)],
        ["Sträckor med avbruten inspektion (hinder)", sum(1 for s in strackor if s.avbruten)],
        ["Relinade sträckor", sum(1 for s in strackor if s.relinad)],
        ["", ""],
        ["Prioritetsklass", "Antal sträckor", "Andel sträckor", "Längd (m)", "Andel längd"],
    ]
    for k in "ABCDE":
        rader.append([KLASS_TEXT[k], klasser.get(k, 0),
                      klasser.get(k, 0) / len(strackor) if strackor else 0,
                      round(klass_m.get(k, 0)), klass_m.get(k, 0) / tot_m if tot_m else 0])
    rader += [["", ""], ["Poängmodell", ""],
              ["Grad 1 / 2 / 3 / 4", " / ".join(f"{GRADPOANG[g]} p" for g in (1, 2, 3, 4))],
              ["Konstruktionskoder (faktor 1,0)", ", ".join(k for k, v in KODER.items() if v[1] == "K")],
              ["Driftkoder (faktor %s)" % str(DRIFTFAKTOR).replace(".", ","),
               ", ".join(k for k, v in KODER.items() if v[1] == "D")],
              ["Index", "poäng per 100 m ledning"],
              ["Klass A", f"konstruktionsgrad 4 eller konstruktionsindex ≥ {TROSKEL_A:g}"],
              ["Klass B", f"konstruktionsgrad 3 eller konstruktionsindex ≥ {TROSKEL_B:g}"],
              ["Klass C", "övriga sträckor med skador"],
              ["Klass D", "inga skadeobservationer"],
              ["Klass E", "ej bedömd (ingen inspekterad längd)"],
              ["Kort sträcka", f"sträckor < {MINLANGD:g} m normeras som {MINLANGD:g} m"]]
    for r in rader:
        ws.append(r)
    ws["A1"].font = Font(bold=True, size=14)
    for r in (11, 22):
        for c in ws[r]:
            c.font = Font(bold=True)
    for r in range(12, 17):
        ws.cell(r, 3).number_format = "0%"
        ws.cell(r, 5).number_format = "0%"
        ws.cell(r, 1).fill = PatternFill("solid", fgColor=KLASS_FARG[ws.cell(r, 1).value[0]])
        ws.cell(r, 1).font = Font(bold=True, color="FFFFFF" if ws.cell(r, 1).value[0] in "AB" else "000000")
    ws.column_dimensions["A"].width = 42
    ws.column_dimensions["B"].width = 40
    for col in "CDE":
        ws.column_dimensions[col].width = 14
    rad = 34
    for namn in ("klasser", "topp"):
        if namn in diagram and os.path.exists(diagram[namn]):
            img = XLImage(diagram[namn])
            img.width, img.height = img.width * 0.55, img.height * 0.55
            ws.add_image(img, f"A{rad}")
            rad += int(img.height / 20) + 3

    # ---- Prioritering ----
    ws = wb.create_sheet("Prioritering")
    kol = ["Rang", "Prioritetsklass", "Fil", "Nr", "Startbrunn", "Slutbrunn", "Område", "Ledningstyp",
           "Material", "Dim (mm)", "Längd (m)", "Datum",
           "Totalindex (p/100 m)", "Konstruktionsindex (p/100 m)", "Driftindex (p/100 m)",
           "Konstr. maxgrad", "Drift maxgrad", "Antal skador", "Antal anslutningar", "Skador (kod+grad)",
           "Driftåtgärd", "Avbruten inspektion", "Inspekterad flera ggr", "Relinad",
           "Svackdjup (m)", "Svackdjup/diameter", "Svacklängd (m)", "Bakfall längd (m)", "Lutning (‰)", "Profil osäker",
           "Rapport", "Videofil"]
    sorterade = sorterade_strackor(strackor)
    rader = []
    for rang, s in enumerate(sorterade, 1):
        pa = s.profil_analys
        lut = s.lutning_promille
        rader.append([rang, KLASS_TEXT[s.klass], s.fil, s.nr, s.startbrunn, s.slutbrunn, s.omrade,
                      s.ledningstyp.capitalize(), s.material, s.dimension + (f"/{s.dimension2}" if s.dimension2 else ""),
                      round(s.langd, 1), s.datum, round(s.index(), 1), round(s.index("K"), 1),
                      round(s.index("D"), 1), s.maxgrad("K") or None, s.maxgrad("D") or None,
                      len(s.skador()), s.antal_anslutningar, s.sammanfattning_skador(), s.driftatgard,
                      "Ja" if s.avbruten else "", "Ja" if s.flerinspekterad else "", "Ja" if s.relinad else "",
                      round(pa["svackdjup"], 2) if pa else None,
                      s.svacka_andel,
                      round(pa["svacklangd"], 1) if pa else None,
                      round(pa["bakfall"], 1) if pa else None,
                      round(lut, 1) if lut is not None else None,
                      ("Ja" if pa["osaker"] else "") if pa else "",
                      "Öppna rapport" if s.rapport_fil else "", s.videofil])
    video_urls = [fil_url(s.video_sokvag) if s.video_sokvag else None for s in sorterade]
    rapport_urls = [s.rapport_fil.replace("\\", "/") if s.rapport_fil else None for s in sorterade]   # relativ länk
    tabell(ws, kol, rader, {"Skador (kod+grad)": 45, "Prioritetsklass": 24, "Driftåtgärd": 28, "Rapport": 15},
           klasskol=1, lankar={len(kol) - 1: video_urls, len(kol) - 2: rapport_urls})
    ci = kol.index("Svackdjup/diameter") + 1
    for r in range(2, ws.max_row + 1):
        ws.cell(r, ci).number_format = "0%"

    # ---- Observationer ----
    ws = wb.create_sheet("Observationer")
    kol = ["Fil", "Nr", "Startbrunn", "Slutbrunn", "Prioritetsklass", "Läge (m)", "Tid i film", "Kod",
           "Beskrivning", "Typ", "Grad", "Poäng", "Löpande", "Löpande längd (m)",
           "Klocka från", "Klocka till", "Vattennivå (%)", "Bild", "Kommentar", "Videofil"]
    rader, bild_urls, video_urls = [], [], []
    for s in sorterade:
        for o in s.observationer:
            if not (o.kod or o.infokod):
                continue
            bs = o.bild_sokvag or (o.bild_b_sokvag if not o.bild else None)
            bild_urls.append(fil_url(bs) if bs else None)
            video_urls.append(fil_url(s.video_sokvag) if s.video_sokvag else None)
            rader.append([s.fil, s.nr, s.startbrunn, s.slutbrunn, KLASS_TEXT[s.klass], o.lage, o.tid,
                          o.kod or o.infokod, o.beskrivning(),
                          {"K": "Konstruktion", "D": "Drift", "I": "Info"}.get(o.typ, "Info"),
                          o.grad, o.poang or None, o.lopande, o.lopande_langd,
                          o.klocka_fran, o.klocka_till, o.vattenniva, o.bild or o.bild_b, o.kommentar, s.videofil])
    tabell(ws, kol, rader, {"Beskrivning": 50, "Kommentar": 30, "Prioritetsklass": 24}, klasskol=4,
           lankar={kol.index("Bild"): bild_urls, kol.index("Videofil"): video_urls})

    # ---- Kodstatistik ----
    ws = wb.create_sheet("Kodstatistik")
    c = Counter()
    for s in strackor:
        for o in s.skador():
            c[(o.kod, o.grad)] += 1
    kod_tot = Counter()
    for (k, g), n in c.items():
        kod_tot[k] += n
    kol = ["Kod", "Beskrivning", "Typ", "Grad 1", "Grad 2", "Grad 3", "Grad 4", "Totalt", "Sträckor berörda"]
    rader = []
    for k, n in kod_tot.most_common():
        rader.append([k, KODER.get(k, (k, ""))[0], "Konstruktion" if KODER.get(k, ("", ""))[1] == "K" else "Drift",
                      c.get((k, 1), 0), c.get((k, 2), 0), c.get((k, 3), 0), c.get((k, 4), 0), n,
                      sum(1 for s in strackor if any(o.kod == k for o in s.skador()))])
    tabell(ws, kol, rader)

    # ---- Per fil ----
    ws = wb.create_sheet("Per fil")
    grpf = defaultdict(lambda: {"n": 0, "m": 0.0, "p": 0.0, "A": 0, "B": 0, "C": 0, "D": 0, "E": 0,
                                "skador": 0, "avbr": 0, "projekt": set(), "omr": set(), "datum": []})
    for s in strackor:
        g = grpf[s.fil]
        g["n"] += 1
        g["m"] += s.langd
        g["p"] += s.poang()
        g[s.klass] += 1
        g["skador"] += len(s.skador())
        g["avbr"] += s.avbruten
        if s.projekt:
            g["projekt"].add(s.projekt)
        if s.omrade:
            g["omr"].add(s.omrade)
        if s.datum:
            g["datum"].append(s.datum)
    kol = ["Fil", "Projekt", "Område", "Period", "Sträckor", "Längd (m)", "Skadeobservationer",
           "Index (p/100 m)", "Klass A", "Klass B", "Klass C", "Klass D", "Klass E",
           "Andel A+B (sträckor)", "Längd A+B (m)", "Avbrutna inspektioner"]
    rader = []
    for fil, g in sorted(grpf.items()):
        ab_m = sum(s.langd for s in strackor if s.fil == fil and s.klass in "AB")
        d = sorted(g["datum"])
        rader.append([fil, ", ".join(sorted(g["projekt"])), ", ".join(sorted(g["omr"])),
                      f"{d[0]} – {d[-1]}" if d else "", g["n"], round(g["m"]), g["skador"],
                      round(g["p"] / g["m"] * 100, 1) if g["m"] else 0,
                      g["A"], g["B"], g["C"], g["D"], g["E"],
                      (g["A"] + g["B"]) / g["n"] if g["n"] else 0, round(ab_m), g["avbr"]])
    tabell(ws, kol, rader, {"Område": 30})
    for r in range(2, ws.max_row + 1):
        ws.cell(r, 14).number_format = "0%"

    # ---- Material & dimension ----
    ws = wb.create_sheet("Material")
    grp = defaultdict(lambda: {"n": 0, "m": 0.0, "p": 0.0, "A": 0, "B": 0, "urspr": Counter()})
    for s in strackor:
        g = grp[(s.material_grupp, s.ledningstyp.capitalize())]
        g["n"] += 1
        g["m"] += s.langd
        g["p"] += s.poang()
        g["urspr"][s.material or "Okänt"] += 1
        if s.klass in "AB":
            g[s.klass] += 1

    def ursprung(mat, g):
        """Ursprungsmaterial för relinade rör, t.ex. 'Betong (16)'. Tomt för övriga."""
        if not RELINAD_SOM_MATERIAL or mat != RELINAD_MATERIAL_NAMN:
            return ""
        return ", ".join(f"{m} ({n})" for m, n in g["urspr"].most_common())

    kol = ["Material", "Ursprungsmaterial", "Ledningstyp", "Sträckor", "Längd (m)",
           "Index (p/100 m)", "Klass A", "Klass B", "Andel A+B"]
    rader = [[m, ursprung(m, g), t, g["n"], round(g["m"]), round(g["p"] / g["m"] * 100, 1) if g["m"] else 0,
              g["A"], g["B"], (g["A"] + g["B"]) / g["n"]] for (m, t), g in sorted(grp.items())]
    tabell(ws, kol, rader)
    for r in range(2, ws.max_row + 1):
        ws.cell(r, 9).number_format = "0%"

    wb.save(path)


def skriv_kartunderlag(strackor: list[Stracka], path: str) -> int:
    """Skriver en JSON-fil med en post per sträcka, avsedd för kartframställning.

    Varje post identifierar sträckan med brunnsparet (startbrunn/slutbrunn) så att
    ArcMap-skriptet kan leta upp ledningen mellan brunnarna. Maskinell bedömning =
    prioritetsklassen från poängmodellen; den manuella bedömningen fylls i i kartan."""
    import json

    poster = []
    for rang, s in enumerate(sorterade_strackor(strackor), 1):
        pa = s.profil_analys
        lut = s.lutning_promille
        poster.append({
            "rang": rang,
            "fil": s.fil,
            "nr": s.nr,
            "startbrunn": s.startbrunn,          # uppströms
            "slutbrunn": s.slutbrunn,            # nedströms
            "utgangsbrunn": s.fran_brunn,        # där kameran startade
            "omrade": s.omrade,
            "datum": s.datum,
            "maskinell_bedomning": s.klass,
            "klasstext": KLASS_TEXT[s.klass],
            "totalindex": round(s.index(), 1),
            "konstruktionsindex": round(s.index("K"), 1),
            "driftindex": round(s.index("D"), 1),
            "maxgrad_konstruktion": s.maxgrad("K"),
            "maxgrad_drift": s.maxgrad("D"),
            "antal_skador": len(s.skador()),
            "antal_anslutningar": s.antal_anslutningar,
            "skador": s.sammanfattning_skador(),
            "driftatgard": s.driftatgard,
            "langd_m": round(s.langd, 1),
            "material": s.material,
            "material_grupp": s.material_grupp,
            "dimension": s.dimension,
            "ledningstyp": s.ledningstyp.capitalize(),
            "relinad": s.relinad,
            "avbruten": s.avbruten,
            "flerinspekterad": s.flerinspekterad,
            "svackdjup_m": round(pa["svackdjup"], 2) if pa else None,
            "svacklangd_m": round(pa["svacklangd"], 1) if pa else None,
            "bakfall_m": round(pa["bakfall"], 1) if pa else None,
            "lutning_promille": round(lut, 1) if lut is not None else None,
            "profil_osaker": bool(pa["osaker"]) if pa else None,
            "tv3_fil": s.tv3_sokvag,
            "rapport": s.rapport_fil.replace("\\", "/") if s.rapport_fil else None,
            "videofil": s.videofil,
            "video_sokvag": s.video_sokvag,
        })

    data = {
        "version": 1,
        "kalla": "tv3_analys.py",
        "genererad": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "klasser": KLASS_TEXT,
        "strackor": poster,
    }
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=1)
    return len(poster)


def sorterade_strackor(strackor):
    ordn = {"A": 0, "B": 1, "C": 2, "D": 3, "E": 4}
    return sorted(strackor, key=lambda s: (ordn[s.klass], -s.index(), -s.maxgrad("K"), s.fil, s.nr))


# ----------------------------------------------------------------------------
# Diagram
# ----------------------------------------------------------------------------

def rita_diagram(strackor: list[Stracka], katalog: str, topp: int) -> dict[str, str]:
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt

    os.makedirs(katalog, exist_ok=True)
    plt.rcParams.update({
        "font.family": "DejaVu Sans", "font.size": 11, "axes.spines.top": False,
        "axes.spines.right": False, "axes.spines.left": False, "axes.edgecolor": "#c3c2b7",
        "axes.grid": True, "axes.grid.axis": "y", "grid.color": "#e1e0d9", "grid.linewidth": 0.8,
        "axes.axisbelow": True, "xtick.color": "#898781", "ytick.color": "#898781",
        "axes.labelcolor": "#52514e", "axes.titlecolor": "#0b0b0b", "axes.titleweight": "bold",
        "axes.titlesize": 13, "axes.titlelocation": "left", "figure.facecolor": "white",
        "axes.facecolor": "white", "ytick.left": False, "xtick.bottom": False,
    })
    ut = {}
    tot_m = sum(s.langd for s in strackor) or 1

    # 1. Fördelning per prioritetsklass (längd)
    klass_m = defaultdict(float)
    klass_n = Counter()
    for s in strackor:
        klass_m[s.klass] += s.langd
        klass_n[s.klass] += 1
    fig, ax = plt.subplots(figsize=(8, 4.2))
    ks = [k for k in "ABCDE" if klass_n[k]]
    vals = [klass_m[k] for k in ks]
    bars = ax.bar([KLASS_TEXT[k].replace(" – ", "\n") for k in ks], vals,
                  color=[KLASS_FARG_HEX[k] for k in ks], width=0.6)
    for b, k in zip(bars, ks):
        ax.text(b.get_x() + b.get_width() / 2, b.get_height() + tot_m * 0.01,
                f"{klass_m[k]:.0f} m · {klass_m[k] / tot_m:.0%}\n{klass_n[k]} sträckor",
                ha="center", va="bottom", fontsize=10, color="#0b0b0b")
    ax.set_title("Inspekterad ledningslängd per prioritetsklass")
    ax.set_ylabel("meter")
    ax.set_ylim(0, max(vals) * 1.3 if max(vals) else 1)
    fig.tight_layout()
    ut["klasser"] = os.path.join(katalog, "1_prioritetsklasser.png")
    fig.savefig(ut["klasser"], dpi=200)
    plt.close(fig)

    # 2. Topp-N mest kritiska sträckor
    top = sorterade_strackor(strackor)[:topp]
    fig, ax = plt.subplots(figsize=(10, 0.42 * len(top) + 1.6))
    flera_projekt = len({s.projekt for s in strackor}) > 1
    etik = [(f"{s.projekt}: " if flera_projekt and s.projekt else "")
            + f"{s.startbrunn} → {s.slutbrunn}  ({s.material} {s.dimension}, {s.langd:.0f} m)" for s in top][::-1]
    k_idx = [s.index("K") for s in top][::-1]
    d_idx = [s.index("D") for s in top][::-1]
    ax.barh(etik, k_idx, color="#2a78d6", height=0.62, label="Konstruktion (SPR, RBR, DEF, YTS, FOG)")
    ax.barh(etik, d_idx, left=k_idx, color="#eb6834", height=0.62, label="Drift (ROT, INL, SED, UTF, INH)")
    for i, s in enumerate(top[::-1]):
        ax.text(k_idx[i] + d_idx[i] + max(k_idx) * 0.01, i, f"{s.index():.0f}  ·  {KLASS_TEXT[s.klass][:1]}",
                va="center", fontsize=9, color="#52514e")
    ax.grid(axis="y", visible=False)
    ax.grid(axis="x", visible=True)
    ax.set_xlabel("skadepoäng per 100 m")
    ax.set_title(f"De {len(top)} mest kritiska sträckorna")
    ax.legend(loc="lower right", frameon=False, fontsize=9)
    fig.tight_layout()
    ut["topp"] = os.path.join(katalog, "2_topplista.png")
    fig.savefig(ut["topp"], dpi=200)
    plt.close(fig)

    # 3. Observationer per kod och grad
    c = Counter()
    for s in strackor:
        for o in s.skador():
            c[(o.kod, o.grad)] += 1
    koder = [k for k, _ in Counter({k: sum(v for (kk, g), v in c.items() if kk == k) for k in KODER}).most_common() if _ > 0]
    fig, ax = plt.subplots(figsize=(10, 5))
    seq = {1: "#cde2fb", 2: "#86b6ef", 3: "#2a78d6", 4: "#104281"}
    import textwrap
    etik = [f"{k}\n" + "\n".join(textwrap.wrap(KODER[k][0], 11)) for k in koder]
    botten = [0] * len(koder)
    for g in (1, 2, 3, 4):
        v = [c.get((k, g), 0) for k in koder]
        ax.bar(etik, v, bottom=botten, color=seq[g], label=f"Grad {g}",
               width=0.65, edgecolor="white", linewidth=1)
        botten = [b + x for b, x in zip(botten, v)]
    for i, b in enumerate(botten):
        ax.text(i, b + max(botten) * 0.01, str(b), ha="center", va="bottom", fontsize=10)
    ax.set_title("Skadeobservationer per kod och grad")
    ax.set_ylabel("antal")
    ax.tick_params(axis="x", labelsize=8.5)
    ax.legend(frameon=False, ncol=4, loc="upper right", fontsize=9)
    fig.tight_layout()
    ut["koder"] = os.path.join(katalog, "3_observationer_per_kod.png")
    fig.savefig(ut["koder"], dpi=200)
    plt.close(fig)

    # 4. Klassfördelning per material (andel längd)
    grp = defaultdict(lambda: defaultdict(float))
    for s in strackor:
        grp[s.material_grupp or "Okänt"][s.klass] += s.langd
    mats = sorted(grp, key=lambda m: -sum(grp[m].values()))
    fig, ax = plt.subplots(figsize=(8, 0.6 * len(mats) + 1.8))
    left = [0.0] * len(mats)
    for k in "ABCD":
        v = [grp[m][k] / sum(grp[m].values()) for m in mats]
        ax.barh([f"{m} ({sum(grp[m].values()):.0f} m)" for m in mats][::-1], v[::-1], left=left[::-1],
                color=KLASS_FARG_HEX[k], label=KLASS_TEXT[k], height=0.6, edgecolor="white", linewidth=1)
        left = [l + x for l, x in zip(left, v)]
    ax.set_xlim(0, 1)
    ax.xaxis.set_major_formatter(matplotlib.ticker.PercentFormatter(1.0))
    ax.grid(axis="y", visible=False)
    ax.grid(axis="x", visible=True)
    ax.set_title("Prioritetsklass per material (andel av inspekterad längd)")
    ax.legend(frameon=False, ncol=2, loc="upper center", bbox_to_anchor=(0.5, -0.15), fontsize=9)
    fig.tight_layout()
    ut["material"] = os.path.join(katalog, "4_klass_per_material.png")
    fig.savefig(ut["material"], dpi=200)
    plt.close(fig)

    return ut



# ----------------------------------------------------------------------------
# PDF-rapport per sträcka
# ----------------------------------------------------------------------------

GRAD_FARG_HEX = {1: "#0ca30c", 2: "#fab219", 3: "#ec835a", 4: "#d03b3b"}
GRAD_FARG_LJUS = {1: "#e6f6e6", 2: "#fff4d6", 3: "#fde6dc", 4: "#f8d7d7"}


def _pdf_typsnitt() -> tuple[str, str]:
    """Registrerar ett TrueType-typsnitt med stöd för åäö och → om ett hittas.
    Returnerar (normal, fet)."""
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    kandidater = [
        ("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf", "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"),
        ("C:/Windows/Fonts/segoeui.ttf", "C:/Windows/Fonts/segoeuib.ttf"),
        ("C:/Windows/Fonts/arial.ttf", "C:/Windows/Fonts/arialbd.ttf"),
        ("/Library/Fonts/Arial.ttf", "/Library/Fonts/Arial Bold.ttf"),
    ]
    for normal, fet in kandidater:
        if os.path.exists(normal) and os.path.exists(fet):
            try:
                pdfmetrics.registerFont(TTFont("Rapport", normal))
                pdfmetrics.registerFont(TTFont("Rapport-Bold", fet))
                return "Rapport", "Rapport-Bold"
            except Exception:
                pass
    return "Helvetica", "Helvetica-Bold"


def anslutning_sida(klocka: str) -> str:
    """Klockposition -> 'vänster', 'höger', 'hjässa', 'botten' eller '' (sett i inspektionsriktningen)."""
    try:
        k = int(klocka)
    except (TypeError, ValueError):
        return ""
    if k in (7, 8, 9, 10, 11):
        return "vänster"
    if k in (1, 2, 3, 4, 5):
        return "höger"
    if k == 12:
        return "hjässa"
    if k == 6:
        return "botten"
    return ""


def rita_schema(s: Stracka, path: str) -> None:
    """Schematisk bild av sträckan: rör med anslutningar, skador (färg efter grad) och löpande skador."""
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    from matplotlib.patches import FancyBboxPatch

    L = max(s.langd, 1.0)
    fig, ax = plt.subplots(figsize=(7.4, 2.5))
    ax.set_xlim(-L * 0.06, L * 1.06)
    ax.set_ylim(-2.15, 2.0)
    ax.axis("off")
    # röret
    ax.add_patch(FancyBboxPatch((0, -0.25), L, 0.5, boxstyle="round,pad=0,rounding_size=0.02",
                                fc="#e8e8e8", ec="#7f7f7f", lw=1.2))
    # brunnar
    for x, namn, ha in ((0, s.fran_brunn, "right"), (L, s.till_brunn, "left")):
        ax.plot(x, 0, "o", ms=14, mfc="#bfbfbf", mec="#4d4d4d", mew=1.2, zorder=5)
        ax.text(x + (-0.02 if ha == "right" else 0.02) * L, 0.55, namn, ha=ha, va="bottom", fontsize=8, fontweight="bold")
    ax.annotate("", xy=(L * 0.12, -1.2), xytext=(L * 0.02, -1.2),
                arrowprops=dict(arrowstyle="->", color="#2a78d6", lw=1.2))
    ax.text(L * 0.13, -1.2, f"inspektionsriktning ({s.riktning.lower()}), position mätt från {s.fran_brunn}",
            va="center", fontsize=7, color="#2a78d6")
    # löpande skador som band
    band = 0
    for o in s.observationer:
        if o.raknas and o.lopande.startswith("A") and o.lopande_langd:
            y = 0.95 + 0.22 * (band % 3)
            ax.plot([o.lage, o.lage + o.lopande_langd], [y, y], lw=4, color=GRAD_FARG_HEX.get(o.grad, "#888"),
                    solid_capstyle="butt", alpha=0.9)
            ax.text(o.lage, y + 0.08, f"{o.kod}{o.grad} {o.lopande_langd:.1f} m", fontsize=6, va="bottom")
            band += 1
    # punktskador
    for o in s.observationer:
        if o.raknas and not (o.lopande.startswith("A") and o.lopande_langd):
            ax.plot(o.lage, 0, "D", ms=7, mfc=GRAD_FARG_HEX.get(o.grad, "#888"), mec="white", mew=0.8, zorder=6)
        elif o.kod == "KAM":
            ax.plot(o.lage, 0, "X", ms=9, mfc="#4d4d4d", mec="white", zorder=6)
    # anslutningar – sida enligt klockposition sett i inspektionsriktningen:
    # kl 7–11 = vänster (ritas ovanför röret), kl 1–5 = höger (under röret), kl 12/6 = hjässa/botten (på röret)
    ovan = under = 0
    for o in s.observationer:
        if o.infokod in ("AS", "AG"):
            sida = anslutning_sida(o.klocka_fran)
            etikett = f"{o.lage:.1f} m" + (f" kl {o.klocka_fran.lstrip('0')}" if o.klocka_fran else "")
            if sida == "vänster":
                y = 0.42 + 0.28 * (ovan % 2); ovan += 1
                ax.plot(o.lage, y, "v", ms=7, mfc="#2a78d6", mec="white", zorder=6)
                ax.text(o.lage, y + 0.13, etikett, ha="center", va="bottom", fontsize=5.5, color="#2a78d6")
            elif sida == "höger":
                y = -0.42 - 0.28 * (under % 2); under += 1
                ax.plot(o.lage, y, "^", ms=7, mfc="#2a78d6", mec="white", zorder=6)
                ax.text(o.lage, y - 0.13, etikett, ha="center", va="top", fontsize=5.5, color="#2a78d6")
            else:
                ax.plot(o.lage, 0, "s", ms=6, mfc="#2a78d6", mec="white", zorder=6)
                ax.text(o.lage, -0.42, etikett, ha="center", va="top", fontsize=5.5, color="#2a78d6")
    # meterskala
    for x in range(0, int(L) + 1, max(1, int(L // 8) or 1)):
        ax.plot([x, x], [-0.25, -0.33], color="#7f7f7f", lw=0.6)
        ax.text(x, -1.45, f"{x}", ha="center", va="top", fontsize=6, color="#7f7f7f")
    ax.text(L / 2, -1.8, "position (m)", ha="center", va="top", fontsize=6.5, color="#7f7f7f")
    # legend
    for i, g in enumerate((1, 2, 3, 4)):
        ax.plot(L * (0.55 + 0.11 * i), 1.75, "D", ms=6, mfc=GRAD_FARG_HEX[g], mec="white")
        ax.text(L * (0.565 + 0.11 * i), 1.75, f"grad {g}", va="center", fontsize=6.5)
    ax.plot(L * 0.02, 1.75, "v", ms=6, mfc="#2a78d6", mec="white")
    ax.text(L * 0.035, 1.75, "anslutning vänster (kl 7–11)", va="center", fontsize=6.5)
    ax.plot(L * 0.3, 1.75, "^", ms=6, mfc="#2a78d6", mec="white")
    ax.text(L * 0.315, 1.75, "höger (kl 1–5)", va="center", fontsize=6.5)
    ax.text(L * 0.02, 1.5, "vänster/höger sett i inspektionsriktningen", va="center", fontsize=6, color="#52514e")
    fig.tight_layout(pad=0.2)
    fig.savefig(path, dpi=200)
    plt.close(fig)


PROFIL_FIG = (7.4, 3.2)                 # tum
PROFIL_AX = (0.10, 0.17, 0.87, 0.70)    # axelns läge i figuren (andel av bredd/höjd)
LANGDSKALOR = [10, 20, 25, 40, 50, 75, 100, 150, 200, 250, 300, 400, 500, 750, 1000, 1500, 2000]
HOJDSKALOR = [1, 2, 5, 10, 20, 25, 50, 100, 200]


def rita_profil(s: Stracka, path: str, bild_bredd_mm: float) -> bool:
    """Inklinometerprofil ritad i exakt skala. Höjdpunkten läggs alltid till vänster så att
    profilen lutar ned mot höger. Returnerar False om profil saknas."""
    if len(s.profil) < 3:
        return False
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt

    x = [p[0] for p in s.profil]
    z = [p[2] for p in s.profil]
    vanster, hoger = s.fran_brunn, s.till_brunn
    if z[-1] > z[0]:                       # spegla så att den höga änden hamnar till vänster
        L = x[-1]
        x = [L - xi for xi in x][::-1]
        z = z[::-1]
        vanster, hoger = hoger, vanster
    x0, z0, x1, z1 = x[0], z[0], x[-1], z[-1]

    # fysisk storlek på axeln när bilden skrivs ut med bredden bild_bredd_mm
    bild_hojd_mm = bild_bredd_mm * PROFIL_FIG[1] / PROFIL_FIG[0]
    ax_w_mm = bild_bredd_mm * PROFIL_AX[2]
    ax_h_mm = bild_hojd_mm * PROFIL_AX[3]
    xspann = max(x1 - x0, 0.5) * 1.03
    zspann = max(max(z) - min(z), 0.05) * 1.15
    langdskala = next((k for k in LANGDSKALOR if xspann * 1000 / k <= ax_w_mm), LANGDSKALOR[-1])
    hojdskala = next((k for k in HOJDSKALOR if zspann * 1000 / k <= ax_h_mm), HOJDSKALOR[-1])
    xvidd = ax_w_mm * langdskala / 1000    # m som ryms i axeln
    zvidd = ax_h_mm * hojdskala / 1000
    zmitt = (max(z) + min(z)) / 2

    fig = plt.figure(figsize=PROFIL_FIG)
    ax = fig.add_axes(PROFIL_AX)
    ax.set_xlim(x0 - (xvidd - (x1 - x0)) / 2, x0 - (xvidd - (x1 - x0)) / 2 + xvidd)
    ax.set_ylim(zmitt - zvidd / 2, zmitt + zvidd / 2)
    ax.plot([x0, x1], [z0, z1], color="#4d4d4d", lw=0.9, ls="--", label="rät linje mellan brunnarna")
    # Höjderna i TV3-filen är avrundade till hela cm, vilket ger en trappstegsformad linje på flacka
    # ledningar. Linjen jämnas ut med ett glidande medelvärde (±0,3 m) enbart för uppritningen;
    # svacka, bakfall och lutning beräknas på rådata.
    zj = []
    for i, xi in enumerate(x):
        grannar = [zk for xk, zk in zip(x, z) if abs(xk - xi) <= 0.3]
        zj.append(sum(grannar) / len(grannar))
    ax.plot(x, zj, color="#2a78d6", lw=1.6, label="uppmätt profil (utjämnad, cm-upplösning i filen)")
    pa = s.profil_analys
    if pa and pa["svackdjup"] > 0.01:
        xs = pa["svackpos"]
        if z[-1] > z[0] if False else (s.profil[-1][2] > s.profil[0][2]):
            xs = s.profil[-1][0] - xs                       # speglad axel
        # höjd i profilen vid svackans position
        zi = min(z, key=lambda v: v)  # fallback
        j = min(range(len(x)), key=lambda i: abs(x[i] - xs))
        zi = z[j]
        ax.plot([xs, xs], [zi, zi + pa["svackdjup"]], color="#d03b3b", lw=1.5)
        ax.annotate(f"svacka {pa['svackdjup']:.2f} m", (xs, zi), xytext=(0, -14), textcoords="offset points",
                    ha="center", fontsize=7.5, color="#d03b3b")
    if pa and pa["osaker"]:
        ax.text(0.99, 0.03, "OBS: inklinometerprofilen avviker från brunnshöjderna – osäker", transform=ax.transAxes,
                ha="right", va="bottom", fontsize=7, color="#d03b3b")
    ax.plot([x0, x1], [zj[0], zj[-1]], "o", ms=5, color="#4d4d4d")
    ax.annotate(f"{vanster}  {z0:.2f}", (x0, z0), xytext=(4, 6), textcoords="offset points", fontsize=7.5, fontweight="bold")
    ax.annotate(f"{hoger}  {z1:.2f}", (x1, z1), xytext=(-4, -12), textcoords="offset points", ha="right",
                fontsize=7.5, fontweight="bold")
    lut = s.lutning_promille
    titel = "Inklinometerprofil"
    if lut is not None:
        titel += f"   ·   lutning {abs(lut):.1f} ‰"
    titel += f"   ·   höjdskala 1:{hojdskala}   ·   längdskala 1:{langdskala}"
    ax.set_title(titel, fontsize=8.5, loc="left", fontweight="bold")
    ax.set_xlabel(f"position (m), 0 = {vanster}", fontsize=7.5)
    ax.set_ylabel("höjd (m)", fontsize=7.5)
    ax.tick_params(labelsize=7)
    ax.grid(color="#e1e0d9", lw=0.6)
    for sp in ("top", "right"):
        ax.spines[sp].set_visible(False)
    ax.legend(fontsize=7, frameon=False, loc="upper right")
    fig.savefig(path, dpi=200)
    plt.close(fig)
    return True


def skriv_rapport(s: Stracka, path: str, tmp: str) -> None:
    """Skriver en PDF-rapport (inspektionsprotokoll) för en sträcka."""
    from reportlab.lib import colors
    from reportlab.lib.enums import TA_LEFT
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.units import mm
    from reportlab.platypus import (Image, KeepTogether, PageBreak, Paragraph, SimpleDocTemplate, Spacer, Table,
                                    TableStyle)

    normal, fet = _pdf_typsnitt()
    st = ParagraphStyle("n", fontName=normal, fontSize=8.5, leading=11)
    st_liten = ParagraphStyle("l", fontName=normal, fontSize=7.5, leading=9.5, textColor=colors.HexColor("#52514e"))
    st_fet = ParagraphStyle("f", fontName=fet, fontSize=8, leading=10.5)
    st_h1 = ParagraphStyle("h1", fontName=fet, fontSize=15, leading=18, textColor=colors.HexColor("#1f3864"), spaceAfter=2)
    st_h2 = ParagraphStyle("h2", fontName=fet, fontSize=10.5, leading=13, textColor=colors.HexColor("#1f3864"),
                           spaceBefore=8, spaceAfter=4)
    st_vit = ParagraphStyle("v", fontName=fet, fontSize=7.5, leading=9.5, textColor=colors.white)
    st_cell = ParagraphStyle("c", fontName=normal, fontSize=7.5, leading=9.5)

    bredd = A4[0] - 30 * mm
    klassfarg = colors.HexColor(KLASS_FARG_HEX[s.klass])

    def P(t, stil=st):
        return Paragraph(str(t).replace("&", "&amp;").replace("<", "&lt;"), stil)

    def sidfot(canvas, doc):
        canvas.saveState()
        canvas.setFont(normal, 7)
        canvas.setFillColor(colors.HexColor("#898781"))
        canvas.drawString(15 * mm, 8 * mm, f"{s.fil}  ·  sträcka {s.nr}  ·  {s.startbrunn} → {s.slutbrunn}")
        canvas.drawRightString(A4[0] - 15 * mm, 8 * mm, f"sida {doc.page}")
        canvas.drawRightString(A4[0] - 15 * mm, A4[1] - 10 * mm, f"{s.projekt}  ·  {s.omrade}")
        canvas.drawString(15 * mm, A4[1] - 10 * mm, "Inspektionsprotokoll (TV-inspektion, P93)")
        canvas.setStrokeColor(colors.HexColor("#c3c2b7"))
        canvas.line(15 * mm, A4[1] - 12 * mm, A4[0] - 15 * mm, A4[1] - 12 * mm)
        canvas.restoreState()

    doc = SimpleDocTemplate(path, pagesize=A4, leftMargin=15 * mm, rightMargin=15 * mm,
                            topMargin=17 * mm, bottomMargin=15 * mm,
                            title=f"Inspektionsprotokoll sträcka {s.nr} {s.startbrunn}-{s.slutbrunn}",
                            author="tv3_analys")
    el = []

    # --- rubrik + klass ---
    rub = Table([[P(f"Sträcka {s.nr}: {s.startbrunn} → {s.slutbrunn}", st_h1),
                  P(KLASS_TEXT[s.klass], ParagraphStyle("k", parent=st_vit, fontSize=9.5, leading=12))]],
                colWidths=[bredd - 48 * mm, 48 * mm])
    rub.setStyle(TableStyle([("BACKGROUND", (1, 0), (1, 0), klassfarg), ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                             ("ALIGN", (1, 0), (1, 0), "CENTER"), ("LEFTPADDING", (0, 0), (0, 0), 0),
                             ("TOPPADDING", (1, 0), (1, 0), 5), ("BOTTOMPADDING", (1, 0), (1, 0), 5)]))
    el.append(rub)
    el.append(Spacer(1, 4))

    # --- infotabell ---
    pa = s.profil_analys
    lut = s.lutning_promille
    dim = s.dimension + (f"/{s.dimension2}" if s.dimension2 else "") + " mm"
    info = [
        ("Område", s.omrade, "Datum", f"{s.datum} {s.klockslag}".strip()),
        ("Startbrunn (uppströms)", s.startbrunn, "Slutbrunn (nedströms)", s.slutbrunn),
        ("Kamera från", f"{s.fran_brunn} (position 0 m)", "Riktning", s.riktning),
        ("Inspekterad längd", f"{s.langd:.2f} m", "Ledningstyp", s.ledningstyp.capitalize()),
        ("Material", s.material + (f" (foder: {s.foder}, {s.fodermaterial})" if s.foder.strip() else ""),
         "Dimension / form", f"{dim}, {s.form.lower()}"),
        ("Antal anslutningar", str(s.antal_anslutningar), "Antal skador", str(len(s.skador()))),
        ("Prioritetsklass", KLASS_TEXT[s.klass], "Totalindex", f"{s.index():.1f} p/100 m"),
        ("Konstruktionsindex", f"{s.index('K'):.1f} p/100 m (maxgrad {s.maxgrad('K') or '–'})",
         "Driftindex", f"{s.index('D'):.1f} p/100 m (maxgrad {s.maxgrad('D') or '–'})"),
        ("Avbruten inspektion", "Ja" if s.avbruten else "Nej", "Väder", s.vader),
        ("Svacka (djup / längd)", f"{pa['svackdjup']:.2f} m / {pa['svacklangd']:.1f} m" if pa else "–",
         "Lutning", (f"{lut:.1f} ‰" if lut is not None else "–") + ((f"  ·  bakfall {pa['bakfall']:.1f} m" if pa and pa["bakfall"] > 0.5 else "")
                                                                    + ("  ·  profil osäker" if pa and pa["osaker"] else ""))),
        ("Videofil", s.videofil or "–", "TV3-fil", s.fil),
    ]
    rader = [[P(a, st_fet), P(b), P(c, st_fet), P(d)] for a, b, c, d in info]
    kw = 38 * mm
    t = Table(rader, colWidths=[kw, bredd / 2 - kw, kw, bredd / 2 - kw])
    t.setStyle(TableStyle([
        ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#d9d9d9")),
        ("BACKGROUND", (0, 0), (0, -1), colors.HexColor("#f2f4f8")),
        ("BACKGROUND", (2, 0), (2, -1), colors.HexColor("#f2f4f8")),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("TOPPADDING", (0, 0), (-1, -1), 2.5), ("BOTTOMPADDING", (0, 0), (-1, -1), 2.5),
    ]))
    el.append(t)

    # --- schema ---
    schema = os.path.join(tmp, f"schema_{s.nr}.png")
    rita_schema(s, schema)
    el.append(Paragraph("Översikt", st_h2))
    el.append(Image(schema, width=bredd, height=bredd * 2.5 / 7.4))

    # --- observationstabell ---
    el.append(Paragraph("Observationer", st_h2))
    huvud = [P(h, st_vit) for h in ("Pos. m", "Tid", "Kod", "Observation", "Kl.", "Nivå", "Foto", "Grad", "Poäng")]
    rader = [huvud]
    stil = [
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1f3864")),
        ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#d9d9d9")),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("TOPPADDING", (0, 0), (-1, -1), 2), ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
        ("ALIGN", (7, 1), (8, -1), "CENTER"),
    ]
    obs = [o for o in s.observationer if o.kod or o.infokod]
    for i, o in enumerate(obs, 1):
        klocka = o.klocka_fran + (f"–{o.klocka_till}" if o.klocka_till and o.klocka_till != o.klocka_fran else "")
        foto = ", ".join(n for n, _ in o.bilder)
        rader.append([P(f"{o.lage:.2f}", st_cell), P(o.tid, st_cell), P(o.kod or o.infokod, st_cell),
                      P(o.beskrivning(), st_cell), P(klocka, st_cell),
                      P(f"{o.vattenniva}%" if o.vattenniva and o.vattenniva != "0" else "", st_cell),
                      P(foto, st_cell), P(o.grad or "", st_cell), P(f"{o.poang:g}" if o.poang else "", st_cell)])
        if o.grad and o.ar_skada:
            stil.append(("BACKGROUND", (0, i), (-1, i), colors.HexColor(GRAD_FARG_LJUS.get(o.grad, "#ffffff"))))
    if len(rader) == 1:
        rader.append([P("Inga observationer registrerade", st_cell)] + [P("", st_cell)] * 8)
    cw = [13 * mm, 17 * mm, 10 * mm, None, 14 * mm, 11 * mm, 36 * mm, 12 * mm, 12 * mm]
    cw[3] = bredd - sum(w for w in cw if w)
    t = Table(rader, colWidths=cw, repeatRows=1)
    t.setStyle(TableStyle(stil))
    el.append(t)
    el.append(Spacer(1, 4))

    # --- profil ---
    profil = os.path.join(tmp, f"profil_{s.nr}.png")
    if rita_profil(s, profil, bredd / mm):
        el.append(KeepTogether([Paragraph("Profil", st_h2),
                                Image(profil, width=bredd, height=bredd * PROFIL_FIG[1] / PROFIL_FIG[0])]))
    else:
        el.append(Paragraph("Profil", st_h2))
        el.append(P("Ingen inklinometerprofil finns registrerad för sträckan.", st_liten))

    # --- foton ---
    bilder = [(o, n, pth) for o in s.observationer for n, pth in o.bilder]
    hittade = [(o, n, pth) for o, n, pth in bilder if pth]
    saknade = [n for o, n, pth in bilder if not pth]
    if hittade:
        el.append(PageBreak())
        el.append(Paragraph("Inspektionsfotografier", st_h2))
        celler = []
        bw = bredd / 2 - 4 * mm
        for o, n, pth in hittade:
            try:
                from PIL import Image as PILImage
                with PILImage.open(pth) as im:
                    w, h = im.size
                bh = bw * h / w
                if bh > 95 * mm:
                    bh, bw_ = 95 * mm, 95 * mm * w / h
                else:
                    bw_ = bw
                img = Image(pth, width=bw_, height=bh)
            except Exception:
                img = P(f"[kunde inte läsa {n}]", st_liten)
            esc = lambda t: str(t).replace("&", "&amp;").replace("<", "&lt;")
            txt = Paragraph(f"<b>{esc(n)}</b> · {o.lage:.2f} m · {esc(o.tid)}<br/>{esc(o.beskrivning())}", st_liten)
            celler.append([img, txt])
        # två per rad
        rader = []
        for i in range(0, len(celler), 2):
            par = celler[i:i + 2]
            rader.append([Table([[c[0]], [c[1]]], colWidths=[bw]) for c in par] + ([""] if len(par) == 1 else []))
        if rader:
            t = Table(rader, colWidths=[bredd / 2, bredd / 2])
            t.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP"), ("LEFTPADDING", (0, 0), (-1, -1), 0),
                                   ("BOTTOMPADDING", (0, 0), (-1, -1), 6)]))
            el.append(t)

    doc.build(el, onFirstPage=sidfot, onLaterPages=sidfot)


def _saker_filnamn(t: str) -> str:
    return re.sub(r"[^A-Za-z0-9ÅÄÖåäö._-]+", "_", t).strip("_")


def rapport_filnamn(s: Stracka) -> str:
    return _saker_filnamn(f"{os.path.splitext(s.fil)[0]}_{s.nr:03d}_{s.klass}_{s.startbrunn}-{s.slutbrunn}") + ".pdf"


def skriv_rapporter(strackor: list[Stracka], katalog: str, urval: str) -> int:
    """Skriver en PDF per sträcka i <katalog>. urval: alla | AB | A."""
    import tempfile
    try:
        import reportlab  # noqa: F401
    except ImportError:
        print("  reportlab saknas – inga PDF-rapporter skapas (pip install reportlab)")
        return 0
    os.makedirs(katalog, exist_ok=True)
    valda = [s for s in strackor if urval == "alla" or s.klass in urval]
    n = 0
    with tempfile.TemporaryDirectory() as tmp:
        for i, s in enumerate(valda, 1):
            namn = rapport_filnamn(s)
            try:
                skriv_rapport(s, os.path.join(katalog, namn), tmp)
                s.rapport_fil = os.path.join(os.path.basename(katalog), namn)
                n += 1
            except Exception as e:
                print(f"  FEL rapport sträcka {s.nr} ({s.fil}): {e}")
            if i % 25 == 0 or i == len(valda):
                print(f"  rapporter: {i}/{len(valda)}", end="\r")
    print()
    return n

# ----------------------------------------------------------------------------
# main
# ----------------------------------------------------------------------------

def las_listfil(path: str) -> tuple[list[tuple[str, list[str]]], list[str]]:
    """Läser en listfil. Returnerar ([(tv3-sökväg, [mediakataloger för just den filen]), ...],
    [mediakataloger som gäller alla filer]).

    Format (en post per rad, tomma rader och #-kommentarer ignoreras):
        media: D:\\Inspektioner\\Filmer          gäller alla TV3-filer i listan
        DUF 701.TV3                              TV3-fil; media söks i filens egen katalog
        DUF 702.TV3 ; D:\\Filmer\\DUF702          TV3-fil med egen mediakatalog (fler kan
                                                 anges, separerade med ;)
        C:\\Inspektioner\\2022\\                  katalog: alla TV3-filer i den
    Relativa sökvägar tolkas relativt listfilens katalog."""
    bas = os.path.dirname(os.path.abspath(path))

    def abs_(p: str) -> str:
        p = p.strip().strip('"').strip("'")
        return p if os.path.isabs(p) else os.path.normpath(os.path.join(bas, p))

    poster: list[tuple[str, list[str]]] = []
    globala: list[str] = []
    for rad in las_text(path).splitlines():
        rad = rad.strip()
        if not rad or rad.startswith("#"):
            continue
        m = re.match(r"^(media|film|bilder|filmer)\s*[:=]\s*(.+)$", rad, re.IGNORECASE)
        if m:
            globala += [abs_(d) for d in m.group(2).split(";") if d.strip()]
            continue
        delar = [d for d in rad.split(";")]
        tv3 = abs_(delar[0])
        media = [abs_(d) for d in delar[1:] if d.strip()]
        poster.append((tv3, media))
    return poster, globala


def hitta_tv3_filer(argument: list[str], listfiler: list[str]) -> tuple[list[tuple[str, list[str]]], list[str]]:
    """Löser upp argument (filer, kataloger, jokertecken, .txt-listor) till
    [(tv3-fil, [mediakataloger]), ...] samt globala mediakataloger från listfiler."""
    import glob
    kandidater: list[tuple[str, list[str]]] = []
    globala: list[str] = []
    for lf in listfiler:
        p, g = las_listfil(lf)
        kandidater += p
        globala += g
    for arg in argument:
        if arg.lower().endswith(".txt"):          # listfil även utan -l
            p, g = las_listfil(arg)
            kandidater += p
            globala += g
        elif any(ch in arg for ch in "*?["):
            kandidater += [(f, []) for f in sorted(glob.glob(arg))]
        else:
            kandidater.append((arg, []))

    filer: list[tuple[str, list[str]]] = []
    for k, media in kandidater:
        if os.path.isdir(k):
            for rot, _, namn in os.walk(k):
                filer += [(os.path.join(rot, n), media) for n in sorted(namn) if n.lower().endswith(".tv3")]
        else:
            filer.append((k, media))

    # ta bort dubbletter, behåll ordning (mediakataloger slås ihop)
    index: dict[str, int] = {}
    unika: list[tuple[str, list[str]]] = []
    for f, media in filer:
        key = os.path.abspath(f)
        if key in index:
            for m in media:
                if m not in unika[index[key]][1]:
                    unika[index[key]][1].append(m)
        else:
            index[key] = len(unika)
            unika.append((f, list(media)))
    return unika, globala


def main(argv=None):
    ap = argparse.ArgumentParser(
        description="Analys av TV3-filer (Svenskt Vatten TV-fil v3.0)",
        epilog="Exempel: python tv3_analys.py -l filer.txt -o resultat_2024 --topp 20")
    ap.add_argument("filer", nargs="*", help=".TV3-filer, kataloger, jokertecken (*.TV3) eller .txt-listfiler")
    ap.add_argument("-l", "--lista", action="append", default=[], metavar="FIL.TXT",
                    help="textfil med en TV3-sökväg per rad (kan anges flera gånger)")
    ap.add_argument("-o", "--utdata", default="tv3_resultat", help="utdatakatalog (standard: tv3_resultat)")
    ap.add_argument("--topp", type=int, default=15, help="antal sträckor i topplistan (standard: 15)")
    ap.add_argument("--rapporter", choices=["alla", "AB", "A", "inga"], default="alla",
                    help="PDF-rapport per sträcka: alla (standard), bara klass A och B, bara A, eller inga")
    ap.add_argument("--karta", choices=["ja", "nej"], default="ja" if SKRIV_KARTUNDERLAG else "nej",
                    help=f"skriv {KARTUNDERLAG_FIL} för ArcMap-skriptet i arcmap/ (standard: ja)")
    ap.add_argument("--diagram", action="store_true", default=SPARA_DIAGRAM,
                    help="spara diagrammen som PNG i <utdata>/diagram; de bäddas alltid in i Excel")
    ap.add_argument("--media", action="append", default=[], metavar="KATALOG",
                    help="extra katalog att söka video-/bildfiler i (kan anges flera gånger); "
                         "TV3-filens egen katalog söks alltid")
    a = ap.parse_args(argv)

    if not a.filer and not a.lista:
        ap.error("ange minst en TV3-fil, katalog eller listfil (-l filer.txt)")

    filer, globala_media = hitta_tv3_filer(a.filer, a.lista)
    if not filer:
        sys.exit("Inga TV3-filer hittades.")
    print(f"{len(filer)} fil(er) att analysera\n")

    strackor: list[Stracka] = []
    fel: list[str] = []
    for p, media in filer:
        for m in media:
            if not os.path.isdir(m):
                fel.append(f"{p}: mediakatalogen finns inte: {m}")
                print(f"  VARNING mediakatalog saknas: {m}")
        if not os.path.exists(p):
            fel.append(f"{p}: filen finns inte")
            print(f"  SAKNAS  {p}")
            continue
        try:
            st = las_tv3(p)
        except Exception as e:  # trasig fil ska inte stoppa hela körningen
            fel.append(f"{p}: {e}")
            print(f"  FEL     {p}: {e}")
            continue
        for st_ in st:
            st_.media_kataloger = list(media)
        print(f"  OK      {os.path.basename(p)}: {len(st)} sträckor, {sum(s.langd for s in st):.0f} m, "
              f"{sum(len(s.skador()) for s in st)} skadeobservationer")
        strackor += st
    if not strackor:
        sys.exit("Inga sträckor hittades.")
    if fel:
        print(f"\n{len(fel)} varning(ar) – se {os.path.join(a.utdata, 'fel.txt')}")

    for m in globala_media + a.media:
        if not os.path.isdir(m):
            print(f"  VARNING mediakatalog saknas: {m}")
    vh, vt, bh, bt = koppla_media(strackor, globala_media + a.media)
    print(f"\nVideofiler hittade: {vh} av {vt}   Bilder hittade: {bh} av {bt}")
    if vt and vh < vt:
        print("  (ange katalogen med filmerna i listfilen – 'media: KATALOG' eller 'fil.TV3 ; KATALOG' –\n"
              "   eller med --media KATALOG för att få klickbara länkar i Excel)")

    os.makedirs(a.utdata, exist_ok=True)
    # Diagrammen behövs som filer för att kunna bäddas in i Excel. Ska de inte sparas
    # ritas de i en temporär katalog som städas bort efteråt.
    import shutil
    import tempfile
    diagramkatalog = os.path.join(a.utdata, "diagram") if a.diagram else tempfile.mkdtemp(prefix="tv3_diagram_")
    diagram = rita_diagram(strackor, diagramkatalog, a.topp)
    if a.rapporter != "inga":
        print(f"\nSkriver PDF-rapporter ({a.rapporter}) ...")
        n = skriv_rapporter(strackor, os.path.join(a.utdata, "rapporter"), a.rapporter)
        print(f"  {n} rapporter skrivna till {os.path.join(a.utdata, 'rapporter')}")
    skriv_excel(strackor, os.path.join(a.utdata, "prioritering.xlsx"), diagram, a.topp)
    if a.karta == "ja":
        kartfil = os.path.join(a.utdata, KARTUNDERLAG_FIL)
        n_poster = skriv_kartunderlag(strackor, kartfil)
        print(f"\n{n_poster} sträckor skrivna till {kartfil} (underlag för ArcMap)")
    if not a.diagram:
        shutil.rmtree(diagramkatalog, ignore_errors=True)
    if fel:
        with open(os.path.join(a.utdata, "fel.txt"), "w", encoding="utf-8") as f:
            f.write("\n".join(fel) + "\n")

    klasser = Counter(s.klass for s in strackor)
    print("\nPrioritetsklasser: " + ", ".join(f"{KLASS_TEXT[k]}: {klasser.get(k, 0)}" for k in "ABCDE"))
    print(f"\nTopp {a.topp}:")
    for i, s in enumerate(sorterade_strackor(strackor)[:a.topp], 1):
        print(f"{i:>3}. [{s.klass}] {s.id:<28} {s.material:<7}{s.dimension:>4} {s.langd:6.1f} m  "
              f"index {s.index():6.1f}  {s.sammanfattning_skador()}")
    print(f"\nResultat skrivet till: {os.path.abspath(a.utdata)}")


if __name__ == "__main__":
    main()
