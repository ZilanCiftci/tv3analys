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
    python tv3_analys.py -l filer.txt --media D:/Filmer --bilder D:/Foton   # bilder i egen katalog
    python tv3_analys.py -l filer.txt --behall-rapporter  # skriv inte om PDF-rapporter som finns

Listfilen (-l/--lista) är en vanlig textfil. Tomma rader och rader som börjar
med # hoppas över. Relativa sökvägar tolkas relativt listfilens katalog.

    # exempel på listfil
    media: D:\\Inspektioner\\Filmer            filmkatalog som gäller alla filer
    bild: D:\\Inspektioner\\Foton              bildkatalog för alla filer (annars söks bilder
                                               i filmkatalogerna)
    littera: brunnslittera.csv                 ersättningslittera för felmärkta brunnar
    DUF 701.TV3                                TV3-fil (media söks även i filens egen katalog)
    DUF 702.TV3 ; D:\\Filmer\\DUF702 ; bild: E:\\Bilder  TV3-fil med egen film- och bildkatalog
    C:\\Inspektioner\\2022\\                    katalog: alla TV3-filer i den

Alla filer analyseras tillsammans i ett gemensamt underlag; kolumnen "Fil" i
Excel visar varifrån varje sträcka kommer.

Poängmodell (kan justeras i KONFIG nedan):
    grad 1 = 1 p, grad 2 = 3 p, grad 3 = 10 p, grad 4 = 30 p
    Konstruktionsskador (SPR, RBR, DEF, YTS, FOG, FRF, DEA) räknas fullt.
    Driftskador (ROT, INL, UTF, SED, INH) räknas med faktor 0,5 i driftindex, som redovisas
    separat och varken påverkar prioritetsklass eller rangordning.
    Konstruktionskoder viktas inbördes med KODFAKTOR: SPR/RBR/DEF 1,0, YTS 0,7,
    FOG/FRF 0,6, DEA 0,3 (None stänger av). Ytskada grad 4 (genomfrätt vägg) väger
    dock fullt (GRADFAKTOR) och cirkulära sprickor viktas med 0,7 (ATTRIBUTFAKTOR).
    Poängen normeras till poäng per 100 m ledning (sträckor kortare än
    MINLANGD räknas som MINLANGD, så att mycket korta sträckor inte överdrivs).
    Löpande skador (A1…B1) räknas vid startmarkeringen, viktade med längden:
    poäng × längd / 10 m (minst 1, högst 5) – se LOPANDE_ENHET_M / LOPANDE_TAK.

Prioritetsklass (för renoveringsbehov):
    A – Åtgärda           : grad 4 på RBR/DEF (GRAD4_KODER_A), eller konstruktionsindex ≥ 80 p/100 m
    B – Planera renovering: konstruktionsgrad 3, eller konstruktionsindex ≥ 25 p/100 m
    C – Bevaka            : övriga sträckor med registrerade skador
    D – Inga skador       : inga skadeobservationer
    E – Ej bedömd         : ingen inspekterad längd (< 1 m)
Inom varje klass rangordnas sträckorna efter konstruktionsindex.
Driftåtgärd (spolning/rotskärning) flaggas separat när driftgrad ≥ 3.
"""

from __future__ import annotations

import argparse
import hashlib
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

# Viktning av konstruktionskoder inbördes: poäng × KODFAKTOR[kod]. Koder som saknas räknas
# med 1,0. Tom ({}) eller None = ingen viktning, bara graden avgör.
# Syftet är att hitta ledningar som bör strumpinfodras medan det fortfarande går: sprickor,
# rörbrott och deformation väger fullt, ytskada (väggen tunnas ut men röret bär) 0,7,
# fogfel 0,6 och defekt anslutning 0,3 (åtgärdas ändå med hatt efter infodringen).
KODFAKTOR: dict[str, float] | None = {"YTS": 0.7, "FOG": 0.6, "FRF": 0.6, "DEA": 0.3}

# Vikt för en viss kod vid en viss grad – ersätter KODFAKTOR för just den kombinationen.
# Ytskada grad 4 betyder att rörväggen är genomfrätt (P93), i praktiken ett rörbrott, och väger
# därför fullt medan ytskada grad 1–3 väger 0,7. None = bara KODFAKTOR.
GRADFAKTOR: dict[tuple[str, int], float] | None = {("YTS", 4): 1.0}

# Viktning per attribut, utöver KODFAKTOR: poäng × ATTRIBUTFAKTOR[(kod, attribut)]. Cirkulära
# sprickor beror oftast på en sättning vid en fog och är mindre allvarliga för bärigheten än
# komplexa och längsgående. None = ingen attributviktning.
ATTRIBUTFAKTOR: dict[tuple[str, str], float] | None = {("SPR", "CIRK"): 0.7}

# Löpande skador (A1…B1) viktas med längden: poängen multipliceras med längd / LOPANDE_ENHET_M
# (minst 1, högst LOPANDE_TAK). 10 m och 5× betyder att en 30 m löpande skada räknas som tre
# punktskador och att ingen löpande skada räknas som mer än fem. None = räkna en gång, oavsett längd.
LOPANDE_ENHET_M = 10.0
LOPANDE_TAK = 5.0

# Attribut/typkoder (kolumn 8) – för läsbara texter
ATTRIBUT = {
    "KOMPL": "komplexa", "CIRK": "cirkulära", "LÄNGS": "längsgående",
    "TUNNA": "tunna", "GROVA": "grova", "PAKET": "rotpaket",
    "RIKTN": "riktningsavvikelse", "TVÄRS": "tvärsförskjutning", "LÄFÖR": "längsförskjutning",
    "FAFOG": "fel i fog", "UTFÄL": "utfällning", "FINT": "fint", "GROVT": "grovt",
    "HINDE": "hinder", "ANNAN": "annan orsak", "EJÖPP": "ej öppen",
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
    "A": "A – Åtgärda",
    "B": "B – Planera renovering",
    "C": "C – Bevaka",
    "D": "D – Inga skador",
    "E": "E – Ej bedömd",
}
KLASS_FARG = {"A": "D03B3B", "B": "EC835A", "C": "FAB219", "D": "0CA30C", "E": "BFBFBF"}
KLASS_FARG_HEX = {k: "#" + v for k, v in KLASS_FARG.items()}

# Skadekoder där en enda observation av grad 4 räcker för klass A (rörbrott, deformation).
# Grad 4 på andra konstruktionskoder (ytskada, fogfel …) ger poäng som vanligt men inte
# automatiskt klass A. None = alla konstruktionskoder (tidigare beteende).
GRAD4_KODER_A = {"RBR", "DEF"}

TROSKEL_A = 80.0   # konstruktionsindex p/100 m
TROSKEL_B = 25.0
MINLANGD = 20.0    # m – nämnare vid normering av korta sträckor
SVACKA_MAX_M = 1.0  # m – större beräknat svackdjup än så är en inklinometerartefakt (driftande
                    # profil); svackan redovisas då som okänd och profilen markeras osäker
BAKFALL_MIN_PROMILLE = 5.0   # ‰ – lutning mot flödesriktningen som räknas som bakfall
BAKFALL_SEGMENT_MIN_M = 1.0  # m – kortare bakfallssträckor än så redovisas inte som segment i
                             # kartunderlaget (cm-upplösningen ger annars ett segment per trappsteg);
                             # luckor kortare än så mellan två segment slås ihop

# Markprofil (markprofil.json från ArcMap-verktyget "Markprofil"): filmens höjder hängs upp på
# GIS-vattengången i brunnarna (RH2000), eftersom filmerna ofta har ett lokalt nollplan.
HOJD_SAMMA_M = 0.3      # m – avviker filens brunnshöjder mindre än så från GIS är de i samma system
HOJD_FALL_TOL_M = 0.3   # m – skiljer sig fallet mellan brunnarna mer än så från GIS korrigeras
                        #     lutningen linjärt (inklinometerdrift), annars bara en förskjutning
TACKNING_MIN_M = 1.0    # m – mindre täckning (mark − hjässa) än så flaggas
# Rimlighet hos höjderna i TV3-filen (PROFILADM/PROFILDAT): värden under HOJD_SAKNAS_UNDER (SVOA: −9999)
# räknas som saknade; skiljer brunnshöjderna mer än HOJD_ORIMLIG_M, eller spänner inklinometerprofilen över
# mer än så, är höjderna fel i filen – lutning och profil redovisas då som okända ("höjdfel i filen" i
# Höjdflagga, protokollet och kartunderlaget) i stället för som −26 000 ‰ och en lodrät profil.
HOJD_SAKNAS_UNDER = -999
HOJD_ORIMLIG_M = 50
OFULLSTANDIG_ANDEL = 0.85  # är filmad längd kortare än så gånger kartlängden nådde kameran inte
# GIS-data (gisdata.json från ArcMap-verktyget "Exportera GIS-data"): brunnar med locknivå och
# ledningssträckor mellan brunnar med vattengång, dimension, material och år. Analysen kontrollerar
# filmen mot GIS och sätter en GIS-flagga per sträcka (kolumn i Prioritering, fält i JSON).
# Flera filmer av samma brunnspar (omfilmning efter spolning, eller från andra hållet efter avbrott).
# OMFILMNING: en hel film ersätter ofullständiga filmer och äldre hela filmer av samma brunnspar
# ("ersatt av nr X"); ersatta filmer finns kvar i Prioritering och får rapport men räknas inte i
# statistik, karta, åtgärdspaket eller inspektionsgrad. En nyare hel film ersätter en äldre bara om
# den är minst OMFILMNING_MIN_ANDEL av den äldres längd. SAMMANSLAGNING: två ofullständiga filmer
# från var sin brunn slås ihop till en extra sträcka ("sammanslagen av nr 72 + 73") på en gemensam
# axel från startbrunnen – kartlängden (markprofil/GIS) ger totallängd och överlapp, annars antas
# filmerna mötas utan överlapp; delfilmerna markeras "ingår i sammanslagen nr Y".
# OMFILMNING_OVER_FILER: samma brunnspar i olika TV3-filer (olika inspektionsomgångar) – den nyaste filmen
# gäller, äldre filmer markeras "ersatt av ny inspektion nr X i FIL (datum)" och räknas inte i statistiken;
# den gällande sträckan får "Tidigare inspektion" (datum, klass och konstruktionsindex då → nu, förvärrad/
# förbättrad) i Prioritering, protokollet och kartunderlaget. Samma dag i olika filer = dubblett, en räknas.
OMFILMNING = True
OMFILMNING_MIN_ANDEL = 0.85
SAMMANSLAGNING = True
OMFILMNING_OVER_FILER = True
DELFILMER_SEPARAT = False     # True: delfilmerna i en sammanslagning listas även var för sig i Prioritering/
                              # Observationer och får egna protokoll; False: bara den sammanslagna sträckan
TIDIGARE_INDEX_ANDEL = 0.2    # konstruktionsindex ändrat mer än så (och minst 5 p/100 m) inom samma klass = förvärrad/förbättrad
GIS_FALL_TOL_M = 0.3        # m – fallet mellan brunnarna (film mot GIS) får avvika så mycket …
GIS_FALL_TOL_ANDEL = 0.5    # … eller så stor andel av GIS-fallet (det största gäller) innan det flaggas
GIS_RIKTNING_MIN_M = 0.05   # m – stiger GIS-vattengången mer än så från start- till slutbrunn flaggas riktningen
GIS_SAKNAS_UNDER = -999     # vattengång/locknivå i GIS under så här räknas som saknad (SVOA: −9999)
GIS_FALL_ORIMLIGT_M = 50    # m – större fall mellan brunnarna i GIS flaggas som orimligt i stället för riktning/fall
GIS_PLATSHALLARE = {"AG", "STBEXTRA"}   # littera i TV3-filen som inte är riktiga brunnar – kontrolleras inte mot GIS
GRENROR_LITTERA = r"^A{1,2}G\d*$"   # littera som betyder anslutning till grenrör/påstick på en annan ledning (AG, AAG1 …):
                                    # ingen brunn finns – sträckan kopplas till GIS-ledningen från den kända brunnen som
                                    # slutar i en fri ände. Samma hantering för servisanslutning och ände utan brunn
                                    # (propp): andra littera markeras med ratt = grenrör / servis / ände i brunnslittera.csv
GIS_GRENROR_LANGD_TOL = 0.3         # andel – GIS-ledningen med fri ände får avvika så mycket (+ 3 m) från filmens längd
GIS_LITTERA_LIKHET = 0.75   # 0–1 – så lika måste ett GIS-littera vara för att föreslås för ett okänt (0,75 fångar två omkastade siffror)
GIS_LANGD_TOL = 0.15        # andel – en GIS-ledning vars längd ligger inom så många % av filmens föreslås som rätt
GIS_VAG_MAX_HOPP = 4        # saknas direkt ledning i GIS mellan brunnarna söks en väg via andra brunnar (högst så
                            # många ledningar, högst GIS_VAG_MAX_ANDEL × filmens längd + 20 m); annars tas brunnarnas
                            # avstånd fågelvägen som kartlängd ("kartlängd fågelvägen")
GIS_VAG_MAX_ANDEL = 2.0
GIS_SKARV_TOL_M = None      # m – två ledningars fria ändar (ingen brunn) närmare varandra än så är en skarv (t.ex.
                            # materialbyte mitt på sträckan): delarna fogas ihop till en ledning mellan brunnarna.
                            # None = exportens tolerans (tolerans_m i gisdata.json), annars talet.
                            # sträcka när bara en av brunnarna är okänd (den andra brunnens grannar i GIS prövas)
# Materialnamn i film och GIS som ska räknas som samma (versaler; vänster = som det står, höger = grupp)
GIS_MATERIAL = {"BTG": "Betong", "BT": "Betong", "BETONG": "Betong", "BET": "Betong", "CONCRETE": "Betong",
                "SEG": "Segjärn", "GJ": "Gjutjärn", "ODEF": "",     # SVOA: Bt, Seg, PVC, odef (odefinierat)
                "PVC": "Plast", "PE": "Plast", "PEH": "Plast", "PP": "Plast", "PLAST": "Plast", "PLASTIC": "Plast",
                "PEM": "Plast", "PEH": "Plast", "PRC": "Plast",
                "GAP": "GAP", "GRP": "GAP", "LERA": "Lera", "LER": "Lera", "LERGODS": "Lera", "TEGEL": "Tegel",
                "GJJ": "Gjutjärn", "GJUTJÄRN": "Gjutjärn", "SEGJ": "Segjärn", "SEGJÄRN": "Segjärn",
                "STÅL": "Stål", "ST": "Stål", "ASBEST": "Asbestcement", "ETERNIT": "Asbestcement",
                "AC": "Asbestcement"}
                           # fram – profilen hängs då bara upp i den brunn kameran startade i

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

# Åtgärdspaket och kostnad (PLAN_verktyg.md steg 3). Kostnaderna läses ur KOSTNADSFIL
# (post;dimension_fran;dimension_till;enhet;kr;kommentar – användaren äger filen, schablon tills
# vidare). Manuella bedömningar läses ur en tidigare prioritering.xlsx ('manuell:' i listfilen eller
# --manuell): kolumnerna Manuell bedömning (A–E, eller texten schakt/strumpa/ingen som styr metoden),
# Kommentar och Lagning (m).
KOSTNADSFIL = "kostnader.csv"          # söks relativt listfilen, annars bredvid skriptet
ATGARD_KLASSER = ("A",)                # gällande klasser som får åtgärd (strumpa) direkt
OVERBRYGGA_KLASSER = ("B",)            # klasser som bara tas med när de ligger MELLAN två åtgärdssträckor
                                       # (okt 2026, användaren: "huvudsakligen A; en eller två B-ledningar
                                       # mellan två A kan tas med – inte strumpa B bara för att vi är där")
OVERBRYGGA_MAX_STRACKOR = 2            # högst så många sträckor i rad får överbrygga mellan två åtgärdssträckor
SCHAKT_AUTOMATISKT = False             # F3: metoden byts bara manuellt. True = grad 4 på GRAD4_KODER_A
                                       # (rörbrott/deformation) ger schakt automatiskt
ETAPP_OVERBRYGGA_M = 60.0              # F5: C/D-sträcka kortare än så mellan två åtgärdssträckor tas med
                                       # (klasser utanför OVERBRYGGA_KLASSER överbryggar bara om de är så korta)
ETAPP_ORDNING = "index"                # "index" = högsta konstruktionsindex först,
                                       # "konsekvens" = flest serviser uppströms först (kräver uppstroms:)
BRUNNSTYP_ANDE_M = 1.5                 # m – TVDAT-kod NB/TB/RB så nära änden gäller för brunnen där
LAGNINGSKODER = {("RBR", 3), ("RBR", 4), ("DEF", 3), ("DEF", 4), ("FOG", 4), ("YTS", 4)}
                                       # skador som brukar kräva punktlagning före strumpning – listas i
                                       # Åtgärdsflagga när Lagning (m) inte är ifylld (F2: manuell bedömning)
FRAMSCHAKTA_BRUNNAR: list[str] = []    # F9: lås vilka brunnar som schaktas fram (littera); kan också
                                       # anges i fliken Etapper, kolumn "Schakta fram (manuellt)"

# Hoppa över PDF-rapporter som redan finns i utdatakatalogen (spar tid när bara Excel eller
# kartunderlaget ska uppdateras). Motsvarar --behall-rapporter. Ta bort rapporter/ eller kör
# utan flaggan när layouten ändrats.
BEHALL_RAPPORTER = False

# Hastighet för PDF-rapporterna. Fotona förminskas till högst FOTO_MAX_PX pixlar på längsta sidan
# innan de bäddas in (ett 1920×1080-foto på 1,2 MB blir ca 100 kB; i A4 visas fotot 85 mm brett,
# så 960 px är ~290 dpi – ingen synlig skillnad, men rapporten blir 5–8 gånger mindre och snabbare
# att skriva; 960 = halva full-HD, vilket JPEG-avkodaren klarar i ett steg). None = originalstorlek.
# RAPPORT_PROCESSER = antal parallella processer (None = antal kärnor − 1, 1 = en i taget).
FOTO_MAX_PX = 960
FOTO_JPEG_KVALITET = 76          # 76 ger ~30 % mindre filer än 82 utan synlig skillnad i A4
RAPPORT_PROCESSER: int | None = None
# Upplösning på översikts- och profilbilden i PDF-rapporten (dpi vid 180 mm bredd). 150 räcker för
# utskrift i A4 och halverar nästan bildernas storlek och rittid mot 200.
RAPPORT_BILD_DPI = 150
# Diagrambilderna i rapporten sparas som palett-PNG (256 färger). Reportlab bäddar in dem som RGB,
# men de kvantiserade pixlarna komprimeras bättre: ca 35 % mindre per figur i PDF:en utan synlig
# skillnad på linjediagram. Kostar ~0,1 s per rapport. False = vanlig RGB-PNG.
RAPPORT_BILD_PALETT = True

# Kolumner som döljs som standard i Excel (grupperade – fäll ut med plustecknet ovanför
# kolumnrubrikerna, eller Data > Dela upp grupp). Allt finns kvar i filen. Tom lista = visa allt.
# Ange rubriken utan enhet ("Längd", inte "Längd (m)") – enheten står på egen rad i Excel.
DOLDA_KOLUMNER = {
    "Prioritering": ["Fil", "Nr", "Område", "Ledningstyp", "Datum", "Driftindex", "Totalindex",
                     "Drift maxgrad", "Antal skador", "Serviser uppströms", "Längd uppströms", "Driftåtgärd",
                     "Inspekterad flera ggr", "Littera rättat",
                     "Svackdjup/diameter", "Svacklängd", "Bakfall längd", "Lutning", "Profil osäker",
                     "Höjdanpassning", "Täckning min", "Täckning max", "Höjdflagga",
                     "Driftområde", "Lutning GIS", "Djup start", "Djup slut", "Anläggningsår", "Renoveringsår GIS",
                     "Brunnstyp start", "Brunnstyp slut"],
    "Etapper": ["Sträckor (lista)"],
    "Observationer": ["Fil", "Typ", "Löpande", "Klocka till", "Vattennivå (%)"],
}


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
    film_nr: int | None = None           # ursprunglig delfilm (nr) när sträckan är sammanslagen

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
    def lopande_faktor(self) -> float:
        """Längdviktning för en löpande skada: längd / LOPANDE_ENHET_M, minst 1, högst LOPANDE_TAK."""
        if not LOPANDE_ENHET_M or not self.lopande.startswith("A") or not self.lopande_langd:
            return 1.0
        f = max(1.0, self.lopande_langd / LOPANDE_ENHET_M)
        return min(f, LOPANDE_TAK) if LOPANDE_TAK else f

    @property
    def poang(self) -> float:
        if not self.raknas:
            return 0.0
        p = GRADPOANG.get(self.grad, 0)
        if self.typ == "D":
            p *= DRIFTFAKTOR
        else:
            if GRADFAKTOR and (self.kod, self.grad) in GRADFAKTOR:
                p *= GRADFAKTOR[(self.kod, self.grad)]
            elif KODFAKTOR:
                p *= KODFAKTOR.get(self.kod, 1.0)
            if ATTRIBUTFAKTOR:
                p *= ATTRIBUTFAKTOR.get((self.kod, self.attribut), 1.0)
        return p * self.lopande_faktor

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
    hojdfel: str = ""               # höjderna i filen är orimliga (satt vid inläsning), t.ex. "brunnshöjder 20,63 / 2444,33 m"
    profil_start_z: float | None = None
    profil_slut_z: float | None = None
    tv3_sokvag: str = ""                 # absolut sökväg till TV3-filen
    rapport_fil: str | None = None       # relativ sökväg (från utdatakatalogen) till PDF-rapporten
    media_kataloger: list[str] = field(default_factory=list)   # film-/mediakataloger för just denna fil
    bild_kataloger: list[str] = field(default_factory=list)    # bildkataloger för just denna fil (valfritt)
    video_sokvag: str | None = None      # hittad sökväg till videofilen
    videofil_b: str | None = None        # andra filmen när sträckan är sammanslagen av två delfilmer
    video_sokvag_b: str | None = None

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
        grad4_a = any(o.grad == 4 and (GRAD4_KODER_A is None or o.kod in GRAD4_KODER_A)
                      for o in self.skador("K"))
        if grad4_a or kidx >= TROSKEL_A:
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
    filmstatus: str = ""            # "", "ersatt av nr X (…)", "ingår i sammanslagen nr Y", "sammanslagen av nr 72 + 73 …"
    sammanslagen_av: list = field(default_factory=list)   # [nr, nr] för en sammanslagen sträcka
    sammanslagning: dict | None = None   # {a_nr, b_nr, a_fran, b_fran, a_langd, b_langd, L, overlapp} för ritningen
    grenror_markerat: dict = field(default_factory=dict)   # sida ("start"/"slut") -> typ ur brunnslittera.csv

    def grenror(self, sida: str) -> bool:
        """True när änden (sida "start"/"slut") inte är en brunn utan ett grenrör/påstick på en annan
        ledning, en servisanslutning eller en ledningsände utan brunn: littera enligt GRENROR_LITTERA
        eller markerad med ratt = grenrör/servis/ände i litterafilen."""
        return self.grenror_typ(sida) is not None

    def grenror_typ(self, sida: str) -> str | None:
        """"grenrör", "servis" eller "ände utan brunn" för en ände utan brunn, annars None."""
        if sida in self.grenror_markerat:
            return self.grenror_markerat[sida]
        lit = self.startbrunn if sida == "start" else self.slutbrunn
        if GRENROR_LITTERA and re.match(GRENROR_LITTERA, _normlittera(lit), re.I):
            return "grenrör"
        return None
    tidigare: dict | None = None    # närmast föregående inspektion av brunnsparet i en annan TV3-fil (OMFILMNING_OVER_FILER):
                                    # {fil, nr, datum, klass, index, langd, antal_skador, utveckling, antal} – antal = alla äldre filmer
    littera_rattat: str = ""       # t.ex. "BDNB1005633→BDNB1015633" om brunnslittera ersatts från CSV
    # Från markprofil.json (ArcMap-verktyget Markprofil): markhöjder längs kartlinjen
    # [(m från startbrunn, höjd)], GIS-vattengång i brunnarna och kartlinjens längd
    mark: list[tuple[float, float]] = field(default_factory=list)
    gis_vg_start: float | None = None
    gis_vg_slut: float | None = None
    langd_karta: float | None = None
    hojdsystem: str = "RH2000"           # kartans höjdsystem enligt markprofil.json
    # Från gisdata.json (ArcMap-verktyget Exportera GIS-data), satt av koppla_gis:
    # {"ledning": post|None, "brunn_start": post|None, "brunn_slut": post|None,
    #  "vg_min_start", "vg_min_slut" (lägsta vattengång vid brunnen), "fil": sökväg}
    gis: dict | None = None
    # Från uppströmsanalysen i ArcMap (verktyget Uppströms, batchläge; CSV via uppstroms: i listfilen)
    serviser_uppstroms: int | None = None      # serviser/anslutningar uppströms, inkl. sträckans egna
    langd_uppstroms: float | None = None       # m ledning uppströms, inkl. sträckan
    serviser_kalla: str = ""                   # "servislager" eller "skattning (ANT_ANSL)"
    # Manuell bedömning ur en tidigare prioritering.xlsx (manuell: i listfilen / --manuell)
    manuell_bedomning: str = ""                # A–E, eller text som styr metoden (schakt/strumpa/ingen)
    kommentar: str = ""
    lagning_m: float | None = None             # meter punktlagning före strumpning (F2, manuell)
    # Åtgärdsplanering (planera_atgarder)
    etapp: int | None = None
    etapp_flagga: str = ""                     # t.ex. "medtagen för sammanhang"
    kostnad: dict | None = None                # {"strumpa", "hattar", "lagning", "summa"} i kr
    kostnadsflagga: str = ""                   # t.ex. "dimension 225 saknar pris", "kostnad på nr 72"
    _metod: str | None = None                  # satt av planeringen (överbryggning), annars metod_auto
    _cache: dict = field(default_factory=dict, repr=False, compare=False)

    @property
    def diameter_m(self) -> float:
        """Innerdiameter i m ur dimensionsfältet ('225', '225/300', 'Ø 400' …), 0 om otolkbart."""
        m = re.search(r"\d+(?:[.,]\d+)?", self.dimension or "")
        return float(m.group(0).replace(",", ".")) / 1000 if m else 0.0

    # ---- höjdläge: filmens höjder mot GIS ----
    @property
    def ofullstandig(self) -> bool:
        """Kameran nådde inte fram till den andra brunnen: avbruten inspektion (KAM/HINDE) eller
        filmad längd klart kortare än ledningen i kartan. Filens 'sluthöjd' är då höjden där
        kameran stannade, inte slutbrunnens."""
        if self.avbruten:
            return True
        return bool(self.langd_karta) and self.langd < OFULLSTANDIG_ANDEL * self.langd_karta

    def _filens_brunnshojder(self) -> tuple[float | None, float | None]:
        """Filens höjd vid start- och slutbrunn (flödesriktning): PROFILADM i första hand,
        annars inklinometerprofilens ändpunkter."""
        zs, ze = self.profil_start_z, self.profil_slut_z
        if self.ofullstandig and len(self.profil) >= 2:
            # Vid avbrott är PROFILADM:s värden opålitliga (höjden där kameran stannade, eller
            # kopierade från syskonfilmen) – ta ändpunkterna ur inklinometern
            zs = ze = None
        if (zs is None or ze is None) and len(self.profil) >= 2:
            z0, z1 = self.profil[0][2], self.profil[-1][2]
            if self.fran_brunn != self.startbrunn:
                z0, z1 = z1, z0
            zs = z0 if zs is None else zs
            ze = z1 if ze is None else ze
        return zs, ze

    @property
    def hojdanpassning(self) -> dict | None:
        """Hur filens höjder hängs upp på GIS-vattengången (RH2000). None om markprofil saknas
        eller filen inte har några höjder alls. Nycklar: status, offset (m), k (m/m i
        flödesriktningen), z_start/z_slut (korrigerade brunnshöjder)."""
        if "hojd" not in self._cache:
            self._cache["hojd"] = self._hojdanpassning()
        return self._cache["hojd"]

    def _hojdanpassning(self) -> dict | None:
        if self.langd < 1:                   # ej bedömd sträcka – inget att hänga upp
            return None
        zs, ze = self._filens_brunnshojder()
        gs, ge = self.gis_vg_start, self.gis_vg_slut
        if zs is None or ze is None:
            if gs is not None and ge is not None and not self.ofullstandig:
                return {"status": "GIS-vattengång, rät linje", "offset": None, "k": 0.0,
                        "z_start": gs, "z_slut": ge}
            return None
        if self.ofullstandig:
            # Bara brunnen kameran startade i är nådd: förskjut så att den änden hamnar på
            # GIS-nivån, ingen lutningskorrigering (den andra änden är inte en brunn)
            fran_start = self.fran_brunn == self.startbrunn
            g, z = (gs, zs) if fran_start else (ge, ze)
            if g is None:
                if self.langd_karta is None:
                    return None
                return {"status": "okänt nollplan (GIS-vattengång saknas)", "offset": 0.0, "k": 0.0,
                        "z_start": zs, "z_slut": ze}
            d = g - z
            if abs(d) <= HOJD_SAMMA_M:
                return {"status": f"{self.hojdsystem} ur filen", "offset": 0.0, "k": 0.0, "z_start": zs, "z_slut": ze}
            return {"status": f"förskjuten till GIS vid {self.fran_brunn} "
                              + ("(avbruten inspektion)" if self.avbruten else "(ofullständig film)"),
                    "offset": d, "k": 0.0, "z_start": zs + d, "z_slut": ze + d}
        if gs is None and ge is None:
            if self.langd_karta is None:
                return None
            return {"status": "okänt nollplan (GIS-vattengång saknas)", "offset": 0.0, "k": 0.0,
                    "z_start": zs, "z_slut": ze}
        if gs is None or ge is None:
            # Bara en brunn har GIS-nivå: förskjut mot den, ingen lutningskorrigering
            brunn, d = (self.startbrunn, gs - zs) if gs is not None else (self.slutbrunn, ge - ze)
            if abs(d) <= HOJD_SAMMA_M:
                return {"status": f"{self.hojdsystem} ur filen", "offset": 0.0, "k": 0.0, "z_start": zs, "z_slut": ze}
            return {"status": f"förskjuten till GIS vid {brunn} (bara en brunn har GIS-nivå)",
                    "offset": d, "k": 0.0, "z_start": zs + d, "z_slut": ze + d}
        ds, de = gs - zs, ge - ze
        if abs(ds) <= HOJD_SAMMA_M and abs(de) <= HOJD_SAMMA_M:
            return {"status": f"{self.hojdsystem} ur filen", "offset": 0.0, "k": 0.0, "z_start": zs, "z_slut": ze}
        L = self.langd
        if abs(ds - de) <= HOJD_FALL_TOL_M or L <= 0:
            return {"status": "förskjuten till GIS", "offset": ds, "k": 0.0,
                    "z_start": zs + ds, "z_slut": ze + ds}
        return {"status": "förskjuten och lutningskorrigerad", "offset": ds, "k": (de - ds) / L,
                "z_start": gs, "z_slut": ge}

    def _korr(self, x: float, z: float) -> float:
        """Korrigerad höjd för en punkt vid kamerans position x med filens höjd z."""
        h = self.hojdanpassning
        if not h or h["offset"] is None:
            return z
        d = x if self.fran_brunn == self.startbrunn else self.langd - x     # m från startbrunn
        return z + h["offset"] + h["k"] * d

    def profil_korrigerad(self) -> list[tuple[float, float]]:
        """Inklinometerprofilen [(kamerans position, höjd)] upphängd på GIS-nivåerna."""
        if "prof" not in self._cache:
            self._cache["prof"] = [(x, self._korr(x, z)) for x, _, z in self.profil]
        return self._cache["prof"]

    def ledningshojd(self, x: float) -> float | None:
        """Vattengångens höjd vid kamerans position x: ur den korrigerade profilen, annars rät
        linje mellan (korrigerade) brunnshöjderna."""
        prof = self.profil_korrigerad()
        if len(prof) >= 2:
            import bisect
            if "prof_x" not in self._cache:
                self._cache["prof_x"] = [px for px, _ in prof]
            xs = self._cache["prof_x"]
            if x <= xs[0]:
                return prof[0][1]
            if x >= xs[-1]:
                return prof[-1][1]
            i = bisect.bisect_right(xs, x)
            (x0, z0), (x1, z1) = prof[i - 1], prof[i]
            return z0 if x1 == x0 else z0 + (z1 - z0) * (x - x0) / (x1 - x0)
        h = self.hojdanpassning
        if not h or self.langd <= 0:
            return None
        d = x if self.fran_brunn == self.startbrunn else self.langd - x
        return h["z_start"] + (h["z_slut"] - h["z_start"]) * d / self.langd

    def mark_i_filmens_axel(self) -> list[tuple[float, float]]:
        """Markhöjderna omräknade till kamerans positioner: kartlinjens meter från startbrunnen
        skalas till filmens längd och vänds om kameran gick motströms."""
        if not self.mark or self.langd <= 0:
            return []
        fran_start = self.fran_brunn == self.startbrunn
        ut = []
        if self.ofullstandig:
            # Kameran nådde inte fram: ingen skalning, kartmeter från kamerans brunn = filmens
            # position, och bara den filmade delen tas med
            Lk = self.langd_karta or max(p for p, _ in self.mark)
            for pos, z in self.mark:
                x = pos if fran_start else Lk - pos
                if -0.5 <= x <= self.langd + 1.0:
                    ut.append((max(0.0, x), z))
            return sorted(ut)
        sk = (self.langd_karta / self.langd) if self.langd_karta else 1.0
        for pos, z in self.mark:
            d = pos / sk if sk else pos
            x = d if fran_start else self.langd - d
            ut.append((x, z))
        return sorted(ut)

    @property
    def tackning(self) -> dict | None:
        """Täckning = mark − hjässa (vattengång + innerdiameter) per markpunkt.
        {min, max, pos_min (kamerans position), n} eller None."""
        if "tack" not in self._cache:
            self._cache["tack"] = self._tackning()
        return self._cache["tack"]

    def _tackning(self) -> dict | None:
        mark = self.mark_i_filmens_axel()
        if not mark or self.hojdanpassning is None:
            return None
        dia = self.diameter_m
        # Är inklinometerprofilen osäker (driftar mot brunnshöjderna) räknas täckningen mot
        # rät linje mellan de korrigerade brunnshöjderna i stället för mot profilen.
        pa = self.profil_analys
        h = self.hojdanpassning
        rat_linje = (pa is None or pa["osaker"]) and h is not None and self.langd > 0
        varden = []
        for x, zm in mark:
            if rat_linje:
                d = x if self.fran_brunn == self.startbrunn else self.langd - x
                zl = h["z_start"] + (h["z_slut"] - h["z_start"]) * d / self.langd
            else:
                zl = self.ledningshojd(x)
            if zl is not None:
                varden.append((zm - (zl + dia), x))
        if not varden:
            return None
        lagst = min(varden)
        return {"min": lagst[0], "max": max(v for v, _ in varden), "pos_min": lagst[1], "n": len(varden)}

    @property
    def hojdflagga(self) -> str:
        if self.hojdfel:
            return self.hojdfel[0].upper() + self.hojdfel[1:]
        if self.langd_karta and not self.ofullstandig and self.langd > 1.25 * self.langd_karta:
            return f"Filmad längd {self.langd:.0f} m mot {self.langd_karta:.0f} m i kartan – fel sträcka?"
        t = self.tackning
        if not t:
            return ""
        if t["min"] < 0:
            return "Ledning över mark – höjdfel"
        if t["min"] < TACKNING_MIN_M:
            return f"Liten täckning ({t['min']:.1f} m)"
        return ""

    def _profil_i_flodesriktning(self) -> list[tuple[float, float]] | None:
        """Profilen (position, höjd) ordnad i flödesriktningen uppströms → nedströms,
        glesad till ca 0,25 m mellan punkterna för att dämpa mätbrus. Höjderna är
        korrigerade mot GIS när markprofil finns (påverkar bara svackan vid lutningskorrigering)."""
        if len(self.profil) < 3:
            return None
        pts = self.profil_korrigerad()
        if self.fran_brunn != self.startbrunn:      # kameran gick motströms
            pts = pts[::-1]
        ut = [pts[0]]
        for x, z in pts[1:]:
            if abs(x - ut[-1][0]) >= 0.25:
                ut.append((x, z))
        return ut if len(ut) >= 3 else None

    @property
    def profil_analys(self) -> dict | None:
        if "pa" not in self._cache:
            self._cache["pa"] = self._profil_analys()
        return self._cache["pa"]

    def _profil_analys(self) -> dict | None:
        """Svackor och bakfall ur inklinometerprofilen.
        svackdjup  – största stående vattendjup (m). Vattnet kan bara lämna ledningen
                     nedströms, så vattenytan i varje punkt ligger på den högsta punkten
                     nedströms om den (t.ex. utloppet vid bakfall) – djup = den nivån − höjden
        svacklangd – total längd (m) där stående vatten > 1 cm
        svackpos   – position (m från kamerans start) för djupaste punkten
        bakfall    – total längd (m) med lutning mot flödesriktningen (> BAKFALL_MIN_PROMILLE)
        bakfall_segment – sammanhängande bakfallssträckor (se _bakfall_segment), för kartan
        osaker     – True om profilen verkar opålitlig (starthöjd saknas, eller inklinometerns
                     fall avviker kraftigt från brunnshöjderna i PROFILADM)"""
        pts = self._profil_i_flodesriktning()
        if pts is None:
            return None
        n = len(pts)
        z = [p[1] for p in pts]
        ned = [0.0] * n                      # högsta punkt nedströms om (och med) punkt i
        m = -1e9
        for i in range(n - 1, -1, -1):
            m = max(m, z[i]); ned[i] = m
        djup = [max(0.0, ned[i] - z[i]) for i in range(n)]
        i_max = max(range(n), key=lambda i: djup[i])
        svacklangd = sum(abs(pts[i + 1][0] - pts[i][0]) for i in range(n - 1) if djup[i] > 0.01 or djup[i + 1] > 0.01)
        bakfall = 0.0
        stigande = []                        # (från, till) i kamerans positioner, i flödesordning
        for i in range(n - 1):
            dx = abs(pts[i + 1][0] - pts[i][0])
            if dx > 0 and (z[i + 1] - z[i]) / dx > BAKFALL_MIN_PROMILLE / 1000:   # stiger i flödesriktningen
                bakfall += dx
                stigande.append((i, i + 1))
        segment = self._bakfall_segment(pts, z, stigande)
        osaker = self.profil_start_z is None or self.profil_slut_z is None
        if not osaker:
            h = self.hojdanpassning
            if h and h["offset"] is not None and not self.ofullstandig:
                fall_brunnar = h["z_start"] - h["z_slut"]
            else:
                fall_brunnar = self.profil_start_z - self.profil_slut_z
            fall_inkl = z[0] - z[-1]
            if abs(fall_inkl - fall_brunnar) > max(0.3, 0.5 * abs(fall_brunnar)):
                osaker = True
        if djup[i_max] > SVACKA_MAX_M:
            # Orimligt djup = inklinometern har driftat; svackan går inte att bedöma
            return {"svackdjup": None, "svacklangd": None, "svackpos": None,
                    "bakfall": bakfall, "bakfall_segment": segment, "osaker": True}
        return {"svackdjup": djup[i_max], "svacklangd": svacklangd, "svackpos": pts[i_max][0],
                "bakfall": bakfall, "bakfall_segment": segment, "osaker": osaker}

    @staticmethod
    def _bakfall_segment(pts, z, stigande) -> list[dict]:
        """Slår ihop stigande delsträckor till bakfallssegment för kartan.
        Luckor kortare än BAKFALL_SEGMENT_MIN_M slås ihop, segment kortare än så tas bort.
        Varje segment: {"fran_m", "till_m" (kamerans positioner, fran_m < till_m),
        "langd_m", "lutning_promille" (stigning i flödesriktningen över segmentet)}."""
        if not stigande:
            return []
        grupper = []                          # [första index, sista index, Σdx stigande, Σdz stigande]
        for i0, i1 in stigande:
            dx, dz = abs(pts[i1][0] - pts[i0][0]), z[i1] - z[i0]
            if grupper and abs(pts[i0][0] - pts[grupper[-1][1]][0]) < BAKFALL_SEGMENT_MIN_M \
                    and z[i0] >= z[grupper[-1][1]] - 0.005:        # luckan får inte falla (cm-brus tillåts)
                grupper[-1][1] = i1
                grupper[-1][2] += dx
                grupper[-1][3] += dz
            else:
                grupper.append([i0, i1, dx, dz])
        ut = []
        for i0, i1, sdx, sdz in grupper:
            x0, x1 = pts[i0][0], pts[i1][0]
            langd = abs(x1 - x0)
            if langd < BAKFALL_SEGMENT_MIN_M:
                continue
            # Lutningen räknas över de stigande bitarna, inte över eventuella luckor
            ut.append({"fran_m": round(min(x0, x1), 1), "till_m": round(max(x0, x1), 1),
                       "langd_m": round(langd, 1),
                       "lutning_promille": round(sdz / sdx * 1000, 1) if sdx else None})
        return ut

    @property
    def svacka_andel(self) -> float | None:
        """Svackdjup i förhållande till rördiametern (0–1)."""
        a = self.profil_analys
        d = self.diameter_m
        return round(a["svackdjup"] / d, 2) if a and a["svackdjup"] is not None and d > 0 else None

    @property
    def max_svacka(self) -> float | None:
        a = self.profil_analys
        return a["svackdjup"] if a else None

    @property
    def lutning_promille(self) -> float | None:
        """Fall i flödesriktningen (‰) ur brunnshöjderna – de korrigerade när filen hängts upp
        på GIS-vattengången."""
        if self.langd <= 0:
            return None
        h = self.hojdanpassning
        if h and h["offset"] is not None:
            return (h["z_start"] - h["z_slut"]) / self.langd * 1000
        zs, ze = self._filens_brunnshojder()     # PROFILADM, annars inklinometerns ändpunkter
        if zs is None or ze is None:
            return None
        return (zs - ze) / self.langd * 1000

    @property
    def lutning_ur_inklinometer(self) -> bool:
        """Lutningen är räknad på inklinometerns ändpunkter (PROFILADM saknas eller är orimlig)."""
        h = self.hojdanpassning
        if h and h["offset"] is not None:
            return False
        return (self.profil_start_z is None or self.profil_slut_z is None) and len(self.profil) >= 2

    @property
    def fran_brunn(self) -> str:
        """Brunnen där kameran startade (position 0 m). Är utgångsbrunnen inte en av sträckans
        brunnar (saknas eller felstavad) används Riktning: Motströms = kameran startade i
        slutbrunnen."""
        if self.utgangsbrunn in (self.startbrunn, self.slutbrunn) and self.utgangsbrunn:
            return self.utgangsbrunn
        return self.slutbrunn if self.riktning.strip().lower().startswith("mot") else self.startbrunn

    @property
    def brunnar_omvanda(self) -> bool:
        """True om filens start-/slutbrunn verkar stå i kamerans riktning i stället för i
        flödesriktningen: Riktning och Utgångsbrunn motsäger varandra (Medströms men kameran
        startade i slutbrunnen, eller Motströms men i startbrunnen). Konventionen i TV3
        (WinCan) är startbrunn = uppströms; alla 282 sträckor i DUF 700/701 följer den."""
        r = self.riktning.strip().lower()
        u = self.utgangsbrunn.strip()
        if not u or not r or u not in (self.startbrunn, self.slutbrunn) or self.startbrunn == self.slutbrunn:
            return False
        medstroms = r.startswith("med")
        return (medstroms and u == self.slutbrunn) or (r.startswith("mot") and u == self.startbrunn)

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

    # ---- åtgärdspaket (steg 3) ----
    def brunnstyp(self, brunn: str) -> str:
        """NB / TB / RB för en av sträckans brunnar: TVDAT-infokoden vid sträckans ände i första
        hand (inom BRUNNSTYP_ANDE_M från 0 resp. längden), annars litterats prefix, annars ''."""
        typ_film = self.brunnstyp_film(brunn)
        return typ_film or brunnstyp_ur_littera(brunn)

    def brunnstyp_film(self, brunn: str) -> str:
        """Brunnstyp enligt TVDAT-koden (NB/TB/RB) närmast sträckans ände, annars ''."""
        if self.langd <= 2 * BRUNNSTYP_ANDE_M:
            return ""
        ande = 0.0 if brunn == self.fran_brunn else self.langd
        nara = [o for o in self.observationer
                if o.infokod in ("NB", "TB", "RB") and abs(o.lage - ande) <= BRUNNSTYP_ANDE_M]
        return min(nara, key=lambda o: abs(o.lage - ande)).infokod if nara else ""

    def brunnstyp_konflikt(self) -> str:
        """Text när filmens kod och litterats prefix anger olika brunnstyp, t.ex.
        'KNBL62726: NB enligt littera, RB i filmen'."""
        ut = []
        for b in (self.startbrunn, self.slutbrunn):
            f, l = self.brunnstyp_film(b), brunnstyp_ur_littera(b)
            if f and l and f != l:
                ut.append(f"{b}: {l} enligt littera, {f} i filmen")
        return "; ".join(ut)

    @property
    def gallande_klass(self) -> str:
        """Manuell bedömning (A–E) om ifylld, annars maskinell klass."""
        m = manuell_klass(self.manuell_bedomning)
        return m or self.klass

    @property
    def metod_auto(self) -> str:
        """strumpa / schakt / ingen. Texten i Manuell bedömning styr alltid (schakt, strumpa, ingen);
        annars strumpa för gällande klass i ATGARD_KLASSER som inte är relinad, ingen för övriga.
        F3: inget automatiskt byte till schakt (om inte SCHAKT_AUTOMATISKT) – se atgardsflagga."""
        m = self.manuell_metod
        if m:
            return m
        if self.gallande_klass not in ATGARD_KLASSER or self.relinad:
            return "ingen"
        if SCHAKT_AUTOMATISKT and self.grad4_koder:
            return "schakt"
        return "strumpa"

    @property
    def manuell_metod(self) -> str:
        """Metod som texten i Manuell bedömning anger: schakt / strumpa / ingen, annars ''.
        'schakt' vinner över 'strumpa' ('strumpa ej möjlig, schakt'); 'ingen'/'inget'/'ej åtgärd' = ingen."""
        t = (self.manuell_bedomning or "").strip().lower()
        if "schakt" in t:
            return "schakt"
        if any(o in t for o in ("ingen", "inget", "ej åtgärd", "ej atgard", "avvakta")):
            return "ingen"
        if "strump" in t:
            return "strumpa"
        return ""

    @property
    def metod(self) -> str:
        return self._metod or self.metod_auto

    @property
    def grad4_koder(self) -> list[str]:
        """Konstruktionskoder med grad 4 som gör strumpning tveksam (GRAD4_KODER_A)."""
        return sorted({o.kod for o in self.skador("K")
                       if o.grad == 4 and (GRAD4_KODER_A is None or o.kod in GRAD4_KODER_A)})

    @property
    def lagningsbehov(self) -> str:
        """Skador som brukar kräva punktlagning före strumpning, t.ex. '2×YTS4, 1×RBR3'."""
        c = Counter((o.kod, o.grad) for o in self.skador("K") if (o.kod, o.grad) in LAGNINGSKODER)
        return ", ".join(f"{n}×{kod}{grad}" for (kod, grad), n in
                         sorted(c.items(), key=lambda kv: (-kv[0][1], kv[0][0])))

    @property
    def atgardsflagga(self) -> str:
        fl = []
        gk = self.gallande_klass
        if self.relinad and gk in ATGARD_KLASSER:
            fl.append(f"relinad men klass {gk} – fodret skadat?")
        if self.metod == "strumpa":
            if self.grad4_koder:
                fl.append(f"grad 4 {'/'.join(self.grad4_koder)} – går strumpa?")
            if self.lagning_m is None and self.lagningsbehov:
                fl.append(f"lagning? {self.lagningsbehov}")
            typer = (self.brunnstyp(self.startbrunn), self.brunnstyp(self.slutbrunn))
            if "" in typer:
                fl.append("brunnstyp okänd")
            elif "NB" not in typer:
                fl.append("bara tillsyns-/rensbrunnar – framschaktning")
            konflikt = self.brunnstyp_konflikt()
            if konflikt:
                fl.append(f"brunnstyp? {konflikt}")
        if self.kostnadsflagga:
            fl.append(self.kostnadsflagga)
        if self.etapp_flagga:
            fl.append(self.etapp_flagga)
        return "; ".join(fl)

    @property
    def dimension_mm(self) -> int | None:
        m = re.search(r"\d+", self.dimension or "")
        return int(m.group(0)) if m else None

    # --- GIS-data (gisdata.json) ---------------------------------------------------------
    def _gis_vg(self) -> tuple[float | None, float | None]:
        """GIS-vattengång vid start- och slutbrunn (flödesriktning) ur den kopplade ledningen."""
        g = self.gis and self.gis.get("ledning")
        if not g:
            return None, None
        fran, till = _normlittera(g.get("fran")), _normlittera(g.get("till"))
        if fran == _normlittera(self.startbrunn) or (not fran and till == _normlittera(self.slutbrunn)):
            return g.get("vg_fran"), g.get("vg_till")
        return g.get("vg_till"), g.get("vg_fran")

    @property
    def gis_lutning_promille(self) -> float | None:
        """Fall enligt GIS-vattengångarna, ‰ (positiv = fall i flödesriktningen)."""
        a, b = self._gis_vg()
        g = self.gis and self.gis.get("ledning")
        if a is None or b is None or not g or not g.get("langd_m"):
            return None
        return (a - b) / g["langd_m"] * 1000

    def _djup(self, vilken: str) -> float | None:
        """Djup vid brunnen = locknivå − lägsta vattengång bland ledningarna i brunnen (m)."""
        if not self.gis:
            return None
        b = self.gis.get("brunn_" + vilken)
        vg = self.gis.get("vg_min_" + vilken)
        if not b or b.get("lockniva") is None or vg is None:
            return None
        return b["lockniva"] - vg

    @property
    def djup_start(self) -> float | None:
        return self._djup("start")

    @property
    def djup_slut(self) -> float | None:
        return self._djup("slut")

    @property
    def anlaggningsar(self) -> int | None:
        g = self.gis and self.gis.get("ledning")
        return g.get("anlaggningsar") if g else None

    @property
    def gis_renoveringsar(self) -> int | None:
        """Renoveringsår (infodring) enligt GIS, om fältet exporterats."""
        g = self.gis and self.gis.get("ledning")
        return g.get("renoveringsar") if g else None

    @property
    def driftomrade(self) -> str:
        """Driftområde ur GIS-data: ledningens, annars startbrunnens, annars slutbrunnens."""
        if not self.gis:
            return ""
        for k in ("ledning", "brunn_start", "brunn_slut"):
            post = self.gis.get(k)
            if post and post.get("omrade"):
                return post["omrade"]
        return ""

    @property
    def gisflagga(self) -> str:
        """Avvikelser mellan filmen och GIS: brunn/ledning saknas, vattengång, riktning, material,
        dimension. Tom text när allt stämmer eller GIS-data saknas."""
        if "gisflagga" not in self._cache:
            self._cache["gisflagga"] = "; ".join(self._gisflaggor())
        return self._cache["gisflagga"]

    def _gisflaggor(self) -> list[str]:
        if not self.gis:
            return []
        fl = []
        saknas = [b for b, k, sida in ((self.startbrunn, "brunn_start", "start"), (self.slutbrunn, "brunn_slut", "slut"))
                  if not self.gis.get(k) and _normlittera(b) not in GIS_PLATSHALLARE and not self.grenror(sida)]
        if saknas:
            fl.append("brunn saknas i GIS: " + ", ".join(saknas))
        g = self.gis.get("ledning")
        gren = self.gis.get("grenror")
        if gren:
            kand = self.slutbrunn if gren["sida"] == "start" else self.startbrunn
            typ = self.grenror_typ(gren["sida"]) or "grenrör"
            if g:
                fl.append(dk(f"{typ}: GIS-ledningen från {kand} med fri ände ({g.get('langd_m') or 0:.1f} m) antagen"
                             + (f" – längden avviker från filmens {self.langd:.1f} m" if gren.get("langd_avviker") else "")))
            else:
                fl.append(f"{typ}: ingen ledning med fri ände från {kand} i GIS")
                return fl
        if not g:
            if not saknas:
                fl.append("ingen ledning i GIS mellan brunnarna")
            return fl
        if g.get("_syntetisk") == "fagelvag":
            fl.append(dk(f"ingen ledning i GIS mellan brunnarna – kartlängd fågelvägen {g['langd_m']:.1f} m"))
            return fl
        if g.get("_syntetisk") == "via":
            fl.append(dk(f"ingen direkt ledning i GIS – kartlängd {g['langd_m']:.1f} m via "
                         f"{', '.join(g['_via'])} ({len(g['_delar'])} ledningar)"))
        if g.get("delar") and len(g["delar"]) > 1 and len({(d.get("material"), d.get("dimension")) for d in g["delar"]}) > 1:
            fl.append(dk("GIS: ledningen består av " + " + ".join(
                f"{d.get('material') or '?'} {d.get('dimension') or ''} {d.get('langd_m') or 0:.1f} m".replace("  ", " ")
                for d in g["delar"]) + " (skarv utan brunn)"))
        if self.langd >= 1:
            # vattengång: fallet mellan brunnarna enligt filmen mot GIS
            gs, ge = self._gis_vg()
            zs, ze = self._filens_brunnshojder()
            if gs is not None and ge is not None:
                fall_gis = gs - ge
                if abs(fall_gis) > GIS_FALL_ORIMLIGT_M:
                    fl.append(dk(f"vattengång orimlig i GIS: {gs:.2f} / {ge:.2f} m"))
                elif fall_gis < -GIS_RIKTNING_MIN_M:
                    fl.append(dk(f"riktning: GIS-vattengången stiger {-fall_gis:.2f} m från {self.startbrunn} "
                                 f"till {self.slutbrunn}"))
                elif zs is not None and ze is not None and not self.ofullstandig:
                    fall_film = zs - ze
                    if abs(fall_film - fall_gis) > max(GIS_FALL_TOL_M, GIS_FALL_TOL_ANDEL * abs(fall_gis)):
                        fl.append(dk(f"vattengång: fall {fall_film:.2f} m i filmen, {fall_gis:.2f} m i GIS"))
            elif gs is None and ge is None:
                fl.append("vattengång saknas i GIS")
        # material och dimension
        mf, mg = _materialgrupp(self.material), _materialgrupp(g.get("material"))
        if mf and mg and mf != mg:
            fl.append(f"material: {self.material} i filmen, {g.get('material')} i GIS")
        df, dg = self.dimension_mm, g.get("dimension")
        if df and dg and df != dg:
            fl.append(f"dimension: {df} i filmen, {dg} i GIS")
        # renovering: GIS säger infodrad men filmen visar inget foder, eller tvärtom
        ren = g.get("renoveringsar")
        if ren and not self.relinad and self.langd >= 1:
            fl.append(f"renoverad {ren} enligt GIS men inte relinad enligt filmen")
        elif self.relinad and not ren and g.get("_ren_exporterad"):
            fl.append("relinad enligt filmen men inget renoveringsår i GIS")
        return fl


def _materialgrupp(text: str | None) -> str:
    """Materialnamn → grupp enligt GIS_MATERIAL ("BTG"/"Betong" → "Betong"); okänt namn ger sig
    självt med stor bokstav, tomt ger ''. Foderuppgifter inom parentes räknas inte."""
    t = re.sub(r"\(.*?\)", "", text or "").strip().upper()
    if not t:
        return ""
    for nyckel, grupp in GIS_MATERIAL.items():
        if t == nyckel or t.startswith(nyckel + " ") or t.startswith(nyckel + "-"):
            return grupp
    return t.capitalize()


def manuell_klass(text: str) -> str:
    """A–E ur Manuell bedömning: en ensam bokstav ("A", "b", "A – går strumpa"), inte första
    bokstaven i fri text ("Bevaka" är inte klass B). Tomt om ingen klass angetts."""
    m = re.match(r"\s*([A-Ea-e])(?![A-Za-zÅÄÖåäö])", text or "")
    return m.group(1).upper() if m else ""


def brunnstyp_ur_littera(littera: str) -> str:
    """NB / TB / RB ur litterats bokstavsprefix (KNBL, BdNBL → NB; KTB, STB → TB; SRB, KRB → RB),
    annars ''. Rensbrunn behandlas som tillsynsbrunn i kalkylen (F1) men redovisas som RB."""
    m = re.match(r"[A-Za-zÅÄÖåäö]+", (littera or "").strip())
    if not m:
        return ""
    pfx = m.group(0).upper()
    for typ in ("NB", "TB", "RB"):
        if typ in pfx:
            return typ
    return ""


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
    for enc in ("utf-8-sig", "cp1252"):
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            continue
    return raw.decode("latin-1")            # kan inte misslyckas


def _normlittera(t: str) -> str:
    return re.sub(r"[\s\-_]", "", t or "").upper()


LITTERA_KOLUMNER = {
    "fel": "fel", "felaktigt": "fel", "felaktig": "fel", "littera": "fel", "fran": "fel", "från": "fel",
    "gammalt": "fel", "gammal": "fel",
    "ratt": "ratt", "rätt": "ratt", "ny": "ratt", "nytt": "ratt", "till": "ratt", "ersattning": "ratt",
    "ersättning": "ratt",
    "fil": "fil", "tv3": "fil", "tv3_fil": "fil", "tv3-fil": "fil",
    "nr": "nr", "stracka": "nr", "sträcka": "nr", "stracknr": "nr", "sträcknr": "nr",
    "motbrunn": "motbrunn", "mot": "motbrunn", "partner": "motbrunn", "andra_brunn": "motbrunn",
    "andrabrunn": "motbrunn",
    "kommentar": "kommentar", "anm": "kommentar", "anmarkning": "kommentar", "anmärkning": "kommentar",
}


def _litterakolumn(rubrik: str) -> str:
    """Rubriktext -> kolumnnamn (fel/ratt/fil/nr/motbrunn/kommentar). Exakt träff först, annars
    första nyckelord som ingår i texten ('Felaktigt littera' -> fel, 'Rätt littera' -> ratt)."""
    r = rubrik.strip().lower()
    if r in LITTERA_KOLUMNER:
        return LITTERA_KOLUMNER[r]
    ord_ = re.findall(r"[a-zåäö0-9_]+", r)
    for o in ord_:
        if o in LITTERA_KOLUMNER and o != "littera":
            return LITTERA_KOLUMNER[o]
    return r


def las_littera(path: str) -> list[dict]:
    """Läser en CSV med ersättningslittera för brunnar som märkts fel vid filmningen.

    Kolumner (rubrikrad avgör ordningen; utan rubrik antas fel;ratt;kommentar):
        fel        felaktigt littera i TV3-filen (obligatorisk)
        ratt       rätt littera (obligatorisk)
        fil        begränsa till en TV3-fil (filnamn, med eller utan .TV3)         valfri
        nr         begränsa till ett sträcknummer i filen (som i fliken Prioritering) valfri
        motbrunn   begränsa till filmningar där den andra brunnen är denna           valfri
        kommentar  fritext                                                          valfri
    Rader utan fil/nr/motbrunn gäller överallt. Har flera rader samma felaktiga littera
    vinner den med flest villkor. Avgränsare ; , eller tab; tomma rader och #-rader hoppas
    över. Littera jämförs utan hänsyn till versaler, mellanslag, bindestreck och understreck."""
    regler: list[dict] = []
    kolumner: list[str] | None = None
    for rad in las_text(path).splitlines():
        rad = rad.strip()
        if not rad or rad.startswith("#"):
            continue
        delar = [d.strip().strip('"').strip("'") for d in re.split(r"[;,\t]", rad)]
        if kolumner is None:
            kolumner = ["fel", "ratt", "kommentar"]
            rubriker = [_litterakolumn(d) for d in delar]
            if "fel" in rubriker or "ratt" in rubriker:       # rubrikrad
                if "fel" not in rubriker or "ratt" not in rubriker:
                    print(f"  VARNING {os.path.basename(path)}: rubrikraden saknar kolumnen "
                          f"{'fel' if 'fel' not in rubriker else 'ratt'} – antar ordningen fel;ratt;kommentar")
                else:
                    kolumner = rubriker
                continue
        post = {k: (delar[i] if i < len(delar) else "") for i, k in enumerate(kolumner)}
        if not post.get("fel") or not post.get("ratt"):
            continue
        regel = {
            "fel": _normlittera(post["fel"]),
            "ratt": post["ratt"],
            "fil": os.path.splitext(os.path.basename(post.get("fil", "")))[0].strip().lower(),
            "nr": _i(post.get("nr", "")) if post.get("nr", "").strip() else None,
            "motbrunn": _normlittera(post.get("motbrunn", "")),
            "kommentar": post.get("kommentar", ""),
        }
        regel["villkor"] = sum(1 for k in ("fil", "nr", "motbrunn") if regel[k])
        regler.append(regel)
    regler.sort(key=lambda r: -r["villkor"])            # mest specifika först
    return regler


def _andtyp(ratt: str) -> str | None:
    """Ord i litterafilens ratt-kolumn som betyder "ingen brunn": grenrör, servis eller ände
    (propp, fri ände). Returnerar typen som visas i flaggan, annars None."""
    t = _normlittera(ratt)
    if re.fullmatch(r"GRENR[ÖO]R|P[ÅA]STICK", t):
        return "grenrör"
    if re.fullmatch(r"SERVIS(ANSLUTNING)?", t):
        return "servis"
    if re.fullmatch(r"(FRI)?[ÄA]NDE|PROPP|INGENBRUNN", t):
        return "ände utan brunn"
    return None


def ratta_littera(strackor: list[Stracka], regler: list[dict]) -> int:
    """Byter ut brunnslittera enligt reglerna i start-, slut- och utgångsbrunn.
    Returnerar antal sträckor som ändrats; ändringen noteras i Stracka.littera_rattat."""
    n = 0
    for s in strackor:
        fil = os.path.splitext(os.path.basename(s.fil))[0].strip().lower()
        par = {_normlittera(s.startbrunn), _normlittera(s.slutbrunn)}
        andringar: list[str] = []
        for falt in ("startbrunn", "slutbrunn", "utgangsbrunn"):
            v = getattr(s, falt)
            if not v:
                continue
            nv = _normlittera(v)
            for r in regler:
                if r["fel"] != nv:
                    continue
                if r["fil"] and r["fil"] != fil:
                    continue
                if r["nr"] is not None and r["nr"] != s.nr:
                    continue
                if r["motbrunn"] and (r["motbrunn"] not in par or r["motbrunn"] == nv):
                    continue
                typ = _andtyp(r["ratt"])
                if typ:
                    # ingen brunn utan grenrör/servis/ände – litterat behålls, änden kopplas i GIS
                    # till ledningen från den kända brunnen som slutar i en fri ände (se koppla_gis)
                    if falt in ("startbrunn", "slutbrunn"):
                        s.grenror_markerat["start" if falt == "startbrunn" else "slut"] = typ
                        andringar.append(f"{v} = {typ}")
                elif r["ratt"] != v:
                    andringar.append(f"{v}→{r['ratt']}")
                    setattr(s, falt, r["ratt"])
                break                                       # första (mest specifika) träffen gäller
        if andringar:
            s.littera_rattat = ", ".join(dict.fromkeys(andringar))
            n += 1
    return n


def kontrollera_hojder(s: Stracka) -> str:
    """Rimlighetskontroll av filens höjder. Skiljer brunnshöjderna i PROFILADM mer än HOJD_ORIMLIG_M,
    eller spänner inklinometerprofilen över mer än så, tas höjderna bort (lutning, svacka och
    profilbild blir okända) och orsaken returneras som text; annars ""."""
    fel = []
    zs, ze = s.profil_start_z, s.profil_slut_z
    if zs is not None and ze is not None and abs(zs - ze) > HOJD_ORIMLIG_M:
        fel.append(f"brunnshöjder {zs:.2f} / {ze:.2f} m")
        s.profil_start_z = s.profil_slut_z = None
    if s.profil:
        z = [p[2] for p in s.profil]
        if max(z) - min(z) > HOJD_ORIMLIG_M:
            fel.append(f"inklinometerprofil {min(z):.1f} till {max(z):.1f} m")
            s.profil = []
    return ("höjdfel i filen: " + ", ".join(fel)) if fel else ""


def las_tv3(path: str, littera: list[dict] | None = None) -> list[Stracka]:
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

    def _hojd(v):
        z = _f(v)
        return None if z is None or z <= HOJD_SAKNAS_UNDER else z

    for r in sektioner.get("PROFILADM", []):
        nr = _i(g(r, 0))
        if nr in strackor:
            strackor[nr].profil_start_z = _hojd(g(r, 16))
            strackor[nr].profil_slut_z = _hojd(g(r, 17))
    for r in sektioner.get("PROFILDAT", []):
        nr = _i(g(r, 0))
        if nr in strackor:
            x, rz, z = _f(g(r, 1)), _f(g(r, 2)), _hojd(g(r, 3))
            if x is not None and z is not None:
                strackor[nr].profil.append((x, rz or 0.0, z))

    hojdfel = []
    for s in strackor.values():
        s.profil.sort(key=lambda p: p[0])
        s.hojdfel = kontrollera_hojder(s)
        if s.hojdfel:
            hojdfel.append(s)
    if hojdfel:
        print(f"  OBS     {filnamn}: {len(hojdfel)} sträckor har orimliga höjder i filen – lutning/profil redovisas "
              f"som okända (nr {', '.join(str(s.nr) for s in hojdfel[:8])}" + (" …" if len(hojdfel) > 8 else "") + ")")

    if littera:
        ratta_littera(list(strackor.values()), littera)

    # Startbrunn ska vara uppströms. Om Riktning och Utgångsbrunn visar att filen i stället
    # anger brunnarna i kamerans riktning byts de (inkl. brunnshöjderna från PROFILADM).
    omvanda = [s for s in strackor.values() if s.brunnar_omvanda]
    for s in omvanda:
        s.startbrunn, s.slutbrunn = s.slutbrunn, s.startbrunn
        s.profil_start_z, s.profil_slut_z = s.profil_slut_z, s.profil_start_z
    if omvanda:
        print(f"  OBS     {filnamn}: {len(omvanda)} sträckor har start-/slutbrunn i kamerans riktning "
              f"– bytta så att startbrunn är uppströms (nr {', '.join(str(s.nr) for s in omvanda[:8])}"
              + (" …" if len(omvanda) > 8 else "") + ")")

    par = Counter(frozenset((_normlittera(s.startbrunn), _normlittera(s.slutbrunn))) for s in strackor.values())
    for s in strackor.values():
        s.flerinspekterad = par[frozenset((_normlittera(s.startbrunn), _normlittera(s.slutbrunn)))] > 1

    return [strackor[k] for k in sorted(strackor)]


# ----------------------------------------------------------------------------
# Media (video + bilder)
# ----------------------------------------------------------------------------

VIDEO_ANDELSER = {".mp4", ".mp2", ".mpg", ".mpeg", ".avi", ".wmv", ".mov", ".mkv", ".m4v", ".asf", ".ts",
                  ".mts", ".webm", ".vob", ".divx", ".flv", ".3gp"}
MEDIA_ANDELSER = VIDEO_ANDELSER | {".jpg", ".jpeg", ".png", ".bmp", ".gif", ".tif", ".tiff"}
_media_cache: dict[str, dict[str, str]] = {}


def indexera_katalog(katalog: str) -> dict[str, str]:
    """Filnamn (gemener) -> absolut sökväg för alla media-filer under katalogen (rekursivt)."""
    katalog = os.path.abspath(katalog)
    if katalog in _media_cache:
        return _media_cache[katalog]
    idx: dict[str, str] = {}
    if os.path.isdir(katalog):
        for rot, kataloger, namn in os.walk(katalog):
            kataloger.sort()                      # samma ordning oavsett filsystem
            for n in sorted(namn):
                stam, and_ = os.path.splitext(n)
                if and_.lower() in MEDIA_ANDELSER:
                    idx.setdefault(n.lower(), os.path.join(rot, n))          # första träffen vinner
                    if and_.lower() in VIDEO_ANDELSER:
                        idx.setdefault("\x00" + stam.lower(), os.path.join(rot, n))   # video utan ändelse
    _media_cache[katalog] = idx
    return idx


def koppla_media(strackor: list[Stracka], extra_kataloger: list[str],
                 extra_bildkataloger: list[str] | None = None) -> tuple[int, int, int, int]:
    """Letar upp video- och bildfiler. Filmer söks i TV3-filens katalog först, sedan i filens
    egna mediakataloger (från listfilen) och sist i de globala (media: i listfilen eller --media).
    Bilder söks i bildkatalogerna (bild: i listfilen eller --bilder, per fil eller globalt) om
    några angetts, annars i samma kataloger som filmerna. TV3-filens egen katalog söks alltid.
    Returnerar (videor hittade, videor totalt, bilder hittade, bilder totalt)."""
    vh = vt = bh = bt = 0
    saknade_video: dict[str, list[str]] = {}
    kataloger_per_fil: dict[str, list[str]] = {}
    for s in strackor:
        egen = os.path.dirname(s.tv3_sokvag)
        kataloger = [egen] + list(s.media_kataloger) + list(extra_kataloger)
        bildkat = list(s.bild_kataloger) + list(extra_bildkataloger or [])
        bildkataloger = bildkat + [egen] if bildkat else kataloger
        kataloger_per_fil.setdefault(s.fil, kataloger)
        index = [indexera_katalog(k) for k in kataloger]
        bildindex = [indexera_katalog(k) for k in bildkataloger]

        def hitta(namn: str, video: bool = False) -> str | None:
            n = os.path.basename(namn.strip().replace("\\", "/")).lower()   # TV3 kan ha sökväg i namnet
            if not n:
                return None
            for idx in (index if video else bildindex):
                if n in idx:
                    return idx[n]
            if video:                                 # annan filändelse på disk (.mpg vs .mp4)
                stam = "\x00" + os.path.splitext(n)[0]
                for idx in index:
                    if stam in idx:
                        return idx[stam]
            return None

        if s.videofil:
            vt += 1
            s.video_sokvag = hitta(s.videofil, video=True)
            vh += s.video_sokvag is not None
            if s.video_sokvag is None:
                saknade_video.setdefault(s.fil, []).append(s.videofil)
        for o in s.observationer:
            if o.bild:
                bt += 1
                o.bild_sokvag = hitta(o.bild)
                bh += o.bild_sokvag is not None
            if o.bild_b:
                bt += 1
                o.bild_b_sokvag = hitta(o.bild_b)
                bh += o.bild_b_sokvag is not None

    # Diagnostik: vilka filmnamn saknas, och vad heter filmerna som faktiskt finns i mapparna?
    for fil, namn in saknade_video.items():
        print(f"  {fil}: {len(namn)} videofiler saknas, t.ex. {', '.join(namn[:3])}")
        exempel: list[str] = []
        for k in kataloger_per_fil.get(fil, []):
            exempel += [os.path.basename(p) for nyckel, p in indexera_katalog(k).items()
                        if not nyckel.startswith("\x00") and os.path.splitext(nyckel)[1] in VIDEO_ANDELSER]
        exempel = sorted(dict.fromkeys(exempel))
        if exempel:
            print(f"    filmer i mapparna heter t.ex. {', '.join(exempel[:3])} ({len(exempel)} st)")
        else:
            print("    inga videofiler alls i de sökta mapparna: "
                  + ", ".join(kataloger_per_fil.get(fil, [])))
    return vh, vt, bh, bt


def las_markprofil(path: str) -> dict:
    """Läser markprofil.json från ArcMap-verktyget Markprofil. Returnerar {"poster": [...],
    "hojdsystem": str, "fil": path}."""
    import json
    with open(path, encoding="utf-8") as f:
        data = json.load(f)
    return {"poster": data.get("strackor", []), "hojdsystem": data.get("hojdsystem", ""), "fil": path}


def koppla_markprofil(strackor: list[Stracka], filer: list[dict]) -> tuple[int, int]:
    """Kopplar markhöjder och GIS-vattengång till sträckorna: på TV3-fil + sträcknummer i
    första hand, annars på brunnspar. Returnerar (sträckor med markhöjder, med GIS-vattengång)."""
    pa_nr: dict[tuple[str, int], dict] = {}
    pa_filpar: dict[tuple[str, frozenset], dict] = {}
    pa_par: dict[frozenset, dict] = {}
    system: dict[int, str] = {}
    for mf in filer:
        for post in mf["poster"]:
            system[id(post)] = mf["hojdsystem"] or "RH2000"
            fil = os.path.basename(str(post.get("fil") or "")).lower()
            par = frozenset((_normlittera(post.get("startbrunn")), _normlittera(post.get("slutbrunn"))))
            if post.get("nr") is not None and fil:
                pa_nr.setdefault((fil, int(post["nr"])), post)
            pa_filpar.setdefault((fil, par), post)
            pa_par.setdefault(par, post)
    n_mark = n_vg = 0
    for s in strackor:
        fil = os.path.basename(s.fil).lower()
        par = frozenset((_normlittera(s.startbrunn), _normlittera(s.slutbrunn)))
        post = pa_nr.get((fil, s.nr)) or pa_filpar.get((fil, par)) or pa_par.get(par)
        if not post:
            continue
        s.mark = [(float(m), float(z)) for m, z, *_ in post.get("mark", []) if z is not None]
        s.gis_vg_start = post.get("vg_start")
        s.gis_vg_slut = post.get("vg_slut")
        s.langd_karta = post.get("langd_karta_m")
        s.hojdsystem = system[id(post)]
        s._cache.clear()
        n_mark += bool(s.mark)
        n_vg += s.gis_vg_start is not None and s.gis_vg_slut is not None
    return n_mark, n_vg


def las_gis(path: str) -> dict:
    """Läser gisdata.json från ArcMap-verktyget Exportera GIS-data. Returnerar {"brunnar": {littera:
    post}, "ledningar": [post], "hojdsystem", "fil"}; littera normaliseras som i _normlittera."""
    import json
    try:
        with open(path, encoding="utf-8") as f:
            data = json.load(f)
    except (OSError, ValueError) as e:
        raise ValueError(f"{path}: kan inte läsas som gisdata.json ({e})") from e

    def tal(v):
        try:
            v = float(str(v).replace(",", "."))
        except (TypeError, ValueError):
            return None
        return None if v <= GIS_SAKNAS_UNDER else v

    def heltal(v):
        m = re.search(r"\d{3,4}", str(v or ""))
        return int(m.group(0)) if m else None
    brunnar = {}
    for b in data.get("brunnar", []):
        n = _normlittera(b.get("littera"))
        b["lockniva"] = tal(b.get("lockniva"))
        if n and n not in brunnar:
            brunnar[n] = b
    ledningar = []
    har_ren = bool((data.get("lager") or {}).get("renoveringsar"))     # fältet valt i exporten?
    for led in data.get("ledningar", []):
        led["vg_fran"], led["vg_till"] = tal(led.get("vg_fran")), tal(led.get("vg_till"))
        led["langd_m"] = tal(led.get("langd_m"))
        led["dimension"] = heltal(led.get("dimension")) if led.get("dimension") is not None else None
        led["anlaggningsar"] = heltal(led.get("anlaggningsar"))
        led["renoveringsar"] = heltal(led.get("renoveringsar")) if har_ren else None
        led["_ren_exporterad"] = har_ren
        ledningar.append(led)
    tol = GIS_SKARV_TOL_M if GIS_SKARV_TOL_M is not None else (tal(data.get("tolerans_m")) or 1.0)
    n_fog = _sammanfoga_fria_andar(ledningar, tol)
    if n_fog:
        print(f"  {os.path.basename(path)}: {n_fog} ledningar hopfogade av delar som möts i en skarv utan brunn")
    return {"brunnar": brunnar, "ledningar": ledningar,
            "hojdsystem": data.get("hojdsystem", "") or "RH2000", "fil": path}


def _sammanfoga_fria_andar(ledningar: list[dict], tolerans: float) -> int:
    """Ledningsdelar som slutar i en fri ände (ingen brunn) där en annan dels fria ände ligger inom
    toleransen – t.ex. materialbyte mitt på sträckan – fogas ihop till en ledning mellan brunnarna.
    Bara parvisa möten fogas (tre ändar på samma ställe är en förgrening). Den hopfogade posten
    ersätter delarna i listan: längd = summan, vattengång ur ändarna, dimension/material/typ/år bara
    om alla delar är lika, `delar` = [{langd_m, material, dimension, anlaggningsar}], `skarvar` = [xy].
    Returnerar antalet hopfogade ledningar."""
    def namn(led, sida):
        return _normlittera(led.get(sida))

    def xy(led, sida):
        v = led.get(sida + "_xy")
        return tuple(v) if isinstance(v, (list, tuple)) and len(v) == 2 and all(isinstance(c, (int, float)) for c in v) else None

    def andra(sida):
        return "till" if sida == "fran" else "fran"

    cell = max(tolerans, 0.01)
    rut: dict[tuple[int, int], list] = {}
    fria = []
    for led in ledningar:
        for sida in ("fran", "till"):
            if not namn(led, sida) and xy(led, sida):
                x, y = xy(led, sida)
                rut.setdefault((round(x / cell), round(y / cell)), []).append((led, sida))
                fria.append((led, sida))
    partner: dict[tuple[int, str], tuple] = {}
    for led, sida in fria:
        x, y = xy(led, sida)
        cx, cy = round(x / cell), round(y / cell)
        nara = [(l2, s2) for dx in (-1, 0, 1) for dy in (-1, 0, 1) for l2, s2 in rut.get((cx + dx, cy + dy), [])
                if l2 is not led and math.hypot(xy(l2, s2)[0] - x, xy(l2, s2)[1] - y) <= tolerans]
        if len(nara) == 1:
            partner[(id(led), sida)] = nara[0]
    # bara ömsesidiga par
    per_id = {id(l): l for l in ledningar}
    partner = {k: v for k, v in partner.items() if partner.get((id(v[0]), v[1])) == (per_id[k[0]], k[1])}
    anv: set[int] = set()
    nya = []
    for led in ledningar:
        if id(led) in anv:
            continue
        for brunnsida in ("fran", "till"):
            if not namn(led, brunnsida) or (id(led), andra(brunnsida)) not in partner:
                continue
            kedja = [(led, brunnsida)]                 # (del, sidan som vetter mot föregående/brunnen)
            cur, in_sida = led, brunnsida
            klar = False
            while True:
                p = partner.get((id(cur), andra(in_sida)))
                if not p or id(p[0]) in {id(d) for d, _ in kedja} or id(p[0]) in anv:
                    break
                cur, in_sida = p
                kedja.append((cur, in_sida))
                if namn(cur, andra(in_sida)):
                    klar = True
                    break
            if not klar:
                continue
            forsta, fs = kedja[0]
            sista, ss = kedja[-1]

            def gemensamt(nyckel):
                v = {d.get(nyckel) for d, _ in kedja}
                return v.pop() if len(v) == 1 else None
            post = dict(forsta)
            post.update({
                "fran": forsta.get(fs), "till": sista.get(andra(ss)),
                "fran_xy": forsta.get(fs + "_xy"), "till_xy": sista.get(andra(ss) + "_xy"),
                "langd_m": sum(d.get("langd_m") or 0.0 for d, _ in kedja),
                "vg_fran": forsta.get("vg_" + fs), "vg_till": sista.get("vg_" + andra(ss)),
                "dimension": gemensamt("dimension"), "material": gemensamt("material"),
                "ledningstyp": gemensamt("ledningstyp"), "anlaggningsar": gemensamt("anlaggningsar"),
                "renoveringsar": gemensamt("renoveringsar"),
                "omrade": next((d.get("omrade") for d, _ in kedja if d.get("omrade")), None),
                "antal_delar": len(kedja),
                "delar": [{"langd_m": d.get("langd_m"), "material": d.get("material"), "dimension": d.get("dimension"),
                           "anlaggningsar": d.get("anlaggningsar"), "oid": d.get("oid")} for d, _ in kedja],
                "skarvar": [d.get(andra(s) + "_xy") for d, s in kedja[:-1]],
            })
            nya.append(post)
            anv.update(id(d) for d, _ in kedja)
            break
    if nya:
        ledningar[:] = [l for l in ledningar if id(l) not in anv] + nya
    return len(nya)


def _gis_vag(ns: str, ne: str, grannar: dict, langd: float) -> dict | None:
    """Kortaste vägen i GIS från ns till ne via andra brunnar (högst GIS_VAG_MAX_HOPP ledningar, högst
    GIS_VAG_MAX_ANDEL × filmens längd + 20 m). Returnerar en syntetisk ledningspost {fran, till, langd_m,
    vg_fran, vg_till, dimension, material, ledningstyp, omrade, _syntetisk: "via", _via: [brunnar],
    _delar: [ledningar]} eller None."""
    import heapq
    tak = GIS_VAG_MAX_ANDEL * max(langd, 1.0) + 20
    basta: dict[str, float] = {ns: 0.0}
    ko = [(0.0, 0, ns, [])]                 # (längd, hopp, brunn, [(granne, ledning) …])
    while ko:
        L, hopp, n, vag = heapq.heappop(ko)
        if n == ne:
            delar = [led for _, led in vag]
            via = [b for b, _ in vag[:-1]]
            def _riktad(led, fran):          # (vg vid fran, vg vid andra änden)
                if _normlittera(led.get("fran")) == fran:
                    return led.get("vg_fran"), led.get("vg_till")
                return led.get("vg_till"), led.get("vg_fran")
            vg0 = _riktad(delar[0], ns)[0]
            vg1 = _riktad(delar[-1], via[-1] if via else ns)[1]
            def _gemensamt(nyckel):
                v = {d.get(nyckel) for d in delar}
                return v.pop() if len(v) == 1 else None
            return {"fran": ns, "till": ne, "langd_m": L, "vg_fran": vg0, "vg_till": vg1,
                    "dimension": _gemensamt("dimension"), "material": _gemensamt("material"),
                    "ledningstyp": _gemensamt("ledningstyp"), "anlaggningsar": _gemensamt("anlaggningsar"),
                    "renoveringsar": _gemensamt("renoveringsar"),
                    "_ren_exporterad": any(d.get("_ren_exporterad") for d in delar),
                    "omrade": delar[0].get("omrade"), "_syntetisk": "via", "_via": via, "_delar": delar}
        if hopp >= GIS_VAG_MAX_HOPP:
            continue
        for granne, led in grannar.get(n, []):
            L2 = L + (led.get("langd_m") or 0.0)
            if L2 > tak or granne == ns or (granne in basta and basta[granne] <= L2):
                continue
            basta[granne] = L2
            heapq.heappush(ko, (L2, hopp + 1, granne, vag + [(granne, led)]))
    return None


def koppla_gis(strackor: list[Stracka], filer: list[dict]) -> dict:
    """Kopplar GIS-data till sträckorna på brunnspar (båda riktningarna). Finns flera ledningar
    mellan samma brunnar tas den vars längd ligger närmast filmens. Utan markprofil används
    GIS-vattengång och kartlängd även till höjdanpassningen. Returnerar statistik + lista
    över brunnar i TV3-filerna som saknas i GIS med förslag på rätt littera."""
    import difflib
    brunnar: dict[str, dict] = {}
    pa_par: dict[frozenset, list[dict]] = {}
    fria: dict[str, list[dict]] = {}
    vg_min: dict[str, float] = {}
    system = "RH2000"
    sedda_led: set = set()
    for gf in filer:
        system = gf["hojdsystem"] or system
        for n, b in gf["brunnar"].items():
            brunnar.setdefault(n, b)
        for led in gf["ledningar"]:
            a, b = _normlittera(led.get("fran")), _normlittera(led.get("till"))
            # samma ledning i flera exportfiler (t.ex. gränsledningar vid export per område) räknas en gång
            nyckel = ((led.get("lager"), led.get("oid"), led.get("del")) if led.get("oid") is not None
                      else (frozenset((a, b)), round(led.get("langd_m") or 0, 1)))
            if nyckel in sedda_led:
                led["_dubblett"] = True
                continue
            sedda_led.add(nyckel)
            if a and b:
                pa_par.setdefault(frozenset((a, b)), []).append(led)
            elif a or b:
                fria.setdefault(a or b, []).append(led)       # brunn -> ledningar som slutar i en fri ände
            for n, vg in ((a, led.get("vg_fran")), (b, led.get("vg_till"))):
                if n and vg is not None and (n not in vg_min or vg < vg_min[n]):
                    vg_min[n] = vg
    grannar: dict[str, list[tuple[str, dict]]] = {}      # brunn -> [(granne, ledning)] i GIS
    for par, leds in pa_par.items():
        a, b = tuple(par) if len(par) == 2 else (next(iter(par)),) * 2
        for led in leds:
            grannar.setdefault(a, []).append((b, led))
            grannar.setdefault(b, []).append((a, led))
    n_led = n_vg = 0
    okanda: list[dict] = []
    for s in strackor:
        ns, ne = _normlittera(s.startbrunn), _normlittera(s.slutbrunn)
        kand = pa_par.get(frozenset((ns, ne)), [])
        # vid flera ledningar mellan samma brunnar: samma ledningstyp först (S/D parallellt i samma
        # schakt har nästan samma längd), sedan längd närmast filmens
        typ = (s.ledningstyp or "")[:1].upper()
        led = min(kand, key=lambda l: ((l.get("ledningstyp") or "")[:1].upper() != typ if typ else False,
                                       abs((l.get("langd_m") or 0) - s.langd))) if kand else None
        grenror = None
        if led is None and s.grenror("start") != s.grenror("slut"):
            # anslutning till grenrör: ingen brunn i den änden – ta GIS-ledningen från den kända
            # brunnen som slutar i en fri ände, den vars längd ligger närmast filmens
            sida = "start" if s.grenror("start") else "slut"
            kand_n = ne if sida == "start" else ns
            grenror = {"sida": sida}
            if kand_n in brunnar:
                kand = sorted(fria.get(kand_n, []), key=lambda l: abs((l.get("langd_m") or 0) - s.langd))
                tol = GIS_GRENROR_LANGD_TOL * s.langd + 3.0
                if kand and abs((kand[0].get("langd_m") or 0) - s.langd) <= tol:
                    led = kand[0]
                elif len(kand) == 1:
                    led = kand[0]
                    grenror["langd_avviker"] = True
        if led is None and ns in brunnar and ne in brunnar and ns != ne:
            # ingen direkt ledning: väg via andra brunnar, annars brunnarnas avstånd fågelvägen
            led = _gis_vag(ns, ne, grannar, s.langd)
            if led is None:
                b0, b1 = brunnar[ns], brunnar[ne]
                if all(isinstance(b.get(k), (int, float)) for b in (b0, b1) for k in ("x", "y")):
                    d = math.hypot(b0["x"] - b1["x"], b0["y"] - b1["y"])
                    if d >= 1.0:                     # brunnar utan koordinater (0, 0) ger ingen kartlängd
                        led = {"fran": ns, "till": ne, "langd_m": d, "_syntetisk": "fagelvag"}
        s.gis = {"ledning": led, "brunn_start": brunnar.get(ns), "brunn_slut": brunnar.get(ne),
                 "vg_min_start": vg_min.get(ns), "vg_min_slut": vg_min.get(ne), "fil": filer[0]["fil"],
                 "grenror": grenror}
        for b, n, annan, n_annan, sida in ((s.startbrunn, ns, s.slutbrunn, ne, "start"),
                                           (s.slutbrunn, ne, s.startbrunn, ns, "slut")):
            if n and n not in brunnar and n not in GIS_PLATSHALLARE and not s.grenror(sida):
                okanda.append({"littera": b, "fil": s.fil, "nr": s.nr, "langd": s.langd,
                               "motbrunn": annan if n_annan in brunnar else "", "n_mot": n_annan})
        if led:
            n_led += 1
            if s.langd_karta is None:        # ingen markprofil – GIS-exporten ger nivåer och kartlängd
                a, b = s._gis_vg()
                s.gis_vg_start, s.gis_vg_slut = a, b
                s.langd_karta = led.get("langd_m")
                s.hojdsystem = system
                n_vg += a is not None and b is not None
        s._cache.clear()
    # Förslag på rätt littera för okända brunnar, per sträcka. Är den andra brunnen känd prövas
    # dess grannar i GIS: en granne vars ledning har ungefär filmens längd (GIS_LANGD_TOL) är
    # troligen rätt brunn, även om namnet inte liknar. Annars (båda okända) bara liknande namn.
    alla = {n: b.get("littera") for n, b in brunnar.items()}
    for o in okanda:
        nn = _normlittera(o["littera"])
        kand = []                   # (poäng, littera, text)
        if o["motbrunn"] and o["langd"] >= 1:
            for n_granne, led in grannar.get(o["n_mot"], []):
                L = led.get("langd_m")
                if not L:
                    continue
                diff = abs(L - o["langd"]) / max(o["langd"], 1.0)
                likhet = difflib.SequenceMatcher(None, nn, n_granne).ratio()
                if diff <= GIS_LANGD_TOL:
                    kand.append((diff - 0.5 * likhet, alla.get(n_granne, n_granne),
                                 dk(f"ledning {L:.1f} m mot {o['langd']:.1f} m i filmen"
                                    + (f", namnlikhet {likhet:.2f}" if likhet >= GIS_LITTERA_LIKHET else ""))))
            o["metod"] = "längd från " + o["motbrunn"]
        if not kand:
            for n_lik in difflib.get_close_matches(nn, list(alla), n=3, cutoff=GIS_LITTERA_LIKHET):
                likhet = difflib.SequenceMatcher(None, nn, n_lik).ratio()
                kand.append((1 - likhet, alla.get(n_lik, n_lik), dk(f"liknande namn ({likhet:.2f})")))
            o["metod"] = "liknande namn" if kand else "inget förslag"
        kand.sort()
        sedda = set()
        o["forslag"] = [(lit, txt) for _, lit, txt in kand if not (lit in sedda or sedda.add(lit))][:3]
        # 'ratt' förifylls bara när längden pekar ut en tydlig kandidat (ensam, eller klart bättre än nästa)
        o["sakert"] = bool(kand) and o["metod"].startswith("längd") and (
            len(kand) == 1 or kand[1][0] - kand[0][0] >= 0.05)
    okanda.sort(key=lambda o: (o["littera"], os.path.basename(o["fil"]), o["nr"]))
    return {"ledningar": n_led, "vg": n_vg, "okanda": okanda, "brunnar_i_gis": len(brunnar),
            "flaggade": sum(1 for s in strackor if s.gisflagga)}


def delfilm(s: Stracka) -> bool:
    """Sträckan är en delfilm som ingår i en sammanslagen sträcka."""
    return s.filmstatus.startswith("ingår")


def visade(strackor: list[Stracka]) -> list[Stracka]:
    """Sträckor som redovisas var för sig (Prioritering, Observationer, protokoll): alla utom delfilmer,
    om inte DELFILMER_SEPARAT."""
    return [s for s in strackor if DELFILMER_SEPARAT or not delfilm(s)]


def aktiva(strackor: list[Stracka]) -> list[Stracka]:
    """Sträckor som räknas: inte ersatta av en nyare film och inte delfilmer i en sammanslagning."""
    return [s for s in strackor if not (s.filmstatus.startswith("ersatt") or s.filmstatus.startswith("ingår"))]


def _spegla_klocka(k: str) -> str:
    """Klockposition sedd från andra hållet: kl 3 → 9, 12 och 6 oförändrade."""
    m = re.match(r"\s*(\d{1,2})", k or "")
    if not m:
        return k
    v = int(m.group(1)) % 12
    return f"{(12 - v) % 12 or 12:02d}" if v else "12"


def _tidsnyckel(s: Stracka) -> tuple:
    """Sorteringsnyckel för när en film gjordes: datum + klockslag tolkade som tid (ISO, svensk
    punkt/snedstreck eller 'ÅÅÅÅMMDD'), annars strängarna som de är."""
    from datetime import datetime
    d = (s.datum or "").strip()
    k = (s.klockslag or "").strip()
    for fmt in ("%Y-%m-%d", "%Y%m%d", "%d.%m.%Y", "%d/%m/%Y", "%Y/%m/%d", "%d-%m-%Y"):
        try:
            dag = datetime.strptime(d, fmt)
            break
        except ValueError:
            continue
    else:
        return (0, d, k)
    tid = 0
    m = re.match(r"(\d{1,2})[:.](\d{2})", k)
    if m:
        tid = int(m.group(1)) * 60 + int(m.group(2))
    return (1, dag.strftime("%Y-%m-%d"), tid)


def _sammanslagen(a: Stracka, b: Stracka, nr: int) -> Stracka:
    """Slår ihop film a (från startbrunnen) och film b (från slutbrunnen) till en sträcka med
    positioner från startbrunnen. Kartlängden ger totallängd och överlapp; i överlappet behålls den
    nyare filmens observationer och den äldre filmens löpande skador kapas vid snittet. Utan
    kartlängd antas filmerna mötas utan överlapp, och ett hinder som båda kamerorna stannade vid
    räknas en gång. Varje observation får 'film nr X' i kommentaren."""
    import copy
    import dataclasses
    L_karta = a.langd_karta or b.langd_karta
    noter = []
    if L_karta and L_karta < max(a.langd, b.langd) - 0.5:
        noter.append(dk(f"kartlängd {L_karta:.1f} m orimlig (kortare än en delfilm) – används inte"))
        L_karta = None
    L = L_karta if L_karta else a.langd + b.langd
    overlapp = a.langd + b.langd - L if L_karta else 0.0
    nyare_b = _tidsnyckel(b) >= _tidsnyckel(a)
    nyare, aldre = (b, a) if nyare_b else (a, b)

    def marka(o, film):
        o.film_nr = film.nr
        o.kommentar = f"film nr {film.nr}" + (f" – {o.kommentar}" if o.kommentar else "")

    # Snitt i respektive films egen axel: den äldre filmens observationer bortom snittet tas bort
    # och dess löpande skador kapas där (den nyare filmen gäller i överlappet)
    snitt_a = a.langd - overlapp if (overlapp > 0 and nyare_b) else None
    snitt_b = b.langd - overlapp if (overlapp > 0 and not nyare_b) else None
    obs = []
    for o in a.observationer:
        if o.kod == "KAM":
            continue
        if snitt_a is not None and o.lage > snitt_a + 1e-6:
            continue
        o2 = copy.copy(o)
        o2.nr = nr
        if snitt_a is not None and o.lopande.startswith("A") and o.lopande_langd and o.lage + o.lopande_langd > snitt_a:
            o2.lopande_langd = round(max(0.0, snitt_a - o.lage), 2)
        marka(o2, a)
        obs.append(o2)
    # b: löpande skadors längd per löpnummer (för B-radernas nya läge och öppna skador)
    lop_langd = {o.lopande[1:]: o.lopande_langd for o in b.observationer if o.lopande.startswith("A")}
    a_obs = [o for o in obs]
    for o in b.observationer:
        if o.kod == "KAM":
            continue
        if snitt_b is not None and o.lage > snitt_b + 1e-6:
            continue
        o2 = copy.copy(o)
        o2.nr = nr
        if o.lopande.startswith("A"):
            langd = o.lopande_langd if o.lopande_langd is not None else max(0.0, b.langd - o.lage)   # öppen skada: till avbrottet
            if snitt_b is not None and o.lage + langd > snitt_b:
                langd = max(0.0, snitt_b - o.lage)
            o2.lopande_langd = round(langd, 2)
            o2.lage = round(max(0.0, L - o.lage - langd), 2)
        elif o.lopande.startswith("B"):
            o2.lage = round(max(0.0, L - o.lage + (lop_langd.get(o.lopande[1:]) or 0.0)), 2)
        else:
            o2.lage = round(max(0.0, L - o.lage), 2)
        o2.klocka_fran, o2.klocka_till = _spegla_klocka(o.klocka_till or o.klocka_fran), (
            _spegla_klocka(o.klocka_fran) if o.klocka_till else "")
        # Utan kartlängd: båda kamerorna stannade vid samma hinder – samma observation inom 1 m
        # från mötet från den äldre filmen är en dubblett
        if not L_karta and abs(o2.lage - a.langd) <= 1.0:
            dubb = [q for q in a_obs if abs(q.lage - a.langd) <= 1.0 and q.kod == o.kod and q.grad == o.grad
                    and q.infokod == o.infokod and q.attribut == o.attribut and q.lopande == o.lopande]
            if dubb:
                bort = dubb[0] if nyare_b else o2
                if bort is not o2:
                    obs.remove(bort)
                    a_obs.remove(bort)
                if "samma observation vid mötet räknas en gång" not in noter:
                    noter.append("samma observation vid mötet räknas en gång")
                if bort is o2:
                    continue
        marka(o2, b)
        obs.append(o2)
    obs.sort(key=lambda o: (o.lage, o.tid))
    # sträckans längd = största läge; se till att slutbrunnen ligger vid L
    if not obs or max(o.lage for o in obs) < L - 0.01:
        slut = copy.copy((b.observationer or a.observationer)[0])
        slut.nr, slut.lage, slut.kod, slut.grad, slut.infokod, slut.attribut = nr, round(L, 2), "", None, "", ""
        slut.klocka_fran = slut.klocka_till = slut.bild = slut.bild_b = slut.kommentar = ""
        slut.lopande, slut.lopande_langd, slut.bild_sokvag, slut.bild_b_sokvag = "", None, None, None
        obs.append(slut)
    ny = dataclasses.replace(a, nr=nr, observationer=obs, profil=[], profil_start_z=None, profil_slut_z=None,
                             utgangsbrunn=a.startbrunn, riktning="Medströms",
                             datum=nyare.datum, klockslag=nyare.klockslag, flerinspekterad=True, rapport_fil=None,
                             sammanslagen_av=[a.nr, b.nr], _cache={})
    # Film a:s videofil i Videofil, film b:s i Videofil 2 (egen länk i Excel, VIDEO2 i kartan)
    if not a.videofil and b.videofil:
        ny.videofil, ny.video_sokvag = b.videofil, b.video_sokvag
    elif b.videofil and b.videofil != a.videofil:
        ny.videofil_b, ny.video_sokvag_b = b.videofil, b.video_sokvag
    ny.manuell_bedomning = a.manuell_bedomning or b.manuell_bedomning
    ny.kommentar = a.kommentar or b.kommentar
    ny.lagning_m = a.lagning_m if a.lagning_m is not None else b.lagning_m
    if not ny.gis:
        ny.gis = b.gis
    if ny.langd_karta is None:
        ny.langd_karta, ny.gis_vg_start, ny.gis_vg_slut, ny.mark = b.langd_karta, b.gis_vg_start, b.gis_vg_slut, b.mark
    if L_karta:
        if overlapp > 0.5:
            noter.append(dk(f"överlapp {overlapp:.1f} m, nyare filmens observationer gäller där"))
        elif overlapp < -0.5:
            noter.append(dk(f"lucka {-overlapp:.1f} m ofilmad mitt på"))
    elif not any("orimlig" in n for n in noter):
        noter.insert(0, "kartlängd okänd – filmerna antas mötas utan överlapp")
    else:
        noter.insert(1, "filmerna antas mötas utan överlapp")
    ny.filmstatus = f"sammanslagen av nr {a.nr} + {b.nr}" + (" (" + "; ".join(noter) + ")" if noter else "")
    ny.sammanslagning = {"a_nr": a.nr, "b_nr": b.nr, "a_fran": a.fran_brunn, "b_fran": b.fran_brunn,
                         "a_langd": a.langd, "b_langd": b.langd, "L": L, "overlapp": overlapp if L_karta else 0.0}
    return ny


def _ersatt(s: Stracka, gall: Stracka, orsak: str, text: str = "") -> None:
    """Markerar s som ersatt av gall; varnar i statusen när den ersatta filmen hade sämre klass."""
    s.filmstatus = (text or f"ersatt av nr {gall.nr}") + (f" ({orsak})" if orsak else "")
    if s.langd >= 1 and s.klass < gall.klass:
        s.filmstatus += f" – OBS: den ersatta filmen var klass {s.klass}, den gällande är {gall.klass}"


def _utveckling(gammal: Stracka, ny: Stracka) -> str:
    """Hur sträckan förändrats mellan två inspektioner: klassbyte avgör, inom samma klass räknas
    konstruktionsindex ändrat mer än TIDIGARE_INDEX_ANDEL (och minst 5 p/100 m)."""
    if ny.klass < gammal.klass:
        return "förvärrad"
    if ny.klass > gammal.klass:
        return "förbättrad"
    g, n = gammal.index("K"), ny.index("K")
    if n - g >= max(5.0, g * TIDIGARE_INDEX_ANDEL):
        return "förvärrad"
    if g - n >= max(5.0, g * TIDIGARE_INDEX_ANDEL):
        return "förbättrad"
    return "oförändrad"


def tidigare_text(s: Stracka) -> str:
    """Texten i kolumnen Tidigare inspektion: '2021-05-27 (DUF 701.TV3 nr 12): C 21 → B 45 p/100 m – förvärrad'."""
    t = s.tidigare
    if not t:
        return ""
    txt = (f"{t['datum'] or 'okänt datum'} ({os.path.basename(t['fil'])} nr {t['nr']}): "
           f"{t['klass']} {t['index']:.0f} → {s.klass} {s.index('K'):.0f} p/100 m – {t['utveckling']}")
    if t["antal"] > 1:
        txt += f" ({t['antal']} äldre filmer)"
    return txt


def _over_filer(strackor: list[Stracka]) -> None:
    """Samma brunnspar i olika TV3-filer: den nyaste filmen (senare datum) gäller, äldre filmer i andra
    filer markeras 'ersatt av ny inspektion …' och den gällande får `tidigare` med den närmast
    föregående inspektionens klass och index. Filmer utan tolkbart datum rörs inte (varning)."""
    grupper: dict[frozenset, list[Stracka]] = {}
    for s in aktiva(strackor):
        if s.langd >= 1 and s.startbrunn and s.slutbrunn and _normlittera(s.startbrunn) != _normlittera(s.slutbrunn):
            grupper.setdefault(frozenset((_normlittera(s.startbrunn), _normlittera(s.slutbrunn))), []).append(s)
    utan_datum = []
    for par, grupp in grupper.items():
        if len({s.fil for s in grupp}) < 2:
            continue
        if any(_tidsnyckel(s)[0] == 0 for s in grupp):
            utan_datum.append(grupp)
            continue
        nyast = max(grupp, key=lambda s: (_tidsnyckel(s), s.fil, s.nr))
        dag = _tidsnyckel(nyast)[1]
        # gällande = filmerna i den nyaste filen (kvar efter filens egen hantering, normalt en); alla
        # filmer i andra filer ersätts – även samma dag (då troligen samma film i två exporter)
        gall = [s for s in grupp if s.fil == nyast.fil]
        aldre = sorted((s for s in grupp if s.fil != nyast.fil), key=lambda s: (_tidsnyckel(s), s.nr), reverse=True)
        tidigare = [s for s in aldre if _tidsnyckel(s)[1] < dag]      # riktiga tidigare inspektioner, inte dubbletter
        for s in aldre:
            orsak = []
            if _tidsnyckel(s)[1] == dag:
                orsak.append("samma datum – dubblett?")
            if nyast.ofullstandig and not s.ofullstandig:
                orsak.append(f"OBS: den nya filmen är ofullständig, {nyast.langd:.1f} m mot {s.langd:.1f} m")
            _ersatt(s, nyast, "; ".join(orsak),
                    f"ersatt av ny inspektion nr {nyast.nr} i {os.path.basename(nyast.fil)} {nyast.datum}")
        if not tidigare:
            continue
        senast = tidigare[0]
        for g in gall:
            g.tidigare = {"fil": senast.fil, "nr": senast.nr, "datum": senast.datum, "klass": senast.klass,
                          "index": senast.index("K"), "langd": senast.langd, "antal_skador": len(senast.skador()),
                          "utveckling": _utveckling(senast, g), "antal": len(tidigare)}
    for grupp in utan_datum:
        s = grupp[0]
        print(f"  OBS     {s.id} finns i {len({x.fil for x in grupp})} filer men datum saknas/kan inte tolkas "
              f"– alla filmerna räknas")


def hantera_omfilmningar(strackor: list[Stracka]) -> list[Stracka]:
    """Flera filmer av samma brunnspar i samma TV3-fil: hel film ersätter ofullständiga och äldre
    hela filmer (filmstatus "ersatt av nr X (…)"), och två ofullständiga filmer från var sin brunn
    slås ihop till en ny sträcka som läggs sist i listan. Därefter (OMFILMNING_OVER_FILER) samma
    brunnspar i olika filer: den nyaste inspektionen gäller, äldre markeras ersatta och den gällande
    får `tidigare`. Returnerar listan inklusive sammanslagna."""
    if not (OMFILMNING or SAMMANSLAGNING):
        if OMFILMNING_OVER_FILER:
            _over_filer(strackor)
        return strackor
    grupper: dict[tuple[str, frozenset], list[Stracka]] = {}
    for s in strackor:
        if s.startbrunn and s.slutbrunn and _normlittera(s.startbrunn) != _normlittera(s.slutbrunn):
            grupper.setdefault((s.fil, frozenset((_normlittera(s.startbrunn), _normlittera(s.slutbrunn)))), []).append(s)
    nya = []
    max_nr = {}
    for s in strackor:
        max_nr[s.fil] = max(max_nr.get(s.fil, 0), s.nr)
    for (fil, par), grupp in grupper.items():
        if len(grupp) < 2:
            continue
        nyast_forst = sorted(grupp, key=lambda s: (_tidsnyckel(s), s.nr), reverse=True)
        hela = [s for s in nyast_forst if s.langd >= 1 and not s.ofullstandig]
        delar = [s for s in nyast_forst if s.langd >= 1 and s.ofullstandig]
        tomma = [s for s in nyast_forst if s.langd < 1]
        if OMFILMNING and hela:
            # gällande = nyaste hela filmen, om ingen äldre är klart längre (nyare < 85 % av äldre)
            gall = hela[0]
            for s in hela[1:]:
                if s.langd > gall.langd / OMFILMNING_MIN_ANDEL:
                    gall = s
            for s in hela:
                if s is not gall:
                    _ersatt(s, gall, "längre film" if _tidsnyckel(s) > _tidsnyckel(gall) else f"omfilmning {gall.datum}")
            for s in delar:
                _ersatt(s, gall, "hel film")
            for s in tomma:
                _ersatt(s, gall, "")
            continue
        if delar:
            # samma utgångsbrunn: den längsta filmen gäller (nyast vid lika längd)
            per_sida: dict[str, list[Stracka]] = {}
            for s in delar:
                per_sida.setdefault(_normlittera(s.fran_brunn), []).append(s)
            kvar = []
            for sida, lista in per_sida.items():
                basta = max(lista, key=lambda s: (round(s.langd, 1), _tidsnyckel(s)))
                if OMFILMNING:
                    for s in lista:
                        if s is not basta:
                            _ersatt(s, basta, "längre film från samma brunn")
                kvar.append(basta)
            ny = None
            if SAMMANSLAGNING and len(kvar) == 2:
                a = next((s for s in kvar if _normlittera(s.fran_brunn) == _normlittera(s.startbrunn)), None)
                b = next((s for s in kvar if s is not a), None)
                if a is not None and b is not None and _normlittera(b.fran_brunn) == _normlittera(b.slutbrunn):
                    max_nr[fil] += 1
                    ny = _sammanslagen(a, b, max_nr[fil])
                    for s, andra in ((a, b), (b, a)):
                        s.filmstatus = (f"ingår i sammanslagen nr {ny.nr} tillsammans med nr {andra.nr} "
                                        f"(filmad från {andra.fran_brunn}, {andra.langd:.1f} m) – se protokollet för sträcka {ny.nr}")
                    nya.append(ny)
            if OMFILMNING:
                for s in tomma:
                    _ersatt(s, ny if ny is not None else kvar[0], "")
    alla = strackor + nya
    if OMFILMNING_OVER_FILER:
        _over_filer(alla)
    return alla


def inspektionsgrad(strackor: list[Stracka], filer: list[dict]) -> dict:
    """Hur stor del av ledningsnätet i GIS som är filmat, per driftområde ('omrade' i gisdata.json)
    och ledningstyp. Filmad längd räknas med GIS-ledningens längd så andelen blir konsekvent.
    Returnerar {"rader": [{omrade, ledningstyp, ledningar, langd_m, filmade, filmad_m, andel}],
    "ej_inspekterat": [GIS-ledningar utan film], "utan_gis": (sträckor, m) filmade utan GIS-ledning}."""
    # GIS-ledning -> sämsta gällande klass bland sträckorna på den (syskon filmade från båda håll)
    klass_pa: dict[int, str] = {}
    for s in strackor:
        if s.langd >= 1 and s.gis and s.gis.get("ledning"):
            k = s.gallande_klass
            led = s.gis["ledning"]
            if led.get("_syntetisk") == "fagelvag":
                continue                                   # ingen GIS-ledning att räkna som filmad
            for d in led.get("_delar") or [led]:
                i = id(d)
                if i not in klass_pa or k < klass_pa[i]:
                    klass_pa[i] = k
    grupp: dict[tuple[str, str], dict] = {}

    def lagg(nyckel, led, klass):
        g = grupp.setdefault(nyckel, {"ledningar": 0, "langd_m": 0.0, "filmade": 0, "filmad_m": 0.0,
                                      "A": 0, "B": 0, "AB_m": 0.0})
        L = led.get("langd_m") or 0.0
        g["ledningar"] += 1
        g["langd_m"] += L
        if klass:
            g["filmade"] += 1
            g["filmad_m"] += L
            if klass in ("A", "B"):
                g[klass] += 1
                g["AB_m"] += L

    ej = []
    for gf in filer:
        for led in gf["ledningar"]:
            if not led.get("fran") or not led.get("till") or led.get("_dubblett"):
                continue                                   # fri ände eller samma ledning i flera filer
            omr = led.get("omrade") or "(utan område)"
            typ = (led.get("ledningstyp") or "").strip() or "(okänd typ)"
            f = klass_pa.get(id(led))
            lagg((omr, typ), led, f)
            lagg((omr, "alla"), led, f)
            lagg(("alla", typ), led, f)
            lagg(("alla", "alla"), led, f)
            if not f:
                ej.append({"omrade": led.get("omrade") or "", "fran": led["fran"], "till": led["till"],
                           "langd_m": led.get("langd_m"), "dimension": led.get("dimension"),
                           "material": led.get("material") or "", "ledningstyp": led.get("ledningstyp") or "",
                           "anlaggningsar": led.get("anlaggningsar")})
    rader = []
    for (omr, typ), g in sorted(grupp.items(), key=lambda kv: (kv[0][0] == "alla", kv[0][0], kv[0][1] == "alla", kv[0][1])):
        rader.append({"omrade": omr, "ledningstyp": typ, **g,
                      "andel": g["filmad_m"] / g["langd_m"] if g["langd_m"] else 0.0,
                      "andel_AB": g["AB_m"] / g["filmad_m"] if g["filmad_m"] else None})
    ej.sort(key=lambda e: (e["omrade"], e["anlaggningsar"] if isinstance(e["anlaggningsar"], int) else 9999,
                           str(e["material"]), -(e["langd_m"] or 0)))
    utan = [s for s in strackor if s.langd >= 1 and not (s.gis and s.gis.get("ledning")
                                                          and s.gis["ledning"].get("_syntetisk") != "fagelvag")]
    return {"rader": rader, "ej_inspekterat": ej, "utan_gis": (len(utan), sum(s.langd for s in utan))}


def skriv_litteraforslag(forslag: list[dict], path: str) -> int:
    """Skriver okända brunnar med förslag som en brunnslittera-CSV (fel;ratt;fil;nr;motbrunn;kommentar)
    som kan rättas och användas med littera: i listfilen. Returnerar antal rader."""
    rader = ["fel;ratt;fil;nr;motbrunn;kommentar"]
    for o in forslag:
        ratt = o["forslag"][0][0] if o["forslag"] and o.get("sakert") else ""
        komm = (f"förslag ({o['metod']}): " + " | ".join(f"{lit} – {txt}" for lit, txt in o["forslag"])
                + ("" if o.get("sakert") else " | osäkert – välj själv i kolumnen ratt")
                if o["forslag"] else "inget förslag: ingen ledning med passande längd och inget liknande littera i GIS")
        komm = komm.replace(";", "/").replace(",", ".")    # kommentaren får inte innehålla CSV-avgränsare
        rader.append(f"{o['littera']};{ratt};{os.path.basename(o['fil'])};{o['nr']};{o['motbrunn']};{komm}")
    with open(path, "w", encoding="utf-8-sig", newline="") as fh:
        fh.write("\r\n".join(rader) + "\r\n")
    return len(rader) - 1


def las_uppstroms(path: str) -> list[dict]:
    """Läser CSV:n från ArcMap-verktyget Uppströms (batchläge): fil;nr;startbrunn;slutbrunn;
    serviser_uppstroms;langd_uppstroms_m;... med decimalkomma."""
    rader = [r for r in las_text(path).splitlines() if r.strip()]
    if not rader:
        return []
    rubrik = [k.strip().lower() for k in rader[0].split(";")]
    poster = []
    for rad in rader[1:]:
        delar = rad.split(";")
        post = dict(zip(rubrik, (d.strip() for d in delar)))
        poster.append(post)
    return poster


def koppla_uppstroms(strackor: list[Stracka], poster: list[dict]) -> int:
    """Kopplar serviser/längd uppströms till sträckorna: TV3-fil + sträcknummer i första hand,
    annars brunnspar. Returnerar antal kopplade sträckor."""
    def tal(v):
        try:
            return float(str(v).replace(",", "."))
        except (TypeError, ValueError):
            return None
    pa_nr: dict[tuple[str, int], dict] = {}
    pa_par: dict[frozenset, dict] = {}
    for post in poster:
        fil = os.path.basename(post.get("fil") or "").lower()
        par = frozenset((_normlittera(post.get("startbrunn")), _normlittera(post.get("slutbrunn"))))
        nr = tal(post.get("nr"))
        if nr is not None and fil:
            pa_nr.setdefault((fil, int(nr)), post)
        pa_par.setdefault(par, post)
    n = 0
    for s in strackor:
        fil = os.path.basename(s.fil).lower()
        par = frozenset((_normlittera(s.startbrunn), _normlittera(s.slutbrunn)))
        post = pa_nr.get((fil, s.nr)) or pa_par.get(par)
        if not post:
            continue
        serv = tal(post.get("serviser_uppstroms"))
        s.serviser_uppstroms = int(serv) if serv is not None else None
        s.langd_uppstroms = tal(post.get("langd_uppstroms_m"))
        s.serviser_kalla = post.get("serviser_kalla", "")
        n += 1
    return n


def fil_url(sokvag: str) -> str:
    """Absolut sökväg -> file:///-URL som Excel kan öppna."""
    from urllib.parse import quote
    p = os.path.abspath(sokvag).replace("\\", "/")
    if not p.startswith("/"):
        p = "/" + p                     # C:/... -> /C:/...
    return "file://" + quote(p, safe="/:")


# ----------------------------------------------------------------------------
# Åtgärdspaket och kostnad (PLAN_verktyg.md steg 3)
# ----------------------------------------------------------------------------

def las_kostnader(path: str) -> list[dict]:
    """Läser kostnadsfilen: post;dimension_fran;dimension_till;enhet;kr;kommentar (decimalkomma,
    # = kommentar). Returnerar [{"post", "fran", "till", "enhet", "kr", "kommentar"}]."""
    poster = []
    for rad in las_text(path).splitlines():
        rad = rad.split("#", 1)[0].strip()
        if not rad or rad.lower().startswith("post"):
            continue
        d = [x.strip() for x in re.split(r"[;\t]", rad)]
        d += [""] * (6 - len(d))

        def tal(v):
            v = re.sub(r"[^\d,.\-]", "", v)            # "9 000 kr" -> "9000"
            if "," in v:
                v = v.replace(".", "").replace(",", ".")  # "2.500,50" -> 2500.50
            try:
                return float(v) if v else None
            except ValueError:
                return None
        kr = tal(d[4])
        if not d[0]:
            continue
        if kr is None:
            print(f"  VARNING kostnadsfil: kan inte tolka beloppet '{d[4]}' för posten {d[0]} – raden hoppas över")
            continue
        poster.append({"post": d[0].lower(), "fran": tal(d[1]), "till": tal(d[2]),
                       "enhet": d[3], "kr": kr, "kommentar": d[5]})
    return poster


def pris(kostnader: list[dict], post: str, dim_mm: int | None = None) -> float | None:
    """kr för en post; för dimensionsberoende poster den rad vars intervall täcker dim_mm."""
    kandidater = [k for k in kostnader if k["post"] == post]
    if not kandidater:
        return None
    med_intervall = [k for k in kandidater if k["fran"] is not None or k["till"] is not None]
    if not med_intervall:
        return kandidater[0]["kr"]
    if dim_mm is None:
        return next((k["kr"] for k in kandidater if k not in med_intervall), None)
    for k in med_intervall:
        if (k["fran"] is None or dim_mm >= k["fran"]) and (k["till"] is None or dim_mm <= k["till"]):
            return k["kr"]
    # Rad utan intervall gäller när inget intervall täcker dimensionen
    return next((k["kr"] for k in kandidater if k not in med_intervall), None)


def las_manuella(path: str) -> tuple[list[dict], list[str]]:
    """Läser manuella bedömningar ur en tidigare prioritering.xlsx: fliken Prioritering
    (Fil, Nr, Startbrunn, Slutbrunn, Manuell bedömning, Kommentar, Lagning) och fliken Etapper
    (kolumnen 'Schakta fram (manuellt)'). Returnerar (poster, framschakta_brunnar)."""
    from openpyxl import load_workbook
    wb = load_workbook(path, read_only=True, data_only=True)
    poster: list[dict] = []
    fram: list[str] = []
    if "Prioritering" in wb.sheetnames:
        ws = wb["Prioritering"]
        rader = ws.iter_rows(values_only=True)
        rubrik = [str(c or "").strip() for c in next(rader, [])]

        def kol(namn):
            lag = [r.lower() for r in rubrik]
            if namn.lower() in lag:                       # exakt träff först
                return lag.index(namn.lower())
            for i, r in enumerate(lag):
                if r.startswith(namn.lower()):
                    return i
            return None
        ix = {n: kol(n) for n in ("Fil", "Nr", "Startbrunn", "Slutbrunn", "Manuell bedömning",
                                  "Kommentar", "Lagning")}
        for rad in rader:
            if not rad or ix["Startbrunn"] is None:
                continue

            def v(n):
                i = ix[n]
                return rad[i] if i is not None and i < len(rad) else None
            if v("Startbrunn") is None or (ix["Nr"] is not None and not isinstance(v("Nr"), (int, float))):
                continue                      # enhetsrad eller tom rad
            man, kom, lag = v("Manuell bedömning"), v("Kommentar"), v("Lagning")
            if man is None and kom is None and lag is None:
                continue
            lagning = None
            if lag not in (None, ""):
                m_lag = re.search(r"\d+(?:[.,]\d+)?", str(lag))
                if m_lag:
                    lagning = float(m_lag.group(0).replace(",", "."))
                    if str(lag).strip() != m_lag.group(0):
                        print(f"  Lagning '{lag}' ({v('Startbrunn')}→{v('Slutbrunn')}) tolkas som {lagning:g} m")
                else:
                    print(f"  VARNING Lagning '{lag}' ({v('Startbrunn')}→{v('Slutbrunn')}) kunde inte tolkas – ignoreras")
            poster.append({"fil": str(v("Fil") or ""), "nr": int(v("Nr")) if isinstance(v("Nr"), (int, float)) else None,
                           "startbrunn": str(v("Startbrunn") or ""), "slutbrunn": str(v("Slutbrunn") or ""),
                           "manuell": str(man or "").strip(), "kommentar": str(kom or "").strip(),
                           "lagning_m": lagning})
    if "Etapper" in wb.sheetnames:
        ws = wb["Etapper"]
        rader = ws.iter_rows(values_only=True)
        rubrik = [str(c or "").strip().lower() for c in next(rader, [])]
        ci = next((i for i, r in enumerate(rubrik) if r.startswith("schakta fram")), None)
        ce = next((i for i, r in enumerate(rubrik) if r == "etapp"), None)
        if ci is not None:
            for rad in rader:
                if not rad or ci >= len(rad) or not rad[ci]:
                    continue
                if ce is not None and ce < len(rad) and not isinstance(rad[ce], (int, float)):
                    continue                      # enhetsraden ("manuellt")
                fram += [b.strip() for b in re.split(r"[,;\s]+", str(rad[ci])) if b.strip()]
    wb.close()
    return poster, fram


def koppla_manuella(strackor: list[Stracka], poster: list[dict]) -> int:
    """Kopplar manuella bedömningar: TV3-fil + sträcknummer i första hand, annars brunnspar."""
    pa_nr: dict[tuple[str, int], dict] = {}
    pa_par: dict[frozenset, dict] = {}
    for post in poster:
        fil = os.path.basename(post["fil"]).lower()
        if post["nr"] is not None and fil:
            pa_nr.setdefault((fil, post["nr"]), post)
        pa_par.setdefault(frozenset((_normlittera(post["startbrunn"]), _normlittera(post["slutbrunn"]))), post)
    n = 0
    for s in strackor:
        post = pa_nr.get((os.path.basename(s.fil).lower(), s.nr)) or \
            pa_par.get(frozenset((_normlittera(s.startbrunn), _normlittera(s.slutbrunn))))
        if not post:
            continue
        s.manuell_bedomning, s.kommentar, s.lagning_m = post["manuell"], post["kommentar"], post["lagning_m"]
        n += 1
    return n


def planera_atgarder(strackor: list[Stracka], kostnader: list[dict],
                     framschakta_manuellt: list[str] | None = None) -> list[dict]:
    """Delar in åtgärdssträckorna i etapper, väljer brunnar att schakta fram och räknar kostnad.
    Sätter s.etapp, s._metod (överbryggning), s.etapp_flagga, s.kostnad, s.kostnadsflagga.
    Returnerar etapperna (dict per etapp, numrerade efter ETAPP_ORDNING)."""
    fram_manuellt = {_normlittera(b) for b in (framschakta_manuellt or []) + list(FRAMSCHAKTA_BRUNNAR)}
    klassordn = {"A": 0, "B": 1, "C": 2, "D": 3, "E": 4}

    # En representant per brunnspar (värsta gällande klass); syskon följer med utan egen kostnad
    def par_av(s):
        p = frozenset((_normlittera(s.startbrunn), _normlittera(s.slutbrunn)))
        # Samma (eller tom) brunn i båda ändar: sträckan får ett eget "par" så att den ändå
        # kan bli en egen etapp; den kopplas inte till grannar
        return p if len(p) == 2 and all(p) else frozenset((f"#{s.fil}", f"#{s.nr}"))

    for s in strackor:
        s.etapp, s._metod, s.etapp_flagga, s.kostnad, s.kostnadsflagga = None, None, "", None, ""
    # Representant per brunnspar: en sträcka med manuell bedömning går före, sedan värsta
    # gällande klass och högst index. Syskon (flerinspekterat par) följer representanten.
    rep_: dict[frozenset, Stracka] = {}

    def rang(s):
        return (not bool((s.manuell_bedomning or "").strip()), klassordn[s.gallande_klass], -s.index("K"))
    for s in strackor:
        r = rep_.get(par_av(s))
        if r is None or rang(s) < rang(r):
            rep_[par_av(s)] = s

    atgard = {p: r for p, r in rep_.items() if r.metod in ("strumpa", "schakt")}
    brunn_metod: dict[tuple[str, str], list[Stracka]] = defaultdict(list)   # (brunn, metod) -> sträckor
    for r in atgard.values():
        for b in par_av(r):
            brunn_metod[(b, r.metod)].append(r)

    # Överbryggning: en kedja av högst OVERBRYGGA_MAX_STRACKOR sträckor utan egen åtgärd som förbinder
    # två åtgärdssträckor med samma metod tas med – B-sträckor (OVERBRYGGA_KLASSER) oavsett längd,
    # C/D bara om kortare än ETAPP_OVERBRYGGA_M (F5). Manuellt 'ingen' och relinade överbryggas inte.
    def overbryggbar(r):
        if r.metod != "ingen" or r.relinad or r.manuell_metod == "ingen" or r.langd < 1:
            return False
        if r.gallande_klass in OVERBRYGGA_KLASSER:
            return True
        return r.gallande_klass in ("C", "D") and r.langd < ETAPP_OVERBRYGGA_M

    for metod in ("strumpa", "schakt"):
        kand = {p: r for p, r in rep_.items() if p not in atgard and len(p) == 2 and overbryggbar(r)}
        kand_vid: dict[str, list[frozenset]] = defaultdict(list)
        for p in kand:
            for b in p:
                kand_vid[b].append(p)
        atg_brunnar = {b for (b, m) in brunn_metod if m == metod and brunn_metod[(b, m)]}
        medtagna: set[frozenset] = set()
        for start in sorted(atg_brunnar):
            # djupet-först över kandidater: vägen får inte återbesöka en sträcka eller brunn
            stack = [(start, [], {start})]
            while stack:
                brunn, vag, besokta = stack.pop()
                if len(vag) >= OVERBRYGGA_MAX_STRACKOR:
                    continue
                for p in kand_vid.get(brunn, []):
                    if p in vag:
                        continue
                    nasta = next(b for b in p if b != brunn)
                    ny_vag = vag + [p]
                    if nasta in atg_brunnar and nasta != start:
                        medtagna.update(ny_vag)
                    elif nasta not in besokta and nasta not in atg_brunnar:
                        stack.append((nasta, ny_vag, besokta | {nasta}))
        for p in medtagna:
            r = kand[p]
            r._metod = metod
            r.etapp_flagga = ("medtagen för sammanhang (klass %s mellan två åtgärdssträckor)" % r.gallande_klass
                              if r.gallande_klass in OVERBRYGGA_KLASSER else "medtagen för sammanhang")
            atgard[p] = r
            for bb in p:
                brunn_metod[(bb, metod)].append(r)

    # Sammanhängande sträckor med samma metod = etapp (bredd-först över gemensamma brunnar)
    etapper: list[dict] = []
    sedda: set[frozenset] = set()
    for p in sorted(atgard, key=lambda q: (klassordn[atgard[q].gallande_klass], -atgard[q].index("K"))):
        if p in sedda:
            continue
        metod = atgard[p].metod
        grupp, ko = [], [p]
        sedda.add(p)
        while ko:
            q = ko.pop()
            grupp.append(atgard[q])
            for b in q:
                for granne in brunn_metod.get((b, metod), []):
                    gp = par_av(granne)
                    if gp not in sedda:
                        sedda.add(gp)
                        ko.append(gp)
        etapper.append({"metod": metod, "strackor": grupp})

    # Ordning: högsta konstruktionsindex först, eller flest serviser uppströms (konsekvens)
    def nyckel(e):
        idx = max(s.index("K") for s in e["strackor"])
        serv = max((s.serviser_uppstroms or 0) for s in e["strackor"])
        return (-serv, -idx) if ETAPP_ORDNING == "konsekvens" else (-idx, -serv)
    etapper.sort(key=nyckel)

    p_hatt, p_lagn = pris(kostnader, "hatt"), pris(kostnader, "lagning")
    p_fram, p_brunn, p_etab = pris(kostnader, "framschaktning"), pris(kostnader, "ny_brunn"), pris(kostnader, "etablering")
    for nr, e in enumerate(etapper, 1):
        e["nr"] = nr
        strumpa = [s for s in e["strackor"] if s.metod == "strumpa"]
        # F9: brunnar att schakta fram – så få som möjligt, varje strumpsträcka med bara
        # tillsyns-/rensbrunnar ska ha en framschaktad brunn i någon ände (okänd typ = som NB)
        behov = [s for s in strumpa if s.brunnstyp(s.startbrunn) in ("TB", "RB")
                 and s.brunnstyp(s.slutbrunn) in ("TB", "RB")]
        kandidater: dict[str, list[Stracka]] = defaultdict(list)
        for s in behov:
            for b in par_av(s):
                kandidater[b].append(s)
        valda = [b for b in sorted(kandidater) if b in fram_manuellt]
        tackta = {id(s) for b in valda for s in kandidater[b]}
        while len(tackta) < len(behov):
            b = min(kandidater, key=lambda k: (-sum(1 for s in kandidater[k] if id(s) not in tackta),
                                                -len(kandidater[k]), k))
            if not any(id(s) not in tackta for s in kandidater[b]):
                break
            valda.append(b)
            tackta |= {id(s) for s in kandidater[b]}
        littera = {_normlittera(x): x for s in behov for x in (s.startbrunn, s.slutbrunn)}
        e["framschaktade"] = [littera.get(b, b) for b in valda]
        e["framschakt_manuell"] = [littera.get(b, b) for b in valda if b in fram_manuellt]

        # Kostnad per sträcka (strumpa); schakt kalkyleras inte (F8)
        summa = {"strumpa": 0.0, "hattar": 0.0, "lagning": 0.0}
        for s in e["strackor"]:
            s.etapp = nr
            if s.metod != "strumpa":
                continue
            dim = s.dimension_mm
            p_str = pris(kostnader, "strumpa", dim)
            k = {"strumpa": p_str * s.langd if p_str is not None else None,
                 "hattar": p_hatt * s.antal_anslutningar if p_hatt is not None else None,
                 "lagning": p_lagn * s.lagning_m if (p_lagn is not None and s.lagning_m) else 0.0}
            k["summa"] = sum(v for v in k.values() if v is not None)
            s.kostnad = k
            if p_str is None:
                s.kostnadsflagga = (f"dimension {s.dimension or '?'} saknar pris"
                                    if any(k["post"] == "strumpa" for k in kostnader) else "kostnadsfil saknas")
            for n_ in summa:
                summa[n_] += k[n_] or 0.0
        n_fram = len(e["framschaktade"])
        e["kostnad"] = dict(summa)
        e["kostnad"]["brunnar"] = n_fram * ((p_fram or 0.0) + (p_brunn or 0.0))
        e["kostnad"]["etablering"] = (p_etab or 0.0) if strumpa else 0.0
        e["kostnad"]["summa"] = sum(e["kostnad"].values())
        e["schakt_m"] = sum(s.langd for s in e["strackor"] if s.metod == "schakt")
        e["langd_m"] = sum(s.langd for s in e["strackor"])
    # Syskon (samma brunnspar, flera inspektioner) får representantens etapp och metod
    for s in strackor:
        r = rep_.get(par_av(s))
        if r is not None and r is not s:
            if r.etapp is not None:
                s.etapp, s._metod, s.etapp_flagga = r.etapp, r.metod, r.etapp_flagga
                s.kostnadsflagga = f"kostnad räknad på nr {r.nr}"
            if (s.manuell_bedomning or "").strip():
                s.etapp_flagga = (s.etapp_flagga + "; " if s.etapp_flagga else "") + \
                    f"manuell bedömning på nr {r.nr} gäller för brunnsparet"
        if len(frozenset((_normlittera(s.startbrunn), _normlittera(s.slutbrunn)))) < 2 and s.metod != "ingen":
            s.etapp_flagga = (s.etapp_flagga + "; " if s.etapp_flagga else "") + "samma brunn i båda ändar – egen etapp"
    return etapper


# ----------------------------------------------------------------------------
# Excel
# ----------------------------------------------------------------------------

def skriv_excel(strackor: list[Stracka], path: str, diagram: dict[str, str], topp: int,
                etapper: list[dict] | None = None, gisstat: dict | None = None):
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

    def tabell(ws, kolumner, rader, bredder=None, klasskol=None, lankar=None, dolda=None) -> int:
        """Skriver rubrikrad, enhetsrad (om någon rubrik har enhet inom parentes, t.ex.
        "Längd (m)": rubriken blir "Längd" och enheten hamnar på raden under) och data.
        lankar: {kolumnindex: [url eller None per rad]} – gör cellerna klickbara.
        dolda: kolumnrubriker (utan enhet) som döljs, grupperade så de kan fällas ut med plustecknet.
        Returnerar första dataradens radnummer."""
        namn, enheter = [], []
        for k in kolumner:
            m = re.match(r"^(.*?)\s*\((.+)\)$", k)
            namn.append(m.group(1) if m else k)
            enheter.append(m.group(2) if m else "")
        har_enhet = any(enheter)
        ws.append(namn)
        for c in ws[1]:
            c.font, c.fill, c.border = rubrik, rubrikfyll, kant
            c.alignment = Alignment(vertical="center", wrap_text=True)
        if har_enhet:
            ws.append(enheter)
            for c in ws[2]:
                c.font = Font(italic=True, color="595959", size=9)
                c.fill = PatternFill("solid", fgColor="EDEDED")
                c.border = kant
                c.alignment = Alignment(horizontal="center", vertical="center")
        start = 3 if har_enhet else 2
        for rad in rader:
            ws.append(rad)
        for ri, row in enumerate(ws.iter_rows(min_row=start, max_row=ws.max_row)):
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
        ws.freeze_panes = f"A{start}"
        ws.auto_filter.ref = f"A{start - 1}:{get_column_letter(len(kolumner))}{max(ws.max_row, start)}"
        for i, kol in enumerate(kolumner, 1):
            b = (bredder or {}).get(kol, (bredder or {}).get(namn[i - 1]))
            if b is None:
                langsta = max((len(str(r[i - 1])) for r in rader[:200] if i - 1 < len(r) and r[i - 1] is not None),
                              default=0)
                b = min(45, max(10, langsta + 2, max(len(namn[i - 1]), len(enheter[i - 1])) + 2))
            ws.column_dimensions[get_column_letter(i)].width = b
        # Dolda kolumner: sammanhängande block grupperas med outline så att Excel visar ett
        # plustecken ovanför rubrikraden där de kan fällas ut
        dold = set(dolda or [])
        index = sorted(i for i, kol in enumerate(kolumner, 1) if kol in dold or namn[i - 1] in dold)
        block: list[list[int]] = []
        for i in index:
            if block and i == block[-1][-1] + 1:
                block[-1].append(i)
            else:
                block.append([i])
        for b in block:
            for i in b:                      # per kolumn: bredden är redan satt på samma objekt
                cd = ws.column_dimensions[get_column_letter(i)]
                cd.hidden = True
                cd.outlineLevel = 1
        if block:
            ws.sheet_properties.outlinePr.summaryRight = False
        return start

    # ---- Sammanfattning ----
    alla = strackor
    strackor = aktiva(alla)        # ersatta/delfilmer räknas inte i statistiken men listas i Prioritering
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
        ["Omfilmningar", (f"{sum(1 for s in alla if s.filmstatus.startswith('ersatt'))} filmer ersatta av nyare/hel film, "
                          f"{sum(1 for s in alla if s.sammanslagen_av)} brunnspar sammanslagna av två delfilmer "
                          + ("(räknas en gång; delfilmerna listas i Prioritering utan rang)" if DELFILMER_SEPARAT
                             else "(räknas en gång; delfilmerna redovisas bara i den sammanslagna sträckan)"))
         if any(s.filmstatus for s in alla) else "inga"],
        ["Ny inspektion i senare fil", _tidigare_text_summa(strackor)],
        ["Höjdfel i TV3-filen", (f"{sum(1 for s in alla if s.hojdfel)} sträckor med orimliga höjder (brunnshöjder eller "
                                 f"inklinometer skiljer > {HOJD_ORIMLIG_M} m) – lutning och profil redovisas som okända")
         if any(s.hojdfel for s in alla) else "inga"],
        ["Höjdläge (markprofil/GIS från ArcMap)", (f"{sum(1 for s in strackor if s.mark)} sträckor med markhöjder, "
                                     f"{sum(1 for s in strackor if s.hojdanpassning and s.hojdanpassning['offset'] is not None and s.gis_vg_start is not None)} "
                                     f"med GIS-vattengång, {sum(1 for s in strackor if s.hojdflagga)} flaggade")
         if any(s.mark or s.gis_vg_start is not None for s in strackor) else "saknas"],
        ["GIS-data (från ArcMap)", (f"{sum(1 for s in strackor if s.gis and s.gis.get('ledning'))} sträckor med ledning i GIS, "
                                   f"{sum(1 for s in strackor if s.gisflagga)} med GIS-flagga")
         if any(s.gis for s in strackor) else "saknas"],
        ["Inspektionsgrad (GIS)", _inspektionsgrad_text(gisstat) if gisstat else "saknas (kräver GIS-data)"],
        ["", ""],
        ["Prioritetsklass", "Antal sträckor", "Andel sträckor", "Längd (m)", "Andel längd"],
    ]
    for k in "ABCDE":
        rader.append([KLASS_TEXT[k], klasser.get(k, 0),
                      klasser.get(k, 0) / len(strackor) if strackor else 0,
                      round(klass_m.get(k, 0)), klass_m.get(k, 0) / tot_m if tot_m else 0])
    rader += [["", ""], ["Poängmodell", ""],
              ["Grad 1 / 2 / 3 / 4", " / ".join(f"{GRADPOANG[g]} p" for g in (1, 2, 3, 4))],
              ["Konstruktionskoder" + ("" if KODFAKTOR else " (faktor 1,0)"),
               ", ".join(k + (f" ×{KODFAKTOR.get(k, 1.0):g}".replace(".", ",") if KODFAKTOR else "")
                         + ("".join(f" (grad {g} ×{f:g})".replace(".", ",")
                                    for (kk, g), f in (GRADFAKTOR or {}).items() if kk == k))
                         for k, v in KODER.items() if v[1] == "K")],
              ["Attributvikter", ", ".join(f"{k} {ATTRIBUT.get(a, a)} ×{v:g}".replace(".", ",")
                                           for (k, a), v in ATTRIBUTFAKTOR.items()) if ATTRIBUTFAKTOR else "inga"],
              ["Driftkoder (faktor %s)" % str(DRIFTFAKTOR).replace(".", ","),
               ", ".join(k for k, v in KODER.items() if v[1] == "D")],
              ["Löpande skador", (f"poäng × längd / {LOPANDE_ENHET_M:g} m (minst 1, högst {LOPANDE_TAK:g})"
                                  if LOPANDE_ENHET_M and LOPANDE_TAK else
                                  f"poäng × längd / {LOPANDE_ENHET_M:g} m (minst 1)" if LOPANDE_ENHET_M else
                                  "räknas en gång oavsett längd")],
              ["Index", "poäng per 100 m ledning"],
              ["Rangordning", "inom klass efter konstruktionsindex (driftindex påverkar inte)"],
              ["Klass A", (f"grad 4 på {'/'.join(sorted(GRAD4_KODER_A))}" if GRAD4_KODER_A else "konstruktionsgrad 4")
                          + f" eller konstruktionsindex ≥ {TROSKEL_A:g}"],
              ["Klass B", f"konstruktionsgrad 3 eller konstruktionsindex ≥ {TROSKEL_B:g}"],
              ["Klass C", "övriga sträckor med skador"],
              ["Klass D", "inga skadeobservationer"],
              ["Klass E", "ej bedömd (ingen inspekterad längd)"],
              ["Kort sträcka", f"sträckor < {MINLANGD:g} m normeras som {MINLANGD:g} m"]]
    if etapper is not None:
        strumpa = [s for s in strackor if s.metod == "strumpa" and s.kostnad]
        schakt = [s for s in strackor if s.metod == "schakt" and not s.kostnadsflagga.startswith("kostnad räknad")]
        tot = sum(e["kostnad"]["summa"] for e in etapper)
        rader += [[], ["Åtgärdspaket", "", "Antal", "Längd (m)", "Kostnad (kr)"],
                  ["Etapper", "sammanhängande sträckor med samma metod", len(etapper),
                   round(sum(e["langd_m"] for e in etapper)), round(tot)],
                  ["Strumpa", "kostnad enligt kostnadsfilen (schablon)", len(strumpa),
                   round(sum(s.langd for s in strumpa)), round(sum(s.kostnad["summa"] for s in strumpa))],
                  ["Framschaktning + ny brunn", "brunnar som schaktas fram (en per etappdel med bara tillsynsbrunnar)",
                   sum(len(e["framschaktade"]) for e in etapper), "", round(sum(e["kostnad"]["brunnar"] for e in etapper))],
                  ["Etablering", "fast kostnad per etapp", sum(1 for e in etapper if e["kostnad"]["etablering"]), "",
                   round(sum(e["kostnad"]["etablering"] for e in etapper))],
                  ["Schakt", "kostnad ej beräknad – beror på djup och spont", len(schakt),
                   round(sum(s.langd for s in schakt)), "ej beräknad"],
                  ["Manuellt bedömda", "sträckor med Manuell bedömning ifylld",
                   sum(1 for s in strackor if s.manuell_bedomning), "", ""]]
    for r in rader:
        ws.append(r)
    ws["A1"].font = Font(bold=True, size=14)
    rubrikrader = [i for i, r in enumerate(rader, 1) if r and r[0] in ("Prioritetsklass", "Poängmodell", "Åtgärdspaket")]
    for r in rubrikrader:
        for c in ws[r]:
            c.font = Font(bold=True)
    klassrad = rubrikrader[0] + 1
    for r in range(klassrad, klassrad + 5):
        ws.cell(r, 3).number_format = "0%"
        ws.cell(r, 5).number_format = "0%"
        ws.cell(r, 1).fill = PatternFill("solid", fgColor=KLASS_FARG[ws.cell(r, 1).value[0]])
        ws.cell(r, 1).font = Font(bold=True, color="FFFFFF" if ws.cell(r, 1).value[0] in "AB" else "000000")
    ws.column_dimensions["A"].width = 42
    ws.column_dimensions["B"].width = 40
    for col in "CDE":
        ws.column_dimensions[col].width = 14
    rad = len(rader) + 3
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
           "Konstruktionsindex (p/100 m)", "Driftindex (p/100 m)", "Totalindex (p/100 m)",
           "Konstr. maxgrad", "Drift maxgrad", "Antal skador", "Anslutningar",
           "Serviser uppströms", "Längd uppströms (m)", "Skador (kod+grad)",
           "Driftåtgärd", "Avbruten inspektion", "Inspekterad flera ggr", "Filmstatus", "Tidigare inspektion", "Relinad", "Littera rättat",
           "Svackdjup (cm)", "Svackdjup/diameter", "Svacklängd (m)", "Bakfall längd (m)", "Lutning (‰)", "Profil osäker",
           "Höjdanpassning", "Täckning min (m)", "Täckning max (m)", "Höjdflagga",
           "GIS-flagga", "Driftområde", "Lutning GIS (‰)", "Djup start (m)", "Djup slut (m)", "Anläggningsår", "Renoveringsår GIS",
           "Brunnstyp start", "Brunnstyp slut", "Etapp", "Metod", "Kostnad (kr)", "Åtgärdsflagga",
           "Manuell bedömning", "Kommentar", "Lagning (m)", "Rapport", "Videofil", "Videofil 2"]
    raknas = {id(s) for s in strackor}
    sorterade = sorterade_strackor(strackor) + sorted((s for s in visade(alla) if id(s) not in raknas), key=lambda s: (s.fil, s.nr))
    rader = []
    for rang, s in enumerate(sorterade, 1):
        if rang > len(strackor):
            rang = None           # ersatt/delfilm: ingen rang
        pa = s.profil_analys
        lut = s.lutning_promille
        h, tk = s.hojdanpassning, s.tackning
        rader.append([rang, KLASS_TEXT[s.klass], s.fil, s.nr, s.startbrunn, s.slutbrunn, s.omrade,
                      s.ledningstyp.capitalize(), s.material, s.dimension + (f"/{s.dimension2}" if s.dimension2 else ""),
                      round(s.langd, 1), s.datum, round(s.index("K"), 1), round(s.index("D"), 1),
                      round(s.index(), 1), s.maxgrad("K") or None, s.maxgrad("D") or None,
                      len(s.skador()), s.antal_anslutningar, s.serviser_uppstroms,
                      round(s.langd_uppstroms) if s.langd_uppstroms is not None else None,
                      s.sammanfattning_skador(), s.driftatgard,
                      "Ja" if s.avbruten else "", "Ja" if s.flerinspekterad else "", s.filmstatus, tidigare_text(s),
                      "Ja" if s.relinad else "",
                      s.littera_rattat,
                      round(pa["svackdjup"] * 100) if pa and pa["svackdjup"] is not None else None,
                      s.svacka_andel,
                      round(pa["svacklangd"], 1) if pa and pa["svacklangd"] is not None else None,
                      round(pa["bakfall"], 1) if pa else None,
                      round(lut, 1) if lut is not None else None,
                      ("Ja" if pa["osaker"] else "") if pa else "",
                      h["status"] if h else "",
                      round(tk["min"], 2) if tk else None, round(tk["max"], 2) if tk else None,
                      s.hojdflagga,
                      s.gisflagga, s.driftomrade,
                      round(s.gis_lutning_promille, 1) if s.gis_lutning_promille is not None else None,
                      round(s.djup_start, 2) if s.djup_start is not None else None,
                      round(s.djup_slut, 2) if s.djup_slut is not None else None,
                      s.anlaggningsar, s.gis_renoveringsar,
                      s.brunnstyp(s.startbrunn), s.brunnstyp(s.slutbrunn),
                      s.etapp if rang else None,
                      (s.metod if s.metod != "ingen" else "") if rang else "",
                      ((round(s.kostnad["summa"]) if s.kostnad else ("ej beräknad" if s.metod == "schakt" else None))
                       if rang else None),
                      s.atgardsflagga if rang else "",
                      s.manuell_bedomning, s.kommentar, s.lagning_m,
                      "Öppna rapport" if s.rapport_fil else "", s.videofil, s.videofil_b or ""])
    video_urls = [fil_url(s.video_sokvag) if s.video_sokvag else None for s in sorterade]
    video_b_urls = [fil_url(s.video_sokvag_b) if s.video_sokvag_b else None for s in sorterade]
    rapport_urls = [s.rapport_fil.replace("\\", "/") if s.rapport_fil else None for s in sorterade]   # relativ länk
    start = tabell(ws, kol, rader, {"Skador (kod+grad)": 45, "Prioritetsklass": 24, "Driftåtgärd": 28, "Rapport": 15,
                                    "Manuell bedömning": 18, "Kommentar": 30, "Höjdanpassning": 30, "Höjdflagga": 26, "GIS-flagga": 40,
                                    "Åtgärdsflagga": 40, "Kostnad": 12, "Lagning": 10, "Filmstatus": 28, "Tidigare inspektion": 44},
                   klasskol=1, lankar={kol.index("Videofil"): video_urls, kol.index("Videofil 2"): video_b_urls,
                                       kol.index("Rapport"): rapport_urls},
                   dolda=DOLDA_KOLUMNER.get("Prioritering"))
    ci = kol.index("Svackdjup/diameter") + 1
    ck = kol.index("Kostnad (kr)") + 1
    for r in range(start, ws.max_row + 1):
        ws.cell(r, ci).number_format = "0%"
        ws.cell(r, ck).number_format = "#,##0"

    # ---- Etapper ----
    if etapper is not None:
        ws = wb.create_sheet("Etapper")
        kol = ["Etapp", "Metod", "Högsta klass", "Max konstruktionsindex (p/100 m)", "Sträckor", "Längd (m)",
               "Dimensioner (mm)", "Brunnar", "Anslutningar (hattar)", "Lagning (m)",
               "Framschaktade brunnar", "Schakta fram (manuellt)", "Serviser uppströms",
               "Kostnad strumpa (kr)", "Kostnad hattar (kr)", "Kostnad lagning (kr)",
               "Kostnad brunnar (kr)", "Etablering (kr)", "Kostnad totalt (kr)",
               "Schakt – kostnad ej beräknad (m)", "Medtagna för sammanhang", "Flaggor", "Sträckor (lista)"]
        rader = []
        for e in etapper:
            st = e["strackor"]
            brunnar = sorted({b for s in st for b in (s.startbrunn, s.slutbrunn)})
            dims = sorted({s.dimension_mm for s in st if s.dimension_mm})
            serv = [s.serviser_uppstroms for s in st if s.serviser_uppstroms is not None]
            flaggor = sorted({f for s in st for f in s.atgardsflagga.split("; ")
                              if f and not f.startswith("medtagen") and not f.startswith("kostnad räknad")})
            rader.append([e["nr"], e["metod"], min(s.gallande_klass for s in st),
                          round(max(s.index("K") for s in st), 1), len(st), round(e["langd_m"]),
                          ", ".join(str(d) for d in dims), ", ".join(brunnar),
                          sum(s.antal_anslutningar for s in st if s.metod == "strumpa"),
                          sum(s.lagning_m or 0 for s in st) or None,
                          ", ".join(e["framschaktade"]), ", ".join(e["framschakt_manuell"]),
                          max(serv) if serv else None,
                          round(e["kostnad"]["strumpa"]) or None, round(e["kostnad"]["hattar"]) or None,
                          round(e["kostnad"]["lagning"]) or None, round(e["kostnad"]["brunnar"]) or None,
                          round(e["kostnad"]["etablering"]) or None,
                          round(e["kostnad"]["summa"]) if e["metod"] == "strumpa" else "ej beräknad",
                          round(e["schakt_m"]) or None,
                          sum(1 for s in st if s.etapp_flagga) or None,
                          "; ".join(flaggor),
                          "; ".join(f"{s.fil} nr {s.nr}: {s.startbrunn}→{s.slutbrunn} ({s.gallande_klass}, {s.langd:.0f} m)"
                                    for s in st)])
        start = tabell(ws, kol, rader, {"Brunnar": 40, "Framschaktade brunnar": 24, "Schakta fram (manuellt)": 22,
                                        "Flaggor": 45, "Sträckor (lista)": 60, "Metod": 10},
                       klasskol=2, dolda=DOLDA_KOLUMNER.get("Etapper"))
        for r in range(start, ws.max_row + 1):
            for c in range(kol.index("Kostnad strumpa (kr)") + 1, kol.index("Kostnad totalt (kr)") + 2):
                ws.cell(r, c).number_format = "#,##0"

    # ---- Observationer ----
    ws = wb.create_sheet("Observationer")
    kol = ["Fil", "Nr", "Startbrunn", "Slutbrunn", "Prioritetsklass", "Läge (m)", "Tid i film", "Kod",
           "Beskrivning", "Typ", "Grad", "Poäng", "Löpande", "Löpande längd (m)",
           "Klocka från", "Klocka till", "Vattennivå (%)", "Bild", "Kommentar", "Videofil"]
    rader, bild_urls, video_urls = [], [], []
    for s in sorterade:
        if s.sammanslagen_av and DELFILMER_SEPARAT:
            continue                       # observationerna finns redan under delfilmerna
        for o in s.observationer:
            if not (o.kod or o.infokod):
                continue
            bs = o.bild_sokvag or (o.bild_b_sokvag if not o.bild else None)
            bild_urls.append(fil_url(bs) if bs else None)
            fran_b = s.videofil_b and s.sammanslagning and o.film_nr == s.sammanslagning["b_nr"]
            vf, vs = (s.videofil_b, s.video_sokvag_b) if fran_b else (s.videofil, s.video_sokvag)
            video_urls.append(fil_url(vs) if vs else None)
            rader.append([s.fil, s.nr, s.startbrunn, s.slutbrunn, KLASS_TEXT[s.klass], o.lage, o.tid,
                          o.kod or o.infokod, o.beskrivning(),
                          {"K": "Konstruktion", "D": "Drift", "I": "Info"}.get(o.typ, "Info"),
                          o.grad, o.poang or None, o.lopande, o.lopande_langd,
                          o.klocka_fran, o.klocka_till, o.vattenniva, o.bild or o.bild_b, o.kommentar, vf])
    tabell(ws, kol, rader, {"Beskrivning": 50, "Kommentar": 30, "Prioritetsklass": 24}, klasskol=4,
           lankar={kol.index("Bild"): bild_urls, kol.index("Videofil"): video_urls},
           dolda=DOLDA_KOLUMNER.get("Observationer"))

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
        g["p"] += s.poang("K")
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
           "Konstruktionsindex (p/100 m)", "Klass A", "Klass B", "Klass C", "Klass D", "Klass E",
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
    start = tabell(ws, kol, rader, {"Område": 30})
    for r in range(start, ws.max_row + 1):
        ws.cell(r, 14).number_format = "0%"

    # ---- Material & dimension ----
    ws = wb.create_sheet("Material")
    grp = defaultdict(lambda: {"n": 0, "m": 0.0, "p": 0.0, "A": 0, "B": 0, "urspr": Counter()})
    for s in strackor:
        g = grp[(s.material_grupp, s.ledningstyp.capitalize())]
        g["n"] += 1
        g["m"] += s.langd
        g["p"] += s.poang("K")
        g["urspr"][s.material or "Okänt"] += 1
        if s.klass in "AB":
            g[s.klass] += 1

    def ursprung(mat, g):
        """Ursprungsmaterial för relinade rör, t.ex. 'Betong (16)'. Tomt för övriga."""
        if not RELINAD_SOM_MATERIAL or mat != RELINAD_MATERIAL_NAMN:
            return ""
        return ", ".join(f"{m} ({n})" for m, n in g["urspr"].most_common())

    kol = ["Material", "Ursprungsmaterial", "Ledningstyp", "Sträckor", "Längd (m)",
           "Konstruktionsindex (p/100 m)", "Klass A", "Klass B", "Andel A+B"]
    rader = [[m, ursprung(m, g), t, g["n"], round(g["m"]), round(g["p"] / g["m"] * 100, 1) if g["m"] else 0,
              g["A"], g["B"], (g["A"] + g["B"]) / g["n"]] for (m, t), g in sorted(grp.items())]
    start = tabell(ws, kol, rader)
    for r in range(start, ws.max_row + 1):
        ws.cell(r, 9).number_format = "0%"

    # ---- Inspektionsgrad per driftområde (GIS-data) ----
    if gisstat:
        ws = wb.create_sheet("Inspektionsgrad")
        kol = ["Driftområde", "Ledningstyp", "Ledningar i GIS", "Längd i GIS (m)", "Filmade ledningar",
               "Filmad längd (m)", "Andel filmad", "Klass A (st)", "Klass B (st)", "Längd A+B (m)",
               "Andel A+B av filmat"]
        rader = [[r["omrade"], r["ledningstyp"], r["ledningar"], round(r["langd_m"]), r["filmade"],
                  round(r["filmad_m"]), r["andel"], r["A"], r["B"], round(r["AB_m"]), r["andel_AB"]]
                 for r in gisstat["rader"]]
        n_utan, m_utan = gisstat["utan_gis"]
        if n_utan:
            rader.append(["(filmat utan ledning i GIS)", "", None, None, n_utan, round(m_utan), None,
                          None, None, None, None])
        start = tabell(ws, kol, rader, {"Driftområde": 18, "Ledningstyp": 14, "Andel A+B av filmat": 14})
        for r in range(start, ws.max_row + 1):
            ws.cell(r, 7).number_format = "0%"
            ws.cell(r, 11).number_format = "0%"
            if ws.cell(r, 2).value == "alla":
                for c in range(1, 12):
                    ws.cell(r, c).font = Font(bold=True)
        ws = wb.create_sheet("Ej inspekterat")
        kol = ["Driftområde", "Från brunn", "Till brunn", "Längd (m)", "Dim (mm)", "Material", "Ledningstyp",
               "Anläggningsår"]
        rader = [[e["omrade"], e["fran"], e["till"], round(e["langd_m"], 1) if e["langd_m"] is not None else None,
                  e["dimension"], e["material"], e["ledningstyp"], e["anlaggningsar"]]
                 for e in gisstat["ej_inspekterat"]]
        tabell(ws, kol, rader, {"Driftområde": 18, "Från brunn": 16, "Till brunn": 16})

    wb.save(path)


def _tidigare_text_summa(strackor: list[Stracka]) -> str:
    """Sammanfattningsrad: hur många gällande sträckor som har en äldre inspektion i en annan fil."""
    med = [s for s in strackor if s.tidigare]
    if not med:
        return "inga"
    n = Counter(s.tidigare["utveckling"] for s in med)
    return (f"{len(med)} sträckor ersätter en äldre inspektion i annan fil: {n['förvärrad']} förvärrade, "
            f"{n['förbättrad']} förbättrade, {n['oförändrad']} oförändrade (kolumn Tidigare inspektion)")


def _inspektionsgrad_text(gisstat: dict) -> str:
    tot = next((r for r in gisstat["rader"] if r["omrade"] == "alla" and r["ledningstyp"] == "alla"), None)
    if not tot or not tot["langd_m"]:
        return "inga ledningar i GIS-data"
    omr = [r for r in gisstat["rader"] if r["omrade"] != "alla" and r["ledningstyp"] == "alla"]
    def ab(r):
        return f" varav A+B {r['andel_AB']:.0%}" if r["andel_AB"] is not None else ""

    def tusen(x):
        return f"{x:,.0f}".replace(",", " ")
    delar = "; ".join(f"{r['omrade']} {r['andel']:.0%}{ab(r)}" for r in omr[:12])
    return (f"{tusen(tot['filmad_m'])} av {tusen(tot['langd_m'])} m filmade ({tot['andel']:.0%}{ab(tot)})"
            + (f" – per område: {delar}" if omr else ""))


def etapper_for_karta(etapper: list[dict] | None) -> list[dict] | None:
    """Etapplistan i kartunderlag.json (för ArcMap-verktyget Etapplager): en post per etapp med metod,
    längd, kostnad, brunnar att schakta fram, flaggor och sträckorna (fil + nr + brunnspar)."""
    if not etapper:
        return None
    ut = []
    for e in etapper:
        st = e["strackor"]
        flaggor = sorted({f for s in st for f in s.atgardsflagga.split("; ")
                          if f and not f.startswith("medtagen") and not f.startswith("kostnad räknad")})
        serv = [s.serviser_uppstroms for s in st if s.serviser_uppstroms is not None]
        ut.append({
            "nr": e["nr"], "metod": e["metod"],
            "hogsta_klass": min(s.gallande_klass for s in st),
            "max_konstruktionsindex": round(max(s.index("K") for s in st), 1),
            "antal_strackor": len(st), "langd_m": round(e["langd_m"], 1),
            "schakt_m": round(e["schakt_m"], 1),
            "dimensioner": sorted({s.dimension_mm for s in st if s.dimension_mm}),
            "brunnar": sorted({b for s in st for b in (s.startbrunn, s.slutbrunn)}),
            "framschaktade": list(e["framschaktade"]),
            "framschakt_manuell": list(e["framschakt_manuell"]),
            "serviser_uppstroms": max(serv) if serv else None,
            "kostnad_kr": round(e["kostnad"]["summa"]) if e["metod"] == "strumpa" else None,
            "kostnad": {k: round(v) for k, v in e["kostnad"].items()},
            "flaggor": flaggor,
            "strackor": [{"fil": s.fil, "nr": s.nr, "startbrunn": s.startbrunn, "slutbrunn": s.slutbrunn,
                          "bedomning": s.gallande_klass, "langd_m": round(s.langd, 1)} for s in st],
        })
    return ut


def _fri_ande_xy(s: Stracka) -> list | None:
    """Koordinaten för GIS-ledningens fria ände (grenrörsänden) när sträckan kopplats så, annars None."""
    g = s.gis and s.gis.get("ledning")
    if not g or not (s.gis.get("grenror") or {}):
        return None
    xy = g.get("till_xy") if not g.get("till") else g.get("fran_xy")
    return [round(float(xy[0]), 3), round(float(xy[1]), 3)] if isinstance(xy, (list, tuple)) and len(xy) == 2 else None


def _gis_ledning_id(s: Stracka) -> dict | None:
    """{lager, oid, del} för den kopplade GIS-ledningen (inte syntetiska vägar), annars None."""
    g = s.gis and s.gis.get("ledning")
    if not g or g.get("_syntetisk") or g.get("oid") is None:
        return None
    return {"lager": g.get("lager"), "oid": g.get("oid"), "del": g.get("del")}


def skriv_kartunderlag(strackor: list[Stracka], path: str, gisstat: dict | None = None,
                       etapper: list[dict] | None = None) -> int:
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
            "serviser_uppstroms": s.serviser_uppstroms,
            "langd_uppstroms_m": s.langd_uppstroms,
            "manuell_bedomning": s.manuell_bedomning,
            "gallande_bedomning": s.gallande_klass,
            "etapp": s.etapp,
            "metod": s.metod if s.metod != "ingen" else None,
            "kostnad_kr": round(s.kostnad["summa"]) if s.kostnad else None,
            "lagning_m": s.lagning_m,
            "brunnstyp_start": s.brunnstyp(s.startbrunn),
            "brunnstyp_slut": s.brunnstyp(s.slutbrunn),
            "atgardsflagga": s.atgardsflagga,
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
            "filmstatus": s.filmstatus or None,
            "sammanslagen_av": s.sammanslagen_av or None,
            "tidigare_inspektion": ({"fil": os.path.basename(s.tidigare["fil"]), "nr": s.tidigare["nr"],
                                     "datum": s.tidigare["datum"], "bedomning": s.tidigare["klass"],
                                     "konstruktionsindex": round(s.tidigare["index"], 1),
                                     "langd_m": round(s.tidigare["langd"], 1), "antal_skador": s.tidigare["antal_skador"],
                                     "utveckling": s.tidigare["utveckling"]} if s.tidigare else None),
            "littera_rattat": s.littera_rattat,
            "svackdjup_cm": round(pa["svackdjup"] * 100) if pa and pa["svackdjup"] is not None else None,
            "svacklangd_m": round(pa["svacklangd"], 1) if pa and pa["svacklangd"] is not None else None,
            "bakfall_m": round(pa["bakfall"], 1) if pa else None,
            # Positioner nedan är kamerans (m från utgångsbrunnen), som i protokollet.
            # ArcMap-skriptet lägger ut dem längs kartlinjen från utgångsbrunnen, skalat med
            # kartlängd/filmlängd (ingen skalning vid avbruten inspektion).
            "svackpos_m": round(pa["svackpos"], 1) if pa and pa["svackpos"] is not None else None,
            "svacka_andel": s.svacka_andel,
            "bakfall_segment": pa["bakfall_segment"] if pa else [],
            "lutning_promille": round(lut, 1) if lut is not None else None,
            "profil_osaker": bool(pa["osaker"]) if pa else None,
            "tv3_fil": s.tv3_sokvag,
            "rapport": s.rapport_fil.replace("\\", "/") if s.rapport_fil else None,
            "videofil": s.videofil,
            "video_sokvag": s.video_sokvag,
            "videofil_b": s.videofil_b,
            "video_sokvag_b": s.video_sokvag_b,
            "hojdanpassning": s.hojdanpassning["status"] if s.hojdanpassning else None,
            "hojd_offset_m": (round(s.hojdanpassning["offset"], 2)
                              if s.hojdanpassning and s.hojdanpassning["offset"] is not None else None),
            "gis_vg_start": s.gis_vg_start,
            "gis_vg_slut": s.gis_vg_slut,
            "tackning_min_m": round(s.tackning["min"], 2) if s.tackning else None,
            "tackning_max_m": round(s.tackning["max"], 2) if s.tackning else None,
            "hojdflagga": s.hojdflagga,
            "hojdfel": s.hojdfel or None,
            "gis_flagga": s.gisflagga or None,
            "grenror": (s.gis or {}).get("grenror", {}).get("sida") if s.gis and s.gis.get("grenror") else None,
            "grenror_xy": _fri_ande_xy(s),
            "gis_ledning": _gis_ledning_id(s),
            "gis_lutning_promille": round(s.gis_lutning_promille, 1) if s.gis_lutning_promille is not None else None,
            "djup_start_m": round(s.djup_start, 2) if s.djup_start is not None else None,
            "djup_slut_m": round(s.djup_slut, 2) if s.djup_slut is not None else None,
            "anlaggningsar": s.anlaggningsar,
            "renoveringsar_gis": s.gis_renoveringsar,
            "driftomrade": s.driftomrade or None,
        })

    data = {
        "version": 1,
        "kalla": "tv3_analys.py",
        "genererad": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "klasser": KLASS_TEXT,
        # Mappar där film och bilder hittades (tv3_pptx.py använder dem när den körs fristående)
        "mediakataloger": sorted({d for s in strackor for d in s.media_kataloger}
                                 | {os.path.dirname(s.video_sokvag) for s in strackor if s.video_sokvag}),
        "bildkataloger": sorted({d for s in strackor for d in s.bild_kataloger}
                                | {os.path.dirname(o.bild_sokvag) for s in strackor for o in s.observationer
                                   if o.bild_sokvag}),
        "strackor": poster,
        # Inspektionsgrad per driftområde och GIS-ledningar utan film (kräver gis: i listfilen)
        "inspektionsgrad": gisstat["rader"] if gisstat else None,
        "ej_inspekterat": gisstat["ej_inspekterat"] if gisstat else None,
        # Etapper (åtgärdspaket) för ArcMap-verktyget Etapplager
        "etapper": etapper_for_karta(etapper),
    }
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=1)
    return len(poster)


def sorterade_strackor(strackor):
    ordn = {"A": 0, "B": 1, "C": 2, "D": 3, "E": 4}
    # Rangordning inom klass efter konstruktionsindex – driftskador (rötter, sediment …)
    # påverkar varken klass eller ordning, de redovisas bara som driftindex/driftåtgärd.
    return sorted(strackor, key=lambda s: (ordn[s.klass], -s.index("K"), -s.maxgrad("K"), s.fil, s.nr))


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
        ax.text(k_idx[i] + d_idx[i] + max(k_idx) * 0.01, i, f"{s.index('K'):.0f}  ·  {KLASS_TEXT[s.klass][:1]}",
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
    mats = sorted((m for m in grp if sum(grp[m].values()) > 0), key=lambda m: -sum(grp[m].values()))
    fig, ax = plt.subplots(figsize=(8, 0.6 * max(len(mats), 1) + 1.8))
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


def _spara_rapportbild(fig, path: str) -> None:
    """Sparar en matplotlib-figur för rapporten: RAPPORT_BILD_DPI och, om RAPPORT_BILD_PALETT,
    som palett-PNG med 256 färger (mycket mindre fil, samma utseende på linjediagram)."""
    fig.savefig(path, dpi=RAPPORT_BILD_DPI)
    if not RAPPORT_BILD_PALETT:
        return
    try:
        from PIL import Image as PILImage
        with PILImage.open(path) as im:
            im.convert("RGB").quantize(colors=256).save(path, "PNG", optimize=False)
    except Exception:
        pass                                   # RGB-bilden ligger kvar


SCHEMA_BREDD_TUM = 7.4
SCHEMA_HOJD_TUM = 2.5          # vid upp till tre rader löpande skador; växer med fler rader
SCHEMA_RADHOJD = 0.30          # dataenheter per rad löpande skador (band + etikett)


def dk(text) -> str:
    """Decimalkomma i rapporttext: '79.05 m' → '79,05 m' (bara punkt mellan två siffror)."""
    return re.sub(r"(?<=\d)\.(?=\d)", ",", str(text))


def _packa_band(band: list[tuple[float, float]]) -> list[int]:
    """Rad per band (0 = närmast röret) så att inga band eller etiketter överlappar i samma rad.
    band = [(från, till inkl. etikett)] i bandens ordning; första lediga raden tas."""
    rader: list[float] = []            # längst högra upptagna x per rad
    ut = []
    for a, b in band:
        for i, slut in enumerate(rader):
            if a > slut:
                rader[i] = b
                ut.append(i)
                break
        else:
            rader.append(b)
            ut.append(len(rader) - 1)
    return ut


def rita_schema(s: Stracka, path: str) -> tuple[float, float]:
    """Schematisk bild av sträckan: rör med anslutningar, skador (färg efter grad) och löpande skador.
    Returnerar figurens (bredd, höjd) i tum – höjden växer när band eller anslutningsetiketter behöver
    fler rader. Allt som kan krocka packas i rader (`_packa_band`): löpande skador ovanför röret,
    anslutningsetiketter per sida och romber för punktskador som ligger inom samma bredd."""
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    from matplotlib.patches import FancyBboxPatch

    L = max(s.langd, 1.0)
    per_pt = L * 1.12 / (SCHEMA_BREDD_TUM * 72)        # dataenheter (x) per punkt
    tecken6 = per_pt * 3.9                              # bredd per tecken vid 6 pt
    tecken55 = per_pt * 3.6                             # vid 5,5 pt
    bla = "#2a78d6"

    # löpande skador: rad per band så de inte hamnar på varandra (etikettbredden räknas med)
    lopande = [o for o in s.observationer if o.raknas and o.lopande.startswith("A") and o.lopande_langd]
    lopande.sort(key=lambda o: o.lage)
    intervall = []
    for o in lopande:
        text = dk(f"{o.kod}{o.grad} {o.lopande_langd:.1f} m")
        intervall.append((o.lage, max(o.lage + o.lopande_langd, o.lage + len(text) * tecken6) + L * 0.015))
    band_rad = _packa_band(intervall)

    # anslutningar – sida enligt klockposition sett i inspektionsriktningen:
    # kl 7–11 = vänster (ritas ovanför röret), kl 1–5 = höger (under röret), kl 12/6 = hjässa/botten (på röret)
    ansl = {"vänster": [], "höger": [], "": []}
    for o in sorted((o for o in s.observationer if o.infokod in ("AS", "AG")), key=lambda o: o.lage):
        etikett = dk(f"{o.lage:.1f} m") + (f" kl {o.klocka_fran.lstrip('0')}" if o.klocka_fran else "")
        sida = anslutning_sida(o.klocka_fran)
        ansl["vänster" if sida == "vänster" else "höger" if sida == "höger" else ""].append((o, etikett))
    # avbrott (KAM) packas med hjässa/botten-etiketterna så texten inte hamnar på en anslutning i samma punkt
    for o in s.observationer:
        if o.kod == "KAM":
            ansl[""].append((o, "inspektion avbruten"))
    ansl[""].sort(key=lambda par: par[0].lage)
    ansl_rad = {}
    halv = [len(e) * tecken55 / 2 + L * 0.008 for _, e in ansl["vänster"]]
    ansl_rad["vänster"] = _packa_band([(o.lage - h, o.lage + h) for (o, _), h in zip(ansl["vänster"], halv)])
    # höger och hjässa/botten har etiketterna under röret – packas tillsammans så de inte krockar
    under = sorted([(o, e, "höger") for o, e in ansl["höger"]] + [(o, e, "") for o, e in ansl[""]],
                   key=lambda t: t[0].lage)
    halv = [len(e) * tecken55 / 2 + L * 0.008 for _, e, _ in under]
    rad_under = _packa_band([(o.lage - h, o.lage + h) for (o, _, _), h in zip(under, halv)])
    ansl["höger"] = [(o, e) for o, e, sida in under if sida == "höger"]
    ansl[""] = [(o, e) for o, e, sida in under if sida == ""]
    ansl_rad["höger"] = [r for (o, e, sida), r in zip(under, rad_under) if sida == "höger"]
    ansl_rad[""] = [r for (o, e, sida), r in zip(under, rad_under) if sida == ""]
    rader_ovan = max(ansl_rad["vänster"], default=-1) + 1
    rader_under = max(rad_under, default=-1) + 1

    # punktskador: romber inom samma bredd läggs i egna rader innanför röret
    punkt = sorted((o for o in s.observationer if o.raknas and not (o.lopande.startswith("A") and o.lopande_langd)),
                   key=lambda o: o.lage)
    romb = per_pt * 7.5
    punkt_rad = _packa_band([(o.lage - romb / 2, o.lage + romb / 2) for o in punkt])

    ANSL_RAD = 0.28
    y_band0 = 0.42 + ANSL_RAD * max(rader_ovan, 1) + 0.25           # banden börjar ovanför etiketterna
    antal_band = max(1, max(band_rad) + 1 if band_rad else 0)
    y_topp = y_band0 + SCHEMA_RADHOJD * antal_band + 0.55           # plats för band + förklaring
    y_skala = -(0.42 + ANSL_RAD * max(rader_under, 1) + 0.3)        # skalan under alla etiketter
    y_pil = y_skala - 0.75
    y_botten = y_skala - (1.25 if s.sammanslagning else 0.95)
    hojd_tum = SCHEMA_HOJD_TUM * (y_topp - y_botten) / 4.3
    fig, ax = plt.subplots(figsize=(SCHEMA_BREDD_TUM, hojd_tum))
    ax.set_xlim(-L * 0.06, L * 1.06)
    ax.set_ylim(y_botten, y_topp)
    ax.axis("off")
    # röret – vid sammanslagning av två avbrutna filmer ritas delfilmernas täckning i olika ton,
    # överlapp skrafferat och en streckad skarvlinje där filmerna möts
    sm = s.sammanslagning
    ax.add_patch(FancyBboxPatch((0, -0.25), L, 0.5, boxstyle="round,pad=0,rounding_size=0.02",
                                fc="#e8e8e8", ec="#7f7f7f", lw=1.2))
    if sm:
        from matplotlib.patches import Rectangle
        a_slut = min(sm["a_langd"], L)
        b_start = max(0.0, L - sm["b_langd"])
        ax.add_patch(Rectangle((b_start, -0.25), L - b_start, 0.5, fc="#d6dde8", ec="none", zorder=1.5))
        if sm["overlapp"] > 0.5:
            ax.add_patch(Rectangle((b_start, -0.25), a_slut - b_start, 0.5, fc="none", ec="#7f7f7f",
                                   hatch="////", lw=0, zorder=1.6))
        elif sm["overlapp"] < -0.5:
            ax.add_patch(Rectangle((a_slut, -0.25), b_start - a_slut, 0.5, fc="white", ec="none", zorder=1.5))
        for x in sorted({round(a_slut, 2), round(b_start, 2)}):
            ax.plot([x, x], [-0.32, 0.32], color="#333333", lw=1.0, ls=(0, (3, 2)), zorder=4)
    # brunnar – större än rörets diameter; vid avbruten inspektion nåddes inte slutbrunnen (ihålig)
    for x, namn, ha, nadd in ((0, s.fran_brunn, "right", True), (L, s.till_brunn, "left", not s.ofullstandig)):
        ax.plot(x, 0, "o", ms=26, mfc="#bfbfbf" if nadd else "white", mec="#4d4d4d", mew=1.2,
                ls="none", zorder=5, alpha=1.0 if nadd else 0.6)
        ax.text(x + (-0.03 if ha == "right" else 0.03) * L, 0.55, namn + ("" if nadd else " (ej nådd)"),
                ha=ha, va="bottom", fontsize=8, fontweight="bold", color="#000000" if nadd else "#7f7f7f")
    # löpande skador som band
    for o, r in zip(lopande, band_rad):
        y = y_band0 + SCHEMA_RADHOJD * r
        ax.plot([o.lage, o.lage + o.lopande_langd], [y, y], lw=4, color=GRAD_FARG_HEX.get(o.grad, "#888"),
                solid_capstyle="butt", alpha=0.9)
        ax.text(o.lage, y + 0.08, dk(f"{o.kod}{o.grad} {o.lopande_langd:.1f} m"), fontsize=6, va="bottom")
    # punktskador
    for o, r in zip(punkt, punkt_rad):
        y = (0, 0.15, -0.15)[r % 3]
        ax.plot(o.lage, y, "D", ms=7, mfc=GRAD_FARG_HEX.get(o.grad, "#888"), mec="white", mew=0.8, zorder=6)
    # anslutningar
    for (o, etikett), r in zip(ansl["vänster"], ansl_rad["vänster"]):
        y = 0.42 + ANSL_RAD * r
        ax.plot(o.lage, y, "v", ms=7, mfc=bla, mec="white", zorder=6)
        ax.text(o.lage, y + 0.13, etikett, ha="center", va="bottom", fontsize=5.5, color=bla)
    for (o, etikett), r in zip(ansl["höger"], ansl_rad["höger"]):
        y = -0.42 - ANSL_RAD * r
        ax.plot(o.lage, y, "^", ms=7, mfc=bla, mec="white", zorder=6)
        ax.text(o.lage, y - 0.13, etikett, ha="center", va="top", fontsize=5.5, color=bla)
    for (o, etikett), r in zip(ansl[""], ansl_rad[""]):
        if o.kod == "KAM":      # avbrott – rött kryss där kameran stannade
            ax.plot(o.lage, 0, "X", ms=10, mfc="#c00000", mec="white", mew=0.8, zorder=7)
            farg = "#c00000"
        else:
            ax.plot(o.lage, 0, "s", ms=6, mfc=bla, mec="white", zorder=6)
            farg = bla
        ax.text(o.lage, -0.42 - ANSL_RAD * r, etikett, ha="center", va="top", fontsize=5.5, color=farg, zorder=7)
    # meterskala – egen linje under anslutningarna, siffrorna direkt under ticksen
    ax.plot([0, L], [y_skala, y_skala], color="#7f7f7f", lw=0.8)
    for x in range(0, int(L) + 1, max(1, int(L // 8) or 1)):
        ax.plot([x, x], [y_skala, y_skala - 0.1], color="#7f7f7f", lw=0.6)
        ax.text(x, y_skala - 0.13, f"{x}", ha="center", va="top", fontsize=6, color="#7f7f7f")
    if sm:
        a_slut = min(sm["a_langd"], L)
        b_start = max(0.0, L - sm["b_langd"])
        ax.annotate("", xy=(a_slut, y_pil), xytext=(0, y_pil), arrowprops=dict(arrowstyle="->", color=bla, lw=1.2))
        ax.annotate("", xy=(b_start, y_pil), xytext=(L, y_pil), arrowprops=dict(arrowstyle="->", color="#4a5d7a", lw=1.2))
        ax.text(a_slut / 2, y_pil - 0.12, dk(f"film nr {sm['a_nr']} från {sm['a_fran']}, 0–{a_slut:.1f} m"),
                ha="center", va="top", fontsize=6.5, color=bla)
        ax.text((b_start + L) / 2, y_pil - 0.12, dk(f"film nr {sm['b_nr']} från {sm['b_fran']}, {b_start:.1f}–{L:.1f} m"),
                ha="center", va="top", fontsize=6.5, color="#4a5d7a")
        skarv = (dk(f"överlapp {sm['overlapp']:.1f} m") if sm["overlapp"] > 0.5 else
                 dk(f"lucka {-sm['overlapp']:.1f} m") if sm["overlapp"] < -0.5 else dk(f"skarv vid {a_slut:.1f} m"))
        ax.text(L / 2, y_pil + 0.14, f"sammanslagen av två avbrutna filmer – {skarv}, position (m) mätt från {s.fran_brunn}",
                ha="center", va="bottom", fontsize=6.5, color="#333333")
    else:
        ax.annotate("", xy=(L * 0.10, y_pil), xytext=(L * 0.0, y_pil),
                    arrowprops=dict(arrowstyle="->", color=bla, lw=1.2))
        ax.text(L * 0.11, y_pil, f"inspektionsriktning ({s.riktning.lower()}), position (m) mätt från {s.fran_brunn}",
                va="center", fontsize=7, color=bla)
    # legend
    y_leg = y_topp - 0.2
    for i, g in enumerate((1, 2, 3, 4)):
        ax.plot(L * (0.55 + 0.11 * i), y_leg, "D", ms=6, mfc=GRAD_FARG_HEX[g], mec="white")
        ax.text(L * (0.565 + 0.11 * i), y_leg, f"grad {g}", va="center", fontsize=6.5)
    ax.plot(L * 0.02, y_leg, "v", ms=6, mfc=bla, mec="white")
    ax.text(L * 0.035, y_leg, "anslutning vänster (kl 7–11)", va="center", fontsize=6.5)
    ax.plot(L * 0.3, y_leg, "^", ms=6, mfc=bla, mec="white")
    ax.text(L * 0.315, y_leg, "höger (kl 1–5)", va="center", fontsize=6.5)
    ax.text(L * 0.02, y_leg - 0.22, "vänster/höger sett i inspektionsriktningen", va="center", fontsize=6,
            color="#52514e")
    fig.tight_layout(pad=0.2)
    _spara_rapportbild(fig, path)
    plt.close(fig)
    return SCHEMA_BREDD_TUM, hojd_tum


PROFIL_FIG = (7.4, 3.2)                 # tum
PROFIL_AX = (0.10, 0.17, 0.87, 0.70)    # axelns läge i figuren (andel av bredd/höjd)
LANGDSKALOR = [10, 20, 25, 40, 50, 75, 100, 150, 200, 250, 300, 400, 500, 750, 1000, 1500, 2000]
HOJDSKALOR = [1, 2, 5, 10, 20, 25, 50, 100, 200, 500, 1000]


def _dk_axlar(ax) -> None:
    """Decimalkomma på axlarna och lika många decimaler på alla ticks (steget avgör: 22,8 och 23,0)."""
    from matplotlib.ticker import FuncFormatter
    for axel, ticks in ((ax.xaxis, ax.get_xticks()), (ax.yaxis, ax.get_yticks())):
        steg = min((b - a for a, b in zip(ticks, ticks[1:]) if b > a), default=1.0)
        dec = max(0, -int(math.floor(math.log10(steg) + 1e-9)))
        axel.set_major_formatter(FuncFormatter(lambda v, _, d=dec: dk(f"{v:.{d}f}")))


def rita_profil(s: Stracka, path: str, bild_bredd_mm: float) -> bool:
    """Inklinometerprofil ritad i exakt skala. Höjdpunkten läggs alltid till vänster så att
    profilen lutar ned mot höger. Returnerar False om profil saknas."""
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt

    if len(s.profil) >= 3:
        prof = s.profil_korrigerad()
    elif s.hojdanpassning and s.mark and s.langd > 0:
        # ingen inklinometer men GIS-vattengång och markyta: rät linje mellan brunnarna
        prof = [(0.0, s.ledningshojd(0.0)), (s.langd, s.ledningshojd(s.langd))]
        if any(zz is None for _, zz in prof):
            return False
    else:
        return False
    L_prof = prof[-1][0]
    x = [p[0] for p in prof]
    z = [p[1] for p in prof]
    mark = s.mark_i_filmens_axel()
    vanster, hoger = s.fran_brunn, ("avbrott" if s.ofullstandig else s.till_brunn)
    speglad = z[-1] > z[0]
    if speglad:                            # spegla så att den höga änden hamnar till vänster
        L = x[-1]
        x = [L - xi for xi in x][::-1]
        z = z[::-1]
        mark = [(L - xm, zm) for xm, zm in mark][::-1]
        vanster, hoger = hoger, vanster
    x0, z0, x1, z1 = x[0], z[0], x[-1], z[-1]
    zalla = z + [zm for _, zm in mark]

    # fysisk storlek på axeln när bilden skrivs ut med bredden bild_bredd_mm
    bild_hojd_mm = bild_bredd_mm * PROFIL_FIG[1] / PROFIL_FIG[0]
    ax_w_mm = bild_bredd_mm * PROFIL_AX[2]
    ax_h_mm = bild_hojd_mm * PROFIL_AX[3]
    xspann = max(x1 - x0, 0.5) * 1.03
    zspann = max(max(zalla) - min(zalla), 0.05) * 1.15
    langdskala = next((k for k in LANGDSKALOR if xspann * 1000 / k <= ax_w_mm), LANGDSKALOR[-1])
    hojdskala = next((k for k in HOJDSKALOR if zspann * 1000 / k <= ax_h_mm), HOJDSKALOR[-1])
    xvidd = ax_w_mm * langdskala / 1000    # m som ryms i axeln
    zvidd = ax_h_mm * hojdskala / 1000
    zmitt = (max(zalla) + min(zalla)) / 2

    fig = plt.figure(figsize=PROFIL_FIG)
    ax = fig.add_axes(PROFIL_AX)
    ax.set_xlim(x0 - (xvidd - (x1 - x0)) / 2, x0 - (xvidd - (x1 - x0)) / 2 + xvidd)
    ax.set_ylim(zmitt - zvidd / 2, zmitt + zvidd / 2)
    ax.plot([x0, x1], [z0, z1], color="#4d4d4d", lw=0.9, ls="--",
            label="rät linje mellan ändpunkterna" if s.ofullstandig else "rät linje mellan brunnarna")
    # Höjderna i TV3-filen är avrundade till hela cm, vilket ger en trappstegsformad linje på flacka
    # ledningar. Linjen jämnas ut med ett glidande medelvärde (±0,3 m) enbart för uppritningen;
    # svacka, bakfall och lutning beräknas på rådata.
    # Glidande fönster över de positionssorterade punkterna (O(n), inte O(n²) – en lång profil
    # med tusentals punkter tog annars sekunder per rapport)
    ordn = sorted(range(len(x)), key=lambda i: x[i])
    xs_, zs_ = [x[i] for i in ordn], [z[i] for i in ordn]
    zj_sort, lo, hi, summa = [], 0, 0, 0.0
    for i, xi in enumerate(xs_):
        while hi < len(xs_) and xs_[hi] <= xi + 0.3 + 1e-9:     # ±0,3 m inklusive, entydigt vid cm-gränser
            summa += zs_[hi]; hi += 1
        while xs_[lo] < xi - 0.3 - 1e-9:
            summa -= zs_[lo]; lo += 1
        zj_sort.append(summa / (hi - lo))
    zj = [0.0] * len(x)
    for k, i in enumerate(ordn):
        zj[i] = zj_sort[k]
    ax.plot(x, zj, color="#2a78d6", lw=1.6,
            label="uppmätt profil (utjämnad, cm-upplösning i filen)" if len(s.profil) >= 3
            else "vattengång enligt GIS (rät linje, ingen inklinometer)")
    pa = s.profil_analys
    if pa and pa["svackdjup"] is not None and pa["svackdjup"] > 0.01:
        xs = pa["svackpos"]
        if speglad:
            xs = L_prof - xs                                # speglad axel
        j = min(range(len(x)), key=lambda i: abs(x[i] - xs))   # höjd i profilen vid svackans position
        zi = z[j]
        ax.plot([xs, xs], [zi, zi + pa["svackdjup"]], color="#d03b3b", lw=1.5)
        nara_ande = abs(xs - x1) < 0.15 * (x1 - x0) or abs(xs - x0) < 0.15 * (x1 - x0)
        ax.annotate(f"svacka {pa['svackdjup'] * 100:.0f} cm", (xs, zi + (pa["svackdjup"] if nara_ande else 0)),
                    xytext=(0, 10 if nara_ande else -14), textcoords="offset points",   # ovanför nära brunnsnamnen
                    ha="center", fontsize=7.5, color="#d03b3b")
    if pa and pa["osaker"]:
        ax.text(0.01, 0.03, "OBS: brunnshöjder saknas eller är orimliga i filen – profilen kan inte kontrolleras"
                if s.profil_start_z is None or s.profil_slut_z is None
                else "OBS: inklinometerprofilen avviker från brunnshöjderna – osäker", transform=ax.transAxes,
                ha="left", va="bottom", fontsize=7, color="#d03b3b")   # nere till vänster: tomt när höga änden är vänster
    if mark:
        ax.plot([xm for xm, _ in mark], [zm for _, zm in mark], color="#8c6d46", lw=1.4,
                label="markyta (GIS)")
        ax.fill_between([xm for xm, _ in mark], [zm for _, zm in mark], zmitt - zvidd,
                        color="#8c6d46", alpha=0.06, lw=0)
        tk = s.tackning
        if tk:
            xt = L_prof - tk["pos_min"] if speglad else tk["pos_min"]
            zm = min(mark, key=lambda p: abs(p[0] - xt))[1]
            ax.annotate(dk(f"täckning {tk['min']:.2f} m"), (xt, zm), xytext=(0, 6), textcoords="offset points",
                        ha="center", va="bottom", fontsize=7.5, color="#8c6d46")
    ax.plot([x0, x1], [zj[0], zj[-1]], "o", ms=5, color="#4d4d4d")
    ax.annotate(dk(f"{vanster}  {z0:.2f}"), (x0, z0), xytext=(4, 6), textcoords="offset points", fontsize=7.5, fontweight="bold")
    ax.annotate(dk(f"{hoger}  {z1:.2f}"), (x1, z1), xytext=(-4, -12), textcoords="offset points", ha="right",
                fontsize=7.5, fontweight="bold")
    lut = s.lutning_promille
    titel = "Inklinometerprofil"
    if lut is not None:
        titel += dk(f"   ·   lutning {abs(lut):.1f} ‰")
    titel += f"   ·   höjdskala 1:{hojdskala}   ·   längdskala 1:{langdskala}"
    h = s.hojdanpassning
    if h and h["offset"] is not None:
        ax.text(0.01, 0.12 if pa and pa["osaker"] else 0.03, dk("höjdläge: " + h["status"] + (f" ({h['offset']:+.2f} m)" if h["offset"] else "")),
                transform=ax.transAxes, ha="left", va="bottom", fontsize=7, color="#52514e")
    ax.set_title(titel, fontsize=8.5, loc="left", fontweight="bold")
    ax.set_xlabel(f"position (m), 0 = {vanster}", fontsize=7.5)
    ax.set_ylabel("höjd (m)", fontsize=7.5)
    _dk_axlar(ax)
    ax.tick_params(labelsize=7)
    ax.grid(color="#e1e0d9", lw=0.6)
    for sp in ("top", "right"):
        ax.spines[sp].set_visible(False)
    ax.legend(fontsize=7, frameon=False, loc="upper right")
    _spara_rapportbild(fig, path)
    plt.close(fig)
    return True


TABELL_MIN_RADER = 3     # minsta antal datarader på var sida om en sidbrytning i observationstabellen


def _Tabell(*args, **kw):
    """Tabell som inte lämnar en ensam rad på en sida: delningen flyttas uppåt tills båda delarna har
    minst TABELL_MIN_RADER datarader, och en kort tabell flyttas hel till nästa sida."""
    from reportlab.platypus import Table

    class Tabell(Table):
        def split(self, availWidth, availHeight):
            delar = Table.split(self, availWidth, availHeight)
            rubrik = self.repeatRows
            data = len(self._cellvalues) - rubrik
            if len(delar) != 2 or data <= 2 * TABELL_MIN_RADER:
                # för kort att dela snyggt: hel på nästa sida
                return [] if len(delar) == 2 and data > 0 else delar
            hojd = availHeight
            for _ in range(TABELL_MIN_RADER + 1):
                forst = len(delar[0]._cellvalues) - rubrik
                kvar = len(delar[1]._cellvalues) - rubrik
                if forst >= TABELL_MIN_RADER and kvar >= TABELL_MIN_RADER:
                    return delar
                if forst < TABELL_MIN_RADER:
                    return []
                hojd -= self._rowHeights[rubrik + forst - 1]
                delar = Table.split(self, availWidth, hojd)
                if len(delar) != 2:
                    return []
            return delar

    return Tabell(*args, **kw)


def _RubrikTabell(rubrik, tabell):
    """Rubrik + tabell: rubriken följer alltid med tabellens första rader. Får inte rubriken och minst
    TABELL_MIN_RADER rader plats på sidan flyttas båda till nästa sida; en lång tabell börjar ändå
    på samma sida som rubriken (reportlabs KeepTogether flyttar hela tabellen till nästa sida)."""
    from reportlab.platypus import KeepTogether

    class RubrikTabell(KeepTogether):
        def split(self, aW, aH):
            rub, tab = self._content
            _, hh = rub.wrap(aW, aH)
            kvar = aH - hh - rub.getSpaceBefore() - rub.getSpaceAfter()
            _, th = tab.wrap(aW, kvar)
            if th <= kvar:
                return [rub, tab]
            delar = tab.split(aW, kvar)
            ram = getattr(self, "_frame", None)
            if delar or (ram is not None and getattr(ram, "_atTop", False)):
                return [rub] + (delar or [tab])
            return [self.FrameBreak(), rub, tab]      # inte plats för rubrik + några rader: nästa sida

    return RubrikTabell([rubrik, tabell])


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
    st_foto = ParagraphStyle("cf", parent=st_cell, fontSize=7, leading=9)    # långa filnamn på en rad
    st_ankare = ParagraphStyle("a", fontName=normal, fontSize=1, leading=1)

    bredd = A4[0] - 30 * mm

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

    # --- rubrik (prioritetsklass visas inte i rapporten – bara i Excel och kartunderlaget) ---
    el.append(P(f"Sträcka {s.nr}: {s.startbrunn} → {s.slutbrunn}", st_h1))
    el.append(Spacer(1, 4))

    # --- infotabell ---
    pa = s.profil_analys
    lut = s.lutning_promille
    dim = s.dimension + (f"/{s.dimension2}" if s.dimension2 else "") + " mm"
    info = [
        ("Område", s.omrade, "Datum", f"{s.datum} {s.klockslag}".strip() + (f"  ·  {s.vader}" if s.vader.strip() else "")),
        ("Startbrunn (uppströms)", s.startbrunn, "Slutbrunn (nedströms)", s.slutbrunn),
        ("Kamera från", f"{s.fran_brunn} (position 0 m)", "Riktning", s.riktning),
        ("Inspekterad längd", f"{s.langd:.2f} m", "Ledningstyp", s.ledningstyp.capitalize()),
        ("Material", s.material + (f" (foder: {s.foder}, {s.fodermaterial})" if s.foder.strip() else ""),
         "Dimension / form", f"{dim}, {s.form.lower()}"),
        ("Antal anslutningar", str(s.antal_anslutningar), "Antal skador", str(len(s.skador()))),
        ("Konstruktionsindex", f"{s.index('K'):.1f} p/100 m (maxgrad {s.maxgrad('K') or '–'})",
         "Avbruten inspektion", "Ja" if s.avbruten else "Nej"),
        ("Driftindex", f"{s.index('D'):.1f} p/100 m (maxgrad {s.maxgrad('D') or '–'})",
         "Totalindex", f"{s.index():.1f} p/100 m"),
        ("Svacka (djup / längd)", f"{pa['svackdjup'] * 100:.0f} cm / {pa['svacklangd']:.1f} m"
         if pa and pa["svackdjup"] is not None else ("okänd (profil osäker)" if pa else "–"),
         "Lutning", (f"{lut:.1f} ‰" + (" (ur inklinometern)" if s.lutning_ur_inklinometer else "") if lut is not None
                     else ("okänd (höjdfel i filen)" if s.hojdfel else "–"))
         + ((f"  ·  bakfall {pa['bakfall']:.1f} m" if pa and pa["bakfall"] > 0.5 else "")
                                                                    + ("  ·  profil osäker" if pa and pa["osaker"] else ""))),
        ("Videofil", (s.videofil or "–") + (f" + {s.videofil_b}" if s.videofil_b else ""), "TV3-fil", s.fil),
    ]
    h, tk = s.hojdanpassning, s.tackning
    if h:
        info.append(("Höjdläge", h["status"] + (f" ({h['offset']:+.2f} m)" if h["offset"] else ""),
                     "Täckning (min / max)",
                     (f"{tk['min']:.2f} / {tk['max']:.2f} m" + (f"  ·  {s.hojdflagga}" if s.hojdflagga else ""))
                     if tk else "–"))
    if s.hojdfel and not h:
        info.append(("Höjdfel", s.hojdfel[0].upper() + s.hojdfel[1:] + " – de värdena används inte", "", ""))
    if s.littera_rattat:
        info.append(("Littera rättat", s.littera_rattat, "", ""))
    if s.filmstatus:
        info.append(("Filmstatus", s.filmstatus[0].upper() + s.filmstatus[1:], "", ""))
    if s.tidigare:
        info.append(("Tidigare inspektion", tidigare_text(s), "", ""))
    rader = [[P(a, st_fet), P(dk(b)), P(c, st_fet), P(dk(d))] for a, b, c, d in info]
    kw = 38 * mm
    t = Table(rader, colWidths=[kw, bredd / 2 - kw, kw, bredd / 2 - kw])
    stil = [
        ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#d9d9d9")),
        ("BACKGROUND", (0, 0), (0, -1), colors.HexColor("#f2f4f8")),
        ("BACKGROUND", (2, 0), (2, -1), colors.HexColor("#f2f4f8")),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("TOPPADDING", (0, 0), (-1, -1), 2.5), ("BOTTOMPADDING", (0, 0), (-1, -1), 2.5),
    ]
    for i, (a, b, c, d) in enumerate(info):
        if not c and not d:           # textrader (Filmstatus, Tidigare inspektion) får hela radbredden
            stil += [("SPAN", (1, i), (3, i)), ("BACKGROUND", (2, i), (2, i), colors.white)]
    t.setStyle(TableStyle(stil))
    el.append(t)

    # --- schema ---
    schema = os.path.join(tmp, f"schema_{s.nr}.png")
    sb, sh = rita_schema(s, schema)
    el.append(Paragraph("Översikt", st_h2))
    el.append(Image(schema, width=bredd, height=bredd * sh / sb))

    # --- profil ---
    profil = os.path.join(tmp, f"profil_{s.nr}.png")
    if rita_profil(s, profil, bredd / mm):
        el.append(KeepTogether([Paragraph("Profil", st_h2),
                                Image(profil, width=bredd, height=bredd * PROFIL_FIG[1] / PROFIL_FIG[0])]))
    else:
        el.append(Paragraph("Profil", st_h2))
        el.append(P(dk(s.hojdfel[0].upper() + s.hojdfel[1:] + " – profilen ritas inte.") if s.hojdfel
                    else "Ingen inklinometerprofil finns registrerad för sträckan.", st_liten))

    # --- observationstabell ---
    huvud = [P(h, st_vit) for h in ("Pos. m", "Tid", "Kod", "Observation", "Kl.", "Nivå", "Foto", "Grad", "Poäng")]
    rader = [huvud]
    stil = [
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1f3864")),
        ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#d9d9d9")),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("TOPPADDING", (0, 0), (-1, -1), 2), ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
        ("LEFTPADDING", (0, 0), (-1, -1), 3), ("RIGHTPADDING", (0, 0), (-1, -1), 3),
        ("ALIGN", (7, 1), (8, -1), "CENTER"),
    ]
    obs = [o for o in s.observationer if o.kod or o.infokod]

    def esc(t) -> str:
        return str(t).replace("&", "&amp;").replace("<", "&lt;")

    def ankare(namn: str) -> str:
        return "foto_" + re.sub(r"[^A-Za-z0-9]", "_", namn)

    # Bilder som finns med i rapporten (första förekomsten av varje filnamn bär ankaret)
    fotoankare: dict[str, str] = {}
    for o in s.observationer:
        for n, pth in o.bilder:
            if pth and n not in fotoankare:
                fotoankare[n] = ankare(n)

    for i, o in enumerate(obs, 1):
        klocka = o.klocka_fran + (f"–{o.klocka_till}" if o.klocka_till and o.klocka_till != o.klocka_fran else "")
        # klickbart bildnamn → hoppar till fotografiet längre bak i rapporten
        foto = "<br/>".join(f'<a href="#{fotoankare[n]}" color="#1f3864"><u>{esc(n)}</u></a>' if n in fotoankare else esc(n)
                            for n, _ in o.bilder)
        rader.append([P(dk(f"{o.lage:.2f}"), st_cell), P(o.tid, st_cell), P(o.kod or o.infokod, st_cell),
                      P(o.beskrivning(), st_cell), P(klocka, st_cell),
                      P(f"{o.vattenniva}%" if o.vattenniva and o.vattenniva != "0" else "", st_cell),
                      Paragraph(foto, st_foto), P(o.grad or "", st_cell), P(dk(f"{o.poang:.1f}".rstrip("0").rstrip(".")) if o.poang else "", st_cell)])
        if o.grad and o.ar_skada:
            stil.append(("BACKGROUND", (0, i), (-1, i), colors.HexColor(GRAD_FARG_LJUS.get(o.grad, "#ffffff"))))
    if len(rader) == 1:
        rader.append([P("Inga observationer registrerade", st_cell)] + [P("", st_cell)] * 8)
    # bredder efter uppmätt textbredd (DejaVu 7,5 pt) + 2 × 3 pt cellmarginal: "110,04", "00:00:11",
    # "KAM", "06–09", "Nivå", "0706202112412_A.JPG" (ett filnamn per rad, 7 pt), "Grad", "Poäng"
    cw = [13 * mm, 15 * mm, 10 * mm, None, 12 * mm, 10 * mm, 32 * mm, 10 * mm, 12 * mm]
    cw[3] = bredd - sum(w for w in cw if w)
    t = _Tabell(rader, colWidths=cw, repeatRows=1)
    t.setStyle(TableStyle(stil))
    el.append(_RubrikTabell(Paragraph("Observationer", st_h2), t))   # rubriken följer med tabellens första rader
    el.append(Spacer(1, 4))

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
                pth_liten = forminska_foto(pth, tmp)
                with PILImage.open(pth_liten) as im:
                    w, h = im.size
                bh = bw * h / w
                if bh > 95 * mm:
                    bh, bw_ = 95 * mm, 95 * mm * w / h
                else:
                    bw_ = bw
                img = Image(pth_liten, width=bw_, height=bh)
            except Exception:
                img = P(f"[kunde inte läsa {n}]", st_liten)
            txt = Paragraph(f"<b>{esc(n)}</b> · {dk(f'{o.lage:.2f}')} m · {esc(o.tid)}<br/>{esc(o.beskrivning())}", st_liten)
            # ankare ovanför bilden så att länken från tabellen landar med bilden i vy;
            # bara första förekomsten av ett filnamn (samma bild kan höra till flera observationer)
            ank = fotoankare.pop(n, None)
            celler.append([Paragraph(f'<a name="{ank}"/>' if ank else "", st_ankare), img, txt])
        # två per rad
        rader = []
        for i in range(0, len(celler), 2):
            par = celler[i:i + 2]
            rader.append([Table([[c[0]], [c[1]], [c[2]]], colWidths=[bw]) for c in par] + ([""] if len(par) == 1 else []))
        if rader:
            t = Table(rader, colWidths=[bredd / 2, bredd / 2])
            t.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP"), ("LEFTPADDING", (0, 0), (-1, -1), 0),
                                   ("BOTTOMPADDING", (0, 0), (-1, -1), 6)]))
            el.append(t)

    doc.build(el, onFirstPage=sidfot, onLaterPages=sidfot)


def _saker_filnamn(t: str) -> str:
    return re.sub(r"[^A-Za-z0-9ÅÄÖåäö._-]+", "_", t).strip("_")


def rapport_prefix(s: Stracka) -> str:
    """Början på rapportens filnamn, oberoende av klass och brunnar: '<fil>_<nr>_'."""
    return _saker_filnamn(f"{os.path.splitext(s.fil)[0]}_{s.nr:03d}") + "_"


def rapport_filnamn(s: Stracka) -> str:
    return rapport_prefix(s) + _saker_filnamn(f"{s.klass}_{s.startbrunn}-{s.slutbrunn}") + ".pdf"


def rensa_gamla_rapporter(strackor: list[Stracka], katalog: str) -> int:
    """Tar bort PDF:er i katalogen som hör till en sträcka men har ett annat namn än det
    aktuella (t.ex. gammal klassbokstav efter ändrade parametrar, eller rättat littera), samt
    protokoll för delfilmer som inte längre redovisas separat."""
    if not os.path.isdir(katalog):
        return 0
    aktuella = {rapport_filnamn(s) for s in visade(strackor)}
    prefix = {rapport_prefix(s) for s in strackor}
    n = 0
    for namn in os.listdir(katalog):
        if namn.lower().endswith(".pdf") and namn not in aktuella and any(namn.startswith(p) for p in prefix):
            os.remove(os.path.join(katalog, namn))
            n += 1
    return n


_FOTO_CACHE: dict[str, str] = {}


def forminska_foto(pth: str, tmp: str) -> str:
    """Kopia av ett foto för inbäddning i PDF: JPEG, längsta sidan högst FOTO_MAX_PX, och vriden
    enligt EXIF-orienteringen (reportlab läser inte EXIF, så ett foto taget på högkant hamnade
    annars liggande). Originalet returneras när det redan är en liten RGB-JPEG utan rotation, eller
    när något går fel. Kopior cachas per process (samma bild kan höra till flera observationer)."""
    nyckel = os.path.normcase(os.path.abspath(pth))
    if nyckel in _FOTO_CACHE and os.path.isfile(_FOTO_CACHE[nyckel]):
        return _FOTO_CACHE[nyckel]
    try:
        from PIL import Image as PILImage, ImageOps
        with PILImage.open(pth) as im:
            w, h = im.size
            try:
                orientering = im.getexif().get(0x0112, 1)
            except Exception:
                orientering = 1
            liten = not FOTO_MAX_PX or max(w, h) <= FOTO_MAX_PX
            if liten and orientering in (None, 1) and im.mode == "RGB" and (im.format or "").upper() == "JPEG":
                _FOTO_CACHE[nyckel] = pth
                return pth
            if FOTO_MAX_PX and orientering in (None, 1):
                # snabb nedskalad JPEG-avkodning (1/2, 1/4 …); målet måste ha bildens proportioner,
                # annars väljer PIL skala 1 (full-HD mot 960×960 gav ingen nedskalning)
                im.draft("RGB", (FOTO_MAX_PX * w // max(w, h), FOTO_MAX_PX * h // max(w, h)))
            im = ImageOps.exif_transpose(im) or im            # vrid enligt EXIF och ta bort taggen
            if im.mode in ("RGBA", "LA", "PA") or "transparency" in im.info:
                rgba = im.convert("RGBA")                     # genomskinligt → vitt, inte svart
                bg = PILImage.new("RGB", rgba.size, "white")
                bg.paste(rgba, mask=rgba.split()[3])
                im = bg
            if im.mode != "RGB":
                im = im.convert("RGB")                        # gråskala, CMYK, palett → RGB
            if FOTO_MAX_PX and max(im.size) > FOTO_MAX_PX:
                im.thumbnail((FOTO_MAX_PX, FOTO_MAX_PX))
            # entydigt namn per originalfil – ett löpnummer kolliderade när en inaktuell cachepost
            # (raderad tempmapp från en tidigare sträcka) gjordes om och nästa foto fick samma namn
            ut = os.path.join(tmp, "foto_" + hashlib.sha1(nyckel.encode("utf-8")).hexdigest()[:16] + ".jpg")
            im.save(ut, "JPEG", quality=FOTO_JPEG_KVALITET, optimize=True, exif=b"")
        _FOTO_CACHE[nyckel] = ut
        return ut
    except Exception:
        return pth


def _snabb_reportlab() -> None:
    """Stänger av ASCII85-kodningen av bilddata i reportlab. Den körs i ren Python när
    C-tillägget saknas och tog 90 % av rapporttiden; binär inbäddning ger dessutom mindre filer."""
    try:
        from reportlab import rl_config
        rl_config.useA85 = 0
    except Exception:
        pass


def _rapport_jobb(s: "Stracka", sokvag: str) -> tuple[int, str, str]:
    """En rapport i en arbetsprocess. Returnerar (nr, fil, felmeddelande eller '')."""
    import tempfile
    _snabb_reportlab()
    try:
        with tempfile.TemporaryDirectory() as tmp:
            skriv_rapport(s, sokvag, tmp)
        return s.nr, s.fil, ""
    except Exception as e:  # noqa: BLE001 – felet rapporteras i huvudprocessen
        return s.nr, s.fil, str(e)


def skriv_rapporter(strackor: list[Stracka], katalog: str, urval: str,
                    behall: bool = False, processer: int | None = None) -> tuple[int, int]:
    """Skriver en PDF per sträcka i <katalog>. urval: alla | AB | A.
    behall=True hoppar över sträckor vars PDF redan finns (länken sätts ändå).
    processer: antal parallella processer (None = RAPPORT_PROCESSER / antal kärnor − 1, 1 = seriellt).
    Returnerar (antal skrivna, antal befintliga som behölls)."""
    import tempfile
    try:
        import reportlab  # noqa: F401
    except ImportError:
        print("  reportlab saknas – inga PDF-rapporter skapas (pip install reportlab)")
        return 0, 0
    _snabb_reportlab()
    os.makedirs(katalog, exist_ok=True)
    gamla = rensa_gamla_rapporter(strackor, katalog)
    if gamla:
        print(f"  {gamla} inaktuella rapporter borttagna (annan klass eller littera än nu)")
    valda = [s for s in visade(strackor) if urval == "alla" or s.klass in urval]
    n = behallna = 0
    att_skriva: list[tuple[Stracka, str]] = []
    for s in valda:
        namn = rapport_filnamn(s)
        sokvag = os.path.join(katalog, namn)
        if behall and os.path.isfile(sokvag):
            s.rapport_fil = os.path.join(os.path.basename(katalog), namn)
            behallna += 1
        else:
            att_skriva.append((s, sokvag))
    if not att_skriva:
        return 0, behallna

    if processer is None:
        processer = RAPPORT_PROCESSER
    if processer is None:
        processer = max(1, (os.cpu_count() or 2) - 1)
    processer = min(int(processer), len(att_skriva))
    klara = 0

    def framsteg():
        if klara % 25 == 0 or klara == len(att_skriva):
            print(f"  rapporter: {klara}/{len(att_skriva)}", end="\r", flush=True)

    if processer > 1:
        try:
            from concurrent.futures import ProcessPoolExecutor, as_completed
            with ProcessPoolExecutor(max_workers=processer) as pool:
                try:
                    jobb = {pool.submit(_rapport_jobb, s, sokvag): (s, sokvag) for s, sokvag in att_skriva}
                    for f in as_completed(jobb):
                        s, sokvag = jobb[f]
                        nr, fil, fel = f.result()
                        if fel:
                            print(f"\n  FEL rapport sträcka {nr} ({fil}): {fel}")
                        else:
                            s.rapport_fil = os.path.join(os.path.basename(katalog), os.path.basename(sokvag))
                            n += 1
                        klara += 1
                        framsteg()
                except KeyboardInterrupt:
                    # annars väntar with-blocket in alla redan inskickade jobb innan avbrottet syns
                    pool.shutdown(wait=False, cancel_futures=True)
                    raise
            print()
            return n, behallna
        except Exception as e:  # noqa: BLE001 – t.ex. miljö utan processpool: kör seriellt
            print(f"\n  parallell körning misslyckades ({e}) – skriver resterande rapporter en i taget")
            att_skriva = [(s, sokvag) for s, sokvag in att_skriva if not s.rapport_fil]
            klara = 0

    with tempfile.TemporaryDirectory() as tmp:
        for s, sokvag in att_skriva:
            try:
                skriv_rapport(s, sokvag, tmp)
                s.rapport_fil = os.path.join(os.path.basename(katalog), os.path.basename(sokvag))
                n += 1
            except Exception as e:  # noqa: BLE001
                print(f"\n  FEL rapport sträcka {s.nr} ({s.fil}): {e}")
            klara += 1
            framsteg()
    print()
    return n, behallna

# ----------------------------------------------------------------------------
# main
# ----------------------------------------------------------------------------

Listpost = tuple[str, list[str], list[str]]      # (tv3-sökväg, filmkataloger, bildkataloger)
Globala = dict[str, list[str]]                    # {"media": [...], "bild": [...], "littera": [...]}
_BILD_NYCKLAR = ("bild", "bilder", "foto", "foton")
_MEDIA_NYCKLAR = ("media", "film", "filmer", "video", "videor")
_MARK_NYCKLAR = ("markprofil", "mark")
_UPP_NYCKLAR = ("uppstroms", "uppströms", "uppstrom")
_GIS_NYCKLAR = ("gis", "gisdata")
_MANUELL_NYCKLAR = ("manuell", "manuellt", "bedomning", "bedömning")
_KOSTNAD_NYCKLAR = ("kostnader", "kostnad", "priser")


def las_listfil(path: str) -> tuple[list[Listpost], Globala]:
    """Läser en listfil. Returnerar ([(tv3-sökväg, [filmkataloger], [bildkataloger]), ...],
    {"media": [filmkataloger för alla filer], "bild": [bildkataloger för alla filer],
     "littera": [CSV-filer med ersättningslittera]}).

    Format (en post per rad, tomma rader och #-kommentarer ignoreras):
        media: D:\\Inspektioner\\Filmer          filmkatalog för alla TV3-filer i listan
        bild: D:\\Inspektioner\\Foton            bildkatalog för alla filer (valfritt – annars
                                                 söks bilderna i filmkatalogerna)
        littera: brunnslittera.csv               ersättningslittera för felmärkta brunnar
        markprofil: Karta\\markprofil.json        från ArcMap-verktyget Markprofil
        uppstroms: Karta\\uppstroms.csv           från ArcMap-verktyget Uppströms (batchläge)
        gis: Karta\\gisdata.json                  från ArcMap-verktyget Exportera GIS-data
        manuell: forra_korningen\\prioritering.xlsx  manuella bedömningar, kommentarer, Lagning (m)
        kostnader: kostnader.csv                 kostnadsposter för åtgärdspaketet
        DUF 701.TV3                              TV3-fil; media söks i filens egen katalog
        DUF 702.TV3 ; D:\\Filmer\\DUF702          TV3-fil med egen filmkatalog (fler kan
                                                 anges, separerade med ;)
        DUF 703.TV3 ; D:\\Film ; bild: E:\\Foton   egen filmkatalog och egen bildkatalog
        C:\\Inspektioner\\2022\\                  katalog: alla TV3-filer i den
    Relativa sökvägar tolkas relativt listfilens katalog."""
    bas = os.path.dirname(os.path.abspath(path))

    def abs_(p: str) -> str:
        p = p.strip().strip('"').strip("'")
        return p if os.path.isabs(p) else os.path.normpath(os.path.join(bas, p))

    nyckel_re = re.compile(r"^(%s)\s*[:=]\s*(.*)$" % "|".join(_BILD_NYCKLAR + _MEDIA_NYCKLAR + _MARK_NYCKLAR
                                                             + _UPP_NYCKLAR + _MANUELL_NYCKLAR + _KOSTNAD_NYCKLAR
                                                             + _GIS_NYCKLAR
                                                             + ("littera", "brunnslittera", "brunnar")),
                           re.IGNORECASE)
    poster: list[Listpost] = []
    globala: Globala = {"media": [], "bild": [], "littera": [], "markprofil": [], "uppstroms": [],
                        "manuell": [], "kostnader": [], "gis": []}
    for rad in las_text(path).splitlines():
        rad = re.split(r"\s+#", rad, 1)[0].strip()      # kommentar efter blanksteg + # tillåts
        if not rad or rad.startswith("#"):
            continue
        m = nyckel_re.match(rad)
        if m:
            nyckel = m.group(1).lower()
            slag = ("bild" if nyckel in _BILD_NYCKLAR else "media" if nyckel in _MEDIA_NYCKLAR
                    else "markprofil" if nyckel in _MARK_NYCKLAR
                    else "uppstroms" if nyckel in _UPP_NYCKLAR
                    else "gis" if nyckel in _GIS_NYCKLAR
                    else "manuell" if nyckel in _MANUELL_NYCKLAR
                    else "kostnader" if nyckel in _KOSTNAD_NYCKLAR else "littera")
            globala[slag] += [abs_(d) for d in m.group(2).split(";") if d.strip()]
            continue
        delar = rad.split(";")
        tv3 = abs_(delar[0])
        media: list[str] = []
        bild: list[str] = []
        for d in delar[1:]:
            d = d.strip()
            if not d:
                continue
            m = re.match(r"^(%s)\s*[:=]\s*(.+)$" % "|".join(_BILD_NYCKLAR + _MEDIA_NYCKLAR), d, re.IGNORECASE)
            if m and m.group(1).lower() in _BILD_NYCKLAR:
                bild.append(abs_(m.group(2)))
            elif m:
                media.append(abs_(m.group(2)))
            else:
                media.append(abs_(d))
        poster.append((tv3, media, bild))
    return poster, globala


def hitta_tv3_filer(argument: list[str], listfiler: list[str]) -> tuple[list[Listpost], Globala]:
    """Löser upp argument (filer, kataloger, jokertecken, .txt-listor) till
    [(tv3-fil, [filmkataloger], [bildkataloger]), ...] samt de globala film-/bildkatalogerna
    och littera-CSV:erna från listfilerna."""
    import glob
    kandidater: list[Listpost] = []
    globala: Globala = {"media": [], "bild": [], "littera": [], "markprofil": [], "uppstroms": [],
                        "manuell": [], "kostnader": [], "gis": []}

    def lagg_till_lista(lf: str) -> None:
        p, g = las_listfil(lf)
        kandidater.extend(p)
        for slag in globala:
            globala[slag] += g[slag]

    for lf in listfiler:
        lagg_till_lista(lf)
    for arg in argument:
        if arg.lower().endswith(".txt"):          # listfil även utan -l
            lagg_till_lista(arg)
        elif any(ch in arg for ch in "*?["):
            kandidater += [(f, [], []) for f in sorted(glob.glob(arg))]
        else:
            kandidater.append((arg, [], []))

    filer: list[Listpost] = []
    for k, media, bild in kandidater:
        if os.path.isdir(k):
            for rot, _, namn in os.walk(k):
                filer += [(os.path.join(rot, n), media, bild) for n in sorted(namn) if n.lower().endswith(".tv3")]
        else:
            filer.append((k, media, bild))

    # ta bort dubbletter, behåll ordning (katalogerna slås ihop)
    index: dict[str, int] = {}
    unika: list[Listpost] = []
    for f, media, bild in filer:
        key = os.path.abspath(f)
        if key in index:
            for lista, nya in ((unika[index[key]][1], media), (unika[index[key]][2], bild)):
                lista.extend(m for m in nya if m not in lista)
        else:
            index[key] = len(unika)
            unika.append((f, list(media), list(bild)))
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
    ap.add_argument("--behall-rapporter", action="store_true", default=BEHALL_RAPPORTER,
                    help="hoppa över PDF-rapporter som redan finns i utdatakatalogen (snabbare omkörning)")
    ap.add_argument("--processer", type=int, default=None, metavar="N",
                    help="antal parallella processer för PDF-rapporterna (standard: antal kärnor − 1; 1 = en i taget)")
    ap.add_argument("--karta", choices=["ja", "nej"], default="ja" if SKRIV_KARTUNDERLAG else "nej",
                    help=f"skriv {KARTUNDERLAG_FIL} för ArcMap-skriptet i arcmap/ (standard: ja)")
    ap.add_argument("--diagram", action="store_true", default=SPARA_DIAGRAM,
                    help="spara diagrammen som PNG i <utdata>/diagram; de bäddas alltid in i Excel")
    ap.add_argument("--littera", action="append", default=[], metavar="FIL.CSV",
                    help="CSV med ersättningslittera för felmärkta brunnar (fel;rätt per rad); "
                         "kan även anges i listfilen som 'littera: FIL.CSV'")
    ap.add_argument("--media", action="append", default=[], metavar="KATALOG",
                    help="extra katalog att söka videofiler i, och bilder om ingen bildkatalog angetts "
                         "(kan anges flera gånger); TV3-filens egen katalog söks alltid")
    ap.add_argument("--markprofil", action="append", default=[], metavar="FIL.JSON",
                    help="markprofil.json från ArcMap-verktyget Markprofil (markhöjder och GIS-vattengång); "
                         "kan även anges i listfilen som 'markprofil: FIL'")
    ap.add_argument("--uppstroms", action="append", default=[], metavar="FIL.CSV",
                    help="CSV från ArcMap-verktyget Uppströms (batchläge): serviser och längd uppströms "
                         "per sträcka; kan även anges i listfilen som 'uppstroms: FIL'")
    ap.add_argument("--gis", action="append", default=[], metavar="FIL.JSON",
                    help="gisdata.json från ArcMap-verktyget Exportera GIS-data (brunnar och ledningar): "
                         "kontroll av vattengång, littera, riktning, material och dimension samt djup; "
                         "kan även anges i listfilen som 'gis: FIL'")
    ap.add_argument("--manuell", action="append", default=[], metavar="FIL.XLSX",
                    help="tidigare prioritering.xlsx med ifyllda Manuell bedömning, Kommentar och Lagning (m); "
                         "kan även anges i listfilen som 'manuell: FIL'")
    ap.add_argument("--kostnader", metavar="FIL.CSV", default=None,
                    help=f"kostnadsfil för åtgärdspaketet (standard: {KOSTNADSFIL} bredvid listfilen eller skriptet); "
                         "kan även anges i listfilen som 'kostnader: FIL'")
    ap.add_argument("--pptx", action="store_true",
                    help="skriv presentation.pptx i utdatakatalogen ur mall.pptx (kräver python-pptx; "
                         "kan också göras efteråt med tv3_pptx.py UTDATA)")
    ap.add_argument("--etapper", choices=["ja", "nej"], default="ja",
                    help="åtgärdspaket: etappindelning och kostnad (fliken Etapper); nej = hoppa över")
    ap.add_argument("--bilder", action="append", default=[], metavar="KATALOG",
                    help="katalog att söka bilder i (kan anges flera gånger); anges ingen söks "
                         "bilderna i videokatalogerna; kan även anges i listfilen som 'bild: KATALOG'")
    a = ap.parse_args(argv)

    if not a.filer and not a.lista:
        ap.error("ange minst en TV3-fil, katalog eller listfil (-l filer.txt)")

    filer, globala = hitta_tv3_filer(a.filer, a.lista)
    globala_media, globala_bild, littera_filer = globala["media"], globala["bild"], globala["littera"]
    if not filer:
        sys.exit("Inga TV3-filer hittades.")
    print(f"{len(filer)} fil(er) att analysera\n")

    strackor: list[Stracka] = []
    fel: list[str] = []

    littera: list[dict] = []
    for lf in littera_filer + a.littera:
        if not os.path.isfile(lf):
            fel.append(f"{lf}: litterafilen finns inte")
            print(f"  VARNING litterafil saknas: {lf}")
            continue
        littera += las_littera(lf)
    littera.sort(key=lambda r: -r["villkor"])
    if littera:
        print(f"  {len(littera)} ersättningslittera lästa"
              + (f" ({sum(1 for r in littera if r['villkor'])} begränsade till fil/sträcka/motbrunn)"
                 if any(r["villkor"] for r in littera) else ""))
    for p, media, bild in filer:
        for m in media + bild:
            if not os.path.isdir(m):
                fel.append(f"{p}: mediakatalogen finns inte: {m}")
                print(f"  VARNING mediakatalog saknas: {m}")
        if not os.path.exists(p):
            fel.append(f"{p}: filen finns inte")
            print(f"  SAKNAS  {p}")
            continue
        try:
            st = las_tv3(p, littera)
        except Exception as e:  # trasig fil ska inte stoppa hela körningen
            fel.append(f"{p}: {e}")
            print(f"  FEL     {p}: {e}")
            continue
        for st_ in st:
            st_.media_kataloger = list(media)
            st_.bild_kataloger = list(bild)
        print(f"  OK      {os.path.basename(p)}: {len(st)} sträckor, {sum(s.langd for s in st):.0f} m, "
              f"{sum(len(s.skador()) for s in st)} skadeobservationer"
              + (f", {sum(1 for s in st if s.littera_rattat)} sträckor med rättat littera"
                 if any(s.littera_rattat for s in st) else ""))
        strackor += st
    if not strackor:
        sys.exit("Inga sträckor hittades.")
    # Två TV3-filer med samma namn i olika mappar: skilj dem åt med mappnamnet, annars blandas
    # de ihop i Excel, rapportnamnen och kartunderlaget (som matchar på fil + sträcknummer).
    per_namn: dict[str, set[str]] = defaultdict(set)
    for s in strackor:
        per_namn[s.fil].add(s.tv3_sokvag)
    for s in strackor:
        if len(per_namn[s.fil]) > 1:
            s.fil = os.path.basename(os.path.dirname(s.tv3_sokvag)) + "/" + s.fil
            for o in s.observationer:
                o.fil = s.fil
    if fel:
        print(f"\n{len(fel)} varning(ar) – se {os.path.join(a.utdata, 'fel.txt')}")

    for m in globala_media + a.media + globala_bild + a.bilder:
        if not os.path.isdir(m):
            print(f"  VARNING mediakatalog saknas: {m}")
    vh, vt, bh, bt = koppla_media(strackor, globala_media + a.media, globala_bild + a.bilder)
    markfiler = []
    for mf in globala["markprofil"] + a.markprofil:
        if not os.path.isfile(mf):
            fel.append(f"{mf}: markprofilen finns inte")
            print(f"  VARNING markprofil saknas: {mf}")
            continue
        markfiler.append(las_markprofil(mf))
    if markfiler:
        n_mark, n_vg = koppla_markprofil(strackor, markfiler)
        status = Counter(s.hojdanpassning["status"] for s in strackor if s.hojdanpassning)
        print(f"Markprofil: {n_mark} sträckor med markhöjder, {n_vg} med GIS-vattengång"
              + (f" ({markfiler[0]['hojdsystem']})" if markfiler[0]["hojdsystem"] else ""))
        for st_, n in status.most_common():
            print(f"  {n:>4} {st_}")
        flaggade = [s for s in strackor if s.hojdflagga]
        if flaggade:
            print(f"  {len(flaggade)} sträckor med höjdflagga, t.ex. "
                  + ", ".join(f"{s.id} ({s.hojdflagga})" for s in flaggade[:3]))
    gisfiler, gisstat = [], None
    for gf in globala["gis"] + a.gis:
        if not os.path.isfile(gf):
            fel.append(f"{gf}: GIS-datafilen finns inte")
            print(f"  VARNING GIS-data saknas: {gf}")
            continue
        gisfiler.append(las_gis(gf))
    if gisfiler:
        g = koppla_gis(strackor, gisfiler)
        print(f"GIS-data: {g['ledningar']} av {len(strackor)} sträckor har ledning i GIS, "
              f"{g['vg']} fick vattengång ur GIS, {g['flaggade']} flaggade, "
              f"{len(set(o['littera'] for o in g['okanda']))} okända brunnar")
        typer = Counter(fl.split(":")[0] for s in strackor for fl in s.gisflagga.split("; ") if fl)
        for typ, n in typer.most_common():
            print(f"  {n:>4} {typ}")
        lf = os.path.join(a.utdata, "littera_forslag.csv")
        if not g["okanda"] and os.path.isfile(lf):
            os.remove(lf)                                  # inaktuella förslag från en tidigare körning
        if g["okanda"]:
            os.makedirs(a.utdata, exist_ok=True)
            skriv_litteraforslag(g["okanda"], lf)
            print(f"  okända brunnar med förslag ur GIS skrivna till {lf}")
            for o in g["okanda"][:6]:
                print(f"    {o['littera']} (nr {o['nr']}) → "
                      + (", ".join(f"{lit} ({txt})" for lit, txt in o["forslag"]) or "inget förslag"))
    uppposter = []
    for uf in globala["uppstroms"] + a.uppstroms:
        if not os.path.isfile(uf):
            fel.append(f"{uf}: uppströmsfilen finns inte")
            print(f"  VARNING uppströmsfil saknas: {uf}")
            continue
        uppposter += las_uppstroms(uf)
    if uppposter:
        n_upp = koppla_uppstroms(strackor, uppposter)
        kallor = Counter(s.serviser_kalla for s in strackor if s.serviser_uppstroms is not None)
        print(f"Uppströms: {n_upp} av {len(strackor)} sträckor kopplade"
              + (" (" + ", ".join(f"{k}: {n}" for k, n in kallor.most_common()) + ")" if kallor else ""))
    manposter, framschakta = [], []
    for mf in globala["manuell"] + a.manuell:
        if not os.path.isfile(mf):
            fel.append(f"{mf}: filen med manuella bedömningar finns inte")
            print(f"  VARNING fil med manuella bedömningar saknas: {mf}")
            continue
        try:
            po, fr = las_manuella(mf)
        except Exception as e:                 # noqa: BLE001 – trasig/öppen Excelfil ska inte stoppa körningen
            fel.append(f"{mf}: kunde inte läsas ({e})")
            print(f"  VARNING kunde inte läsa manuella bedömningar: {mf} ({e})")
            continue
        manposter += po
        framschakta += fr
    if manposter:
        n_man = koppla_manuella(strackor, manposter)
        print(f"Manuella bedömningar: {n_man} sträckor ({sum(1 for s in strackor if s.manuell_bedomning)} med "
              f"bedömning, {sum(1 for s in strackor if s.lagning_m)} med lagning)")
    strackor = hantera_omfilmningar(strackor)
    ersatta = [s for s in strackor if s.filmstatus.startswith("ersatt")]
    sammanslagna = [s for s in strackor if s.sammanslagen_av]
    if ersatta or sammanslagna:
        print(f"Omfilmningar: {len(ersatta)} filmer ersatta av nyare/hel film, "
              f"{len(sammanslagna)} brunnspar sammanslagna av två delfilmer")
        for s in sammanslagna[:5]:
            print(f"  nr {s.nr} {s.id}: {s.filmstatus}")
    tidigare = [s for s in strackor if s.tidigare]
    if tidigare:
        print(f"Ny inspektion över filgränser: {_tidigare_text_summa(strackor)}")
        for s in sorted(tidigare, key=lambda s: s.tidigare["utveckling"] != "förvärrad")[:5]:
            print(f"  {os.path.basename(s.fil)} nr {s.nr} {s.id}: {tidigare_text(s)}")
    rakn = aktiva(strackor)        # det som räknas i statistik, karta, åtgärdspaket och inspektionsgrad
    gisstat = inspektionsgrad(rakn, gisfiler) if gisfiler else None
    if gisstat:
        print("Inspektionsgrad: " + _inspektionsgrad_text(gisstat))
        if gisstat["utan_gis"][0]:
            print(f"  {gisstat['utan_gis'][0]} filmade sträckor ({gisstat['utan_gis'][1]:.0f} m) saknar ledning i GIS")
    etapper = None
    if a.etapper == "ja":
        kostnadsfil = a.kostnader or (globala["kostnader"][0] if globala["kostnader"] else None)
        if kostnadsfil is None:
            for kandidat in ([os.path.join(os.path.dirname(os.path.abspath(lf)), KOSTNADSFIL) for lf in a.lista]
                             + [os.path.join(os.path.dirname(os.path.abspath(__file__)), KOSTNADSFIL)]):
                if os.path.isfile(kandidat):
                    kostnadsfil = kandidat
                    break
        kostnader = []
        if kostnadsfil and os.path.isfile(kostnadsfil):
            kostnader = las_kostnader(kostnadsfil)
        else:
            print(f"  VARNING kostnadsfil saknas ({kostnadsfil or KOSTNADSFIL}) – etapper utan kostnad")
        etapper = planera_atgarder(rakn, kostnader, framschakta)
        n_str = sum(1 for s in rakn if s.metod == "strumpa")
        n_sch = sum(1 for s in rakn if s.metod == "schakt")
        kr = f"{sum(e['kostnad']['summa'] for e in etapper):,.0f}".replace(",", " ")
        print(f"Åtgärdspaket: {len(etapper)} etapper, {n_str} sträckor strumpa, {n_sch} schakt, "
              f"{sum(len(e['framschaktade']) for e in etapper)} brunnar att schakta fram, {kr} kr"
              + (f" (kostnader: {os.path.basename(kostnadsfil)})" if kostnader else " (ingen kostnadsfil)"))
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
    diagram = rita_diagram(rakn, diagramkatalog, a.topp)
    if a.rapporter != "inga":
        print(f"\nSkriver PDF-rapporter ({a.rapporter}) ...")
        n, behallna = skriv_rapporter(strackor, os.path.join(a.utdata, "rapporter"), a.rapporter,
                                      behall=a.behall_rapporter, processer=a.processer)
        print(f"  {n} rapporter skrivna till {os.path.join(a.utdata, 'rapporter')}"
              + (f", {behallna} befintliga behållna" if behallna else ""))
    excel_fil = os.path.join(a.utdata, "prioritering.xlsx")
    try:
        skriv_excel(strackor, excel_fil, diagram, a.topp, etapper, gisstat)
    except PermissionError:
        sys.exit(f"\nKan inte skriva {excel_fil} – filen är troligen öppen i Excel. "
                 "Stäng den och kör igen.")
    if a.karta == "ja":
        kartfil = os.path.join(a.utdata, KARTUNDERLAG_FIL)
        n_poster = skriv_kartunderlag(rakn, kartfil, gisstat, etapper)
        print(f"\n{n_poster} sträckor skrivna till {kartfil} (underlag för ArcMap)")
    if a.pptx:
        try:
            import tv3_pptx
            pptx_fil = tv3_pptx.bygg_presentation(rakn, etapper, diagram, a.utdata, topp=a.topp,
                                                  ta=sys.modules[__name__])
            print(f"Presentation skriven till {pptx_fil}")
        except ImportError as e:
            print(f"  VARNING PowerPoint hoppas över: {e} (pip install python-pptx)")
        except Exception as e:  # noqa: BLE001 – mall eller innehåll får inte stoppa övriga utdata
            fel.append(f"PowerPoint: {e}")
            print(f"  VARNING PowerPoint hoppas över: {e}")
    if not a.diagram:
        shutil.rmtree(diagramkatalog, ignore_errors=True)
    felfil = os.path.join(a.utdata, "fel.txt")
    if fel:
        with open(felfil, "w", encoding="utf-8") as f:
            f.write("\n".join(fel) + "\n")
    elif os.path.exists(felfil):
        os.remove(felfil)                 # ingen gammal fellista ska ligga kvar från förra körningen

    klasser = Counter(s.klass for s in rakn)
    print("\nPrioritetsklasser: " + ", ".join(f"{KLASS_TEXT[k]}: {klasser.get(k, 0)}" for k in "ABCDE"))
    print(f"\nTopp {a.topp}:")
    for i, s in enumerate(sorterade_strackor(rakn)[:a.topp], 1):
        print(f"{i:>3}. [{s.klass}] {s.id:<28} {s.material:<7}{s.dimension:>4} {s.langd:6.1f} m  "
              f"k-index {s.index('K'):6.1f}  {s.sammanfattning_skador()}")
    print(f"\nResultat skrivet till: {os.path.abspath(a.utdata)}")


if __name__ == "__main__":
    import multiprocessing
    multiprocessing.freeze_support()    # fryst exe (PyInstaller) på Windows: annars spawnas skriptet rekursivt
    main()
