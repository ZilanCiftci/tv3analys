# Plan: ledningsinventering inom ett område – från ArcMap till kostnadsjämförelse strumpa/omläggning

Status: oktober 2026. Steg 1 (Exportera område) är byggt (PR #119), steg 2 (Ledningsinventering →
`ledningsdata.json`, verktyg och modul `arcmap/moduler/inventering.py`) är byggt men oprövat i riktig ArcMap.
Besluten F10–F17 står längst ned; F10 (vattenlagrens namn) och prislistan i F13 väntar på användaren.

## Målet, som jag förstår det

1. Ta ut **allt VA-nät inom ett eller flera områden** (polygoner i ett eget lager i ArcMap) till lokala
   lager, med all attributinformation, så att allt vidare arbete går utan kontakt med databasen.
2. Exportera **ledningarna med all information till Excel** och komplettera med det som inte står på
   ledningen men behövs för en kostnadsjämförelse: antal serviser och deras längd per sträcka,
   medeldjup per sträcka, parallella ledningar (samma schakt) och lutning.
3. Räkna **strumpa kontra omläggning** per sträcka (och per stråk/etapp).
4. **Profiler** per ledningsstråk (inte för serviser) och **planritningar** ur ArcMap.

Nästan allt bygger på moduler som redan finns i `arcmap/moduler/`:

| Behov | Finns redan | Återanvänds i |
|---|---|---|
| Brunn→brunn-nät, serviser kopplade till ledning | `natverk.Graf`, `serviser_vid`, `Segmentindex` | steg 2 |
| Vattengång interpolerad längs ledningen, skarvar utan brunn | `gisexport.exportera`, `markprofil.Ledningar` | steg 2 |
| Markhöjd ur punktlager eller raster (höjdmodell) | `projektering.Markhojd` | steg 2, 4 |
| Profilritning A3 med exakta skalor, brunnar, täckning, nivåtabell | `projektering.rita_profil` | steg 4 |
| Kartor per sträcka/etapp i SVOA:s utskriftsmallar | `kartexport` | steg 4 |
| Kostnadsposter, etappindelning, Excel-byggare | `tv3_analys.py` (kostnader.csv, `planera_atgarder`, `tabell`) | steg 3 |

Ordning (varje steg ger något användbart för sig):

1. **Exportera område** – lokala lager. En kväll.
2. **Ledningsinventering** – `ledningsdata.json` med serviser, djup, parallellitet, lutning. Två pass.
3. **Excel och kostnad** – `ledning_analys.py`: inventeringsbok + strumpa/omläggning. Två pass.
4. **Profiler och planritningar** per stråk. Två pass.
5. **Koppling till TV-inspektionen** och scenarier per etapp. Ett pass.

Sammanlagt sju–åtta arbetspass. Steg 1, 2 och 4 körs i ArcMap (Citrix), steg 3 lokalt i Python 3.

---

## 1. Exportera område (nytt ArcMap-verktyg)

**Syfte.** Alla VA-objekt som berör områdespolygonerna kopieras till en ny filgeodatabas, ett
featureklass per källlager, med alla fält. Därefter arbetar alla andra verktyg mot den lokala gdb:n.

**Dialog.** Områdeslager (polygon; bara valda objekt om några är valda, annars alla), namnfält i
områdeslagret (valfritt), lista över lager att ta med (flervals-rullista med kartans lager, förvalt:
alla lager under gruppen *Avlopp* och *Vatten* – ledningar, serviser, alla brunnslager,
serviskopplingspunkter, ledningsändar, ventiler/brandposter om de finns), utdatamapp och gdb-namn,
buffert (m, standard 0 – se nedan).

**Metod.** `SelectLayerByLocation INTERSECT` mot polygonerna och `CopyFeatures` – **hela objekt, inte
klippta**, så att längder, vattengångar och brunnspar förblir riktiga. Varje objekt får fälten
`OMRADE` (polygonens namn) och `KORSAR_GRANS` (ja när objektet sticker ut ur området). Med buffert > 0
tas även ledningar strax utanför med, så att nätet hänger ihop vid gränsen (serviser och parallella
ledningar till en gränsledning). Lagren läggs in i kartan under en grupp *Lokalt: <gdb-namn>* med
källagrens symbologi (`UpdateLayer`), definitionsfrågor kopieras inte (de är redan tillämpade i urvalet).
En `export_info.json` i mappen anger datum, områden, lager, antal objekt och koordinatsystem – den läses
av de senare verktygen så att rätt lager hittas automatiskt.

**Tid.** SelectLayerByLocation mot SDE är snabb; CopyFeatures av några tusen objekt tar sekunder.
Verktyget loggar per lager (som Lista fält) så att man ser var det står.

## 2. Ledningsinventering (nytt ArcMap-verktyg → `ledningsdata.json` + CSV)

**Syfte.** En post per ledningssträcka (brunn→brunn) med allt som behövs för Excel och kalkyl, ur den
lokala gdb:n. Utökar `gisexport.exportera` snarare än ersätter den – samma brunn/ledningsmodell, så
att `tv3_analys.py` kan läsa filen som `gis:` också.

**Per ledning (utöver det gisexport redan ger: från/till brunn, längd, vg upp/ned, dimension, material,
typ, år, renoveringsår, skarvar):**

- **Alla originalfält** (EntityID, PipeType, SewerFunction, Commentary …) under `attribut`.
- **Lutning** beräknad ur vattengångarna och längden (‰, positiv = fall i ritad riktning) samt fältet
  `Slope` som jämförelse; flaggor *bakfall*, *flack* (< 2 ‰), *vg saknas*, *vg orimlig*.
- **Markhöjd var 1 m längs linjen** (`projektering.Markhojd`: raster/höjdmodell om det finns i kartan,
  annars punktlager; F11) och **djup** = mark − vattengång interpolerad i samma punkt: medel, min, max,
  samt djup vid brunnarna ur locknivå − vg. Djupklass (0–2, 2–3, 3–4, > 4 m; F13) för schaktkalkylen.
  Saknas markhöjd används locknivåerna i ändbrunnarna som rät linje (markerat `djup_kalla`).
- **Serviser**: `natverk.serviser_vid` kopplar varje servis vars ände ligger inom toleransen (1 m) till
  ledningen. Per ledning: antal, summerad längd, lista {EntityID, längd, typ, dimension, djup vid
  anslutningen om vg finns}. Serviser som inte hittar någon huvudledning listas i datakvalitet.
- **Parallella ledningar**: linjen samplas var 1 m och närmaste andra ledning (annan typ eller samma)
  söks inom `PARALLELL_AVSTAND` (3,0 m c–c; F12) med `Segmentindex`. Per par: andel av längden med den
  andra ledningen inom avståndet, medelavstånd, typ. Andel ≥ 60 % = *samschakt* (grupp-id
  `SCHAKTGRUPP` så att S + D + V i samma gata blir en grupp). Redovisas både per ledning
  (`parallella: [{ledning, typ, andel, avstand}]`) och som egen lista.
- **Korsningar** med andra ledningar (geometrisk korsning) och höjdskillnad där båda har vg – viktigt vid
  omläggning (en vattenledning 0,3 m över spilledningen styr schaktdjupet).
- **Brunnar**: alla fält, locknivå, bottennivå om ifylld, antal anslutna ledningar, djup (lock − lägsta vg).

**Vatten.** Tryckledningar har ingen vattengång. Djup räknas som `VATTEN_DJUP_CENTRUM_M` = 1,7 m till
rörets centrum (F17) och markeras `djup_kalla` = schablon. Lutning redovisas inte för vatten.

**Dialog.** Lokal gdb (eller lager i kartan), fältval som i Exportera GIS-data (gissas ur SVOA-namnen:
EntityID, LevelFrom/LevelTo, PipeDimension, PipeMaterial, PipeType, ConstructionYear, CoverLevel …),
markhöjdslager, servislager, tolerans, parallellavstånd, utdata. Minnet av senaste val som i övriga verktyg.

**Utdata.** `ledningsdata.json` (ledningar, brunnar, serviser, parallella par, korsningar, datakvalitet)
och `ledningsdata_ledningar.csv`/`_serviser.csv`/`_parallella.csv` (utf-8-sig, decimalkomma, öppnas
direkt i Excel om man bara vill titta).

## 3. Excel och kostnad (`ledning_analys.py`, Python 3, lokalt)

**Syfte.** Inventeringsboken och kostnadsjämförelsen. Egen fil med KONFIG överst som `tv3_analys.py`,
återanvänder dess Excel-hjälpare (`tabell`, enhetsrad, dolda kolumner, klassfärger) och `las_kostnader`.

**Körning.** `python ledning_analys.py ledningsdata.json [-o mapp] [--kostnader kostnader_omlaggning.csv]
[--karta kartunderlag.json]`.

**`ledningsinventering.xlsx`:**

- **Sammanfattning** – längd per område/typ/material/dimension/decennium, antal serviser, medeldjup,
  andel samschakt, summa strumpa/omläggning, datakvalitet.
- **Ledningar** – en rad per sträcka: EntityID, från/till brunn, typ, dimension, material, år, längd,
  vg upp/ned, lutning (beräknad och fältet), bakfall, djup medel/min/max, djupklass, serviser (st, m),
  parallell med (lista, andel), schaktgrupp, korsningar, kostnad strumpa, kostnad omläggning, kvot,
  rekommendation (se nedan), TV-klass om `--karta` angetts, flaggor, alla originalfält dolda sist.
- **Serviser** – en rad per servis med huvudledning, längd, typ, dimension, djup vid anslutningen.
- **Parallella** – en rad per par med andel, medelavstånd, gemensam längd, schaktgrupp.
- **Brunnar** – alla brunnar med lock, botten, djup, antal ledningar.
- **Kostnad per stråk/schaktgrupp** – summerat: strumpa alla, lägg om alla, och blandat per sträcka.
- **Datakvalitet** – vg saknas, bakfall, orimlig lutning, servis utan huvudledning, dimension saknas,
  ledning utan brunn i någon ände, år saknas.

**Kostnadsmodell (ny fil `kostnader_omlaggning.csv`, samma format som `kostnader.csv` plus djupkolumner):**

```
post;dimension_fran;dimension_till;djup_fran;djup_till;enhet;kr;kommentar
schakt;0;300;0;2;m;9000;schablon – grund schakt, rör ≤ 300
schakt;0;300;2;3;m;12000;schablon
schakt;0;300;3;4;m;16000;schablon – spont
schakt;0;300;4;99;m;22000;schablon
servis_omkoppling;;;;;st;35000;schablon – ny anslutning + omkoppling vid tomtgräns
servis_ledning;;;;;m;4000;schablon – ny servisledning i schakt
brunn_ny;;;;;st;60000;schablon – ny brunn i omlagd sträcka
aterstallning_gata;;;;;m;3500;schablon – asfalt/kantsten (F13: ytskikt okänt i GIS)
samschakt_andel;;;;;faktor;0,4;andra ledningen i samma schakt kostar 40 % av egen schakt
```

F13: användaren har en egen prisdatabas i Excel som räknar med 1, 2 eller 3 ledningar i samma grav.
Kostnadsfilen anpassas till den när priserna kommer – troligen kr/m per (antal ledningar i graven,
dimension, djupklass) i stället för en samschaktfaktor; schaktgruppen i steg 2 ger antalet ledningar.

Strumpa räknas som i åtgärdspaketet (`kostnader.csv`: kr/m per dimension + hatt per servis +
etablering). Omläggning = schakt kr/m efter dimension och **medeldjup** (eller djupklass per
delsträcka om djupet varierar mycket) + återställning + serviser (omkoppling per st + ny servisledning
per m av servisens längd) + ny brunn per brunn som byts; sträckor i samma schaktgrupp får
`samschakt_andel` på schakt och återställning för alla utom den dyraste ledningen. Allt är schabloner
som du byter ut; filen är det enda som behöver ändras.

**Rekommendation per sträcka** (enkel regel, justerbar i KONFIG): strumpa när kvoten omläggning/strumpa
≥ `STRUMPA_KVOT` (1,5) och ledningen är strumpbar (inte RBR/DEF grad 4 om TV-klass finns, dimension
inom strumpintervall); omläggning när ledningen är grund och har få serviser eller när flera ledningar
delar schakt; annars *bedöm*. Det är en sorteringshjälp, inte ett beslut.

**Tillbaka till kartan.** `ledning_analys.py` skriver också `ledningsresultat.json`, och ett litet
ArcMap-verktyg **Uppdatera inventeringslager** fyller fält i den lokala gdb:n (DJUP_MEDEL, DJUPKLASS,
ANT_SERV, SERV_M, SCHAKTGRUPP, KOST_STRUMPA, KOST_OMLAGG, KVOT, REK) så att det går att symbolisera
på kostnadskvot eller djupklass. GeoJSON för webb-GIS som i övriga verktyg.

## 4. Profiler och planritningar per stråk

**Stråk.** Profilen ritas per stråk = kedja av ledningar av samma typ mellan förgreningspunkter,
startbrunn eller valda brunnar. Byggs ur `natverk.Graf` (ritad riktning/vattengång ger ordningen);
`projektering._kedjor` gör redan detta för projekteringslagren och kan användas som den är. Långa stråk
delas i sidor om högst `PROFIL_MAX_M` (300 m) med överlapp.

**Ledningsprofil (nytt verktyg, utbyggnad av Projekteringsprofil).** Indata: lokal gdb (ledningar,
brunnar, markhöjd), urval = valda ledningar i kartan, eller två brunnar (väg söks i grafen), eller *alla
stråk i området*. `projektering.rita_profil` ritar redan vattengång, hjässa, mark, täckning, brunnar som
rektanglar lock→botten och nivåtabell i exakta skalor; tillägg: serviser som små markeringar på
hjässan med littera (position, ingen egen profil), parallella ledningar av andra typer i samma profil
(streckade, som i projekteringsprofilen), korsande ledningar som cirklar med nivå, sträcknamn och
EntityID i tabellen, medeldjup per sträcka under nivåtabellen. Utdata: en PDF per stråk (A3 liggande)
+ `strak.csv` med stråk-id, ledningar, längd, start/slut.

**Planritning per stråk.** `kartexport.exportera` får ett urvalsläge `strak` med det lokala
ledningslagret som bedömda-lager (markering av stråkets ledningar, titel = stråk-id, undertitel = typ,
dimension, längd) i SVOA:s liggande/stående mallar; samlad PDF `kartor_strak.pdf` och fältet KARTA på
ledningarna. Profil-PDF och plan-PDF får samma filnamn med olika suffix så att de ligger parvis.

## 5. Koppling till TV-inspektionen och scenarier

- `--karta kartunderlag.json` i `ledning_analys.py`: TV-klass, konstruktionsindex, grad 4-koder,
  relinad och etapp kopplas per brunnspar (som `koppla_gis` fast omvänt). Rekommendationen använder det
  (RBR/DEF 4 → inte strumpa; klass D/E → ingen åtgärd).
- **Scenarier per etapp/schaktgrupp**: *strumpa allt*, *lägg om allt*, *rekommenderat per sträcka* –
  tre kolumner i fliken Etapper, med skillnaden i kronor och antal serviser som berörs.
- Åtgärdspaketet i `tv3_analys.py` kan sedan läsa omläggningskostnaden från `ledningsresultat.json`
  i stället för "schakt utan kostnad" (F8 i förra planen) – då blir Etapper-fliken komplett.

---

## Verktygen som behöver tas fram

| # | Verktyg | Var | Nytt/utökning |
|---|---|---|---|
| 1 | Exportera område | `.pyt` + `moduler/omrade_export.py` | nytt |
| 2 | Ledningsinventering | `.pyt` + `moduler/inventering.py` (bygger på gisexport, natverk, Markhojd) | nytt |
| 3 | `ledning_analys.py` + `kostnader_omlaggning.csv` | repots rot, Python 3 | nytt |
| 4 | Uppdatera inventeringslager | `.pyt`, liten | nytt |
| 5 | Ledningsprofil | utökning av Projekteringsprofil/`projektering.rita_profil` | utökning |
| 6 | Planritning per stråk | utökning av Exportera kartor (`kartexport`) | utökning |
| 7 | Omläggningskostnad i åtgärdspaketet | `tv3_analys.py` | liten utökning |

Testas som tidigare med låtsas-arcpy i scratchpad och med en syntetisk `ledningsdata.json`; riktig
ArcMap-körning är din del. Handledningen får ett nytt kapitel 9 *Ledningsinventering*.

## Fler saker som kan vara till nytta

- **Förnyelsebehov ur ålder och material** – teknisk livslängd per material (betong 1920-tal, segjärn,
  PVC …) ger en kolumn *restlivslängd* som komplement till TV-klassen; för vatten finns ingen film, så
  där är det ålder/material/läckhistorik som finns.
- **Dimensionskontroll nedströms** – dimensionen ska inte minska i flödesriktningen; en flagga ur
  grafen hittar flaskhalsar och fel i GIS.
- **Brunnsdjup och brunnstyp** per etapp – antal nedstigningsbrunnar som behöver bytas vid omläggning
  och var tillsynsbrunnar behöver schaktas fram vid strumpning (finns redan i åtgärdspaketet).
- **Ytskikt/återställning** – om det finns ett lager med gata/gc-väg/park/tomtmark kan återställningen
  prissättas per ledning (korsning med ytlager, andel per klass). Stor post i verkligheten.
- **Konflikter med andra ledningsslag** – el, fjärrvärme, opto, om ni kan få ut dem som lager; läggs in
  som "korsningar" i steg 2.
- **Serviser per fastighet** – koppla serviskopplingspunkterna till fastighetslagret för att se vilka
  fastigheter som berörs av en omläggning (brev, avstängning).
- **Känslighet** – kör kostnaden med ±30 % på schaktpriset och visa om rekommendationen byter; billigt
  att lägga in när modellen finns.
- **Tidsplan och nuvärde** – strumpa nu + omläggning om 30 år mot omläggning nu, med ränta; en kolumn
  om ni jämför livscykelkostnad, inte bara investering.
- **Webb-GIS** – GeoJSON av inventeringen med kostnader och rekommendation, som de andra lagren.
- **PowerPoint** – sidor ur inventeringen i `tv3_pptx.py` (nyckeltal, karta, topplista på kvot).

## Beslut och frågor (F10–F17)

- **F10 Lager.** Öppen: vattenlagrens namn och fält kommer när användaren sitter vid ArcMap (kör *Lista
  fält* på dem). Exportera område förväljer alla featurelager under grupperna *Avlopp* och *Vatten*, så
  namnen behövs först i steg 2 (fältgissningar för vatten).
- **F11 Markhöjd.** Ja, det finns en höjdmodell/markhöjder i kartan → djup ur markhöjd − vg.
- **F12 Parallellt.** c–c < 3 m på minst 60 % av sträckan = samma schakt.
- **F13 Kostnader.** Djupklasser enligt förslaget. Användaren har en Excel-prisdatabas som tar hänsyn till
  1, 2 eller 3 ledningar i samma grav; priserna klistras in senare och kostnadsfilen anpassas efter den.
- **F14 Sträcka.** Brunn→brunn, eller till ledningsände utan brunn (servis, T-rörsanslutning). Fristående
  från TV3-filmningarna (TV-klassen kopplas bara på när den finns).
- **F15 Profil.** Per stråk, A3 liggande, högst 300 m per sida, alla parallella ledningsslag i samma profil.
- **F16 Planritning.** Det finns A3-mallar, både stående och liggande – de används (A4 som reserv).
- **F17 Vatten.** Djup till rörets centrum = 1,7 m (schablon).
