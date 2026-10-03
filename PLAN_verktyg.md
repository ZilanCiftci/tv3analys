# Implementeringsplan: fyra nya verktyg

Status: plan, oktober 2026. Beslut och svar på frågorna förs in under respektive avsnitt.

Ordning (varje steg bygger på det föregående):

1. **Svackor och bakfall som karta** – liten utökning av JSON och Skapa ledningslager.
2. **Uppströmsanalys i nätet** – ger en återanvändbar nätverksmodul i ArcMap och antal anslutna uppströms per sträcka.
3. **Åtgärdspaket och kostnad** – fristående i `tv3_analys.py`, kan ta antal anslutna uppströms från steg 2.
4. **PowerPoint ur analysen** – sist, plockar ihop allt det andra.

Sammanlagt cirka sju arbetspass. Steg 1 och 2 testas i ArcMap (Citrix), steg 3 och 4 körs lokalt.

---

## 1. Svackor och bakfall som karta

**Syfte.** Ett punktlager med svackor och ett linjelager med bakfall, för att se var sediment och stopp kommer att återkomma, och för att kunna visa det i webb-GIS.

**Var.** `tv3_analys.py` skriver mer i `kartunderlag.json`; `arcmap/skapa_ledningslager.py` får två valfria utdatalager. Ingen ny dialog, två nya valfria utdataparametrar i Skapa ledningslager.

**Indata.** Svackpositionen finns redan i `profil_analys` (kamerans position). Bakfall räknas i dag bara som total längd; `profil_analys` utökas med en lista av bakfallssträckor (från, till i meter, flödesriktning).

**JSON per sträcka.** `svackpos_m` räknat från startbrunnen, `bakfall_segment` (lista med [från, till]), samt det som redan finns: `svackdjup_cm`, `svacklangd_m`, svackdjup/diameter, `profil_osaker`.

**ArcMap.** Vägen brunn→brunn finns som punktlista orienterad startbrunn→slutbrunn. Positionen skalas med kartlängd/filmlängd (samma regel som markprofilen; ingen skalning vid avbruten inspektion) och punkten läggs ut längs linjen. Bakfallssegment klipps ur samma linje.

- Lager `svackor` (punkt): FRAN_BRUNN, TILL_BRUNN, SVACKA_CM, SVACKLANGD, ANDEL_DIAM, MASK_BED, OSAKER, RAPPORT.
- Lager `bakfall` (linje): FRAN_BRUNN, TILL_BRUNN, LANGD_M, LUTNING, MASK_BED, RAPPORT.

**Symbologi.** Storlek efter svackdjup i tre steg, röd när djupet överstiger halva diametern. Sparas som `.lyr` en gång i ArcMap, som för bedömda ledningar.

**Test.** Låtsas-arcpy med en sträcka med känd svacka (punkten på rätt meter). DUF 701: de två kända svackorna ska hamna på KRB68925→KNBL62753 och KTB1003603→KRB1000744.

**Omfattning.** Ett arbetspass.

---

## 2. Uppströmsanalys i nätet

**Syfte.** Peka på en sträcka eller brunn och få allt som ligger uppströms: ledningar, brunnar, serviser, total längd. I batchläge: antal anslutna uppströms per bedömd sträcka, som underlag för konsekvens och etappindelning.

**Var.** Ny modul `arcmap/natverk.py` som bryter ut grafbygget ur `Natverk` och generaliserar det (alla brunnar är noder, inte bara JSON-filens). Nytt verktyg **Uppströms** i toolboxen. Skapa ledningslager byter till den gemensamma modulen i ett senare steg.

**Indata.** Ledningslager, brunnslager, valfritt servislager (eller typkolumn om serviser ligger i samma lager), och startpunkt: vald ledning i kartan eller ett brunnslittera. Valfritt område eller sökavstånd från startpunkten för att begränsa grafen (hela SVOA-nätet är stort).

**Flödesriktning** (kärnfrågan). Ledningslagret har ett **riktningsattribut** (F6): dialogen får ett fältval och en tolkning av värdena (vilket värde som betyder "med ritad riktning" respektive "mot"). Som reserv när attributet saknas på en ledning används vattengångsfälten (högre → lägre), annars ritad riktning; sådana ledningar loggas.

**Algoritm.** Rikta varje kant uppströms→nedströms, gå bakåt från startnoden med bredd-först, samla alla kanter. Varje nod besöks en gång (ringmatning i dagvattennät ger cykler). Ledningstyper som inte får passeras (tryckledningar, pumpstationer) anges i dialogen och stoppar sökningen. Serviser räknas som ledningar i servislagret som ansluter till en uppströmskant inom toleransen. Saknas servislager används fältet Anslutningar ur kartunderlaget som skattning för de inspekterade delarna.

**Utdata.** Markering i kartan, lager `uppstroms_<littera>` med fälten AVSTAND_M (till startpunkten), NIVA (steg i trädet), SERVIS (ja/nej), samt sammanfattning i loggen och CSV: antal ledningar, total längd, antal serviser, antal brunnar, längsta gren.

**Batchläge.** Samma verktyg med bedömda-lagret som indata skriver per sträcka `ANT_SERV_U` (serviser uppströms) och `L_UPPSTR` (längd uppströms) till lagret och till en CSV som `tv3_analys.py` kan läsa in (`uppstroms: FIL` i listfilen) för konsekvens och etappordning.

**Risker.** Riktningsdata av blandad kvalitet; pumpstationer och tryckledningar; prestanda på stora nät (begränsa med område/avstånd).

**Test.** Syntetiskt nät med förgrening, slinga och pumpstation i låtsas-arcpy; verifiera antal och längd.

**Omfattning.** Två arbetspass, varav hälften riktningslogiken.

---

## 3. Åtgärdspaket och kostnad

**Syfte.** Förvandla prioriteringslistan till en plan: sammanhängande sträckor blir etapper med metod, mängder och kostnad.

**Var.** `tv3_analys.py` (ny KONFIG-sektion), ny kostnadsfil `kostnader.csv` i repot som användaren äger. Etapp, metod och kostnad följer med till kartan via `kartunderlag.json` och Skapa ledningslager (fält `ETAPP`, `METOD`, `KOSTNAD`).

### 3.1 Brunnstyp avgör om strumpan kan installeras

Strumpa installeras från en **nedstigningsbrunn** (NB/NBL). Från en **tillsynsbrunn** (TB) går det inte. Brunnstypen tas ur:

1. TVDAT-infokoden vid sträckans ändar (NB, TB, RB) om den finns,
2. annars litterats prefix (…NB/…NBL = nedstigning, …TB = tillsyn, …RB = rensbrunn),
3. annars okänd → flaggas.

Regler per sträcka:

| Brunnar i ändarna | Åtgärd | Kostnadsposter |
|---|---|---|
| Minst en nedstigningsbrunn | Strumpa från den | strumpa kr/m |
| Bara tillsyns-/rensbrunnar | Schakta fram en brunn, strumpa från öppet schakt, ny brunn | framschaktning kr/st + strumpa kr/m + ny brunn kr/st |
| Okänd typ | Som nedstigningsbrunn, men flaggas "brunnstyp okänd" | – |

Rensbrunn (RB) behandlas som tillsynsbrunn: strumpa kan inte installeras därifrån (F1).

### 3.2 Metodval per sträcka

| Villkor | Metod |
|---|---|
| Relinad redan | ingen åtgärd (flaggas om klass A/B – fodret är skadat) |
| Grad 4 på RBR eller DEF, eller manuell bedömning "schakt" | schakt (hela sträckan) |
| Klass A eller B i övrigt | strumpa (+ punktlagningar, se 3.3) |
| Klass C, D | ingen åtgärd (kan tas med i etapp för sammanhang, se 3.4) |

Manuell bedömning i Excel kan alltid styra metoden: värdena A–E styr klassen, texten "schakt", "strumpa" eller "ingen" styr metoden.

### 3.3 Lagning före strumpning (manuell bedömning)

Lokala skador som är så svåra att strumpan inte kan installeras över dem lagas med punktschakt först. Det avgörs **manuellt** (F2): fliken Prioritering får kolumnen **Lagning (m)** där användaren anger hur många meter som behöver lagas med schakt före strumpning. Kostnaden blir meter × kr/m för posten `lagning` i kostnadsfilen. Skriptet föreslår inget automatiskt, men i kolumnen Skador (kod+grad) syns de observationer som brukar kräva lagning (RBR/DEF grad 3, FOG grad 4, YTS grad 4), så att genomgången går fort.

**Ingen automatisk brytpunkt** mot schakt (F3). Metoden byts bara av manuell bedömning ("schakt"). Kalkylen visar däremot alltid båda alternativen per sträcka, strumpa inklusive lagning och framschaktning respektive schakt av hela sträckan, så att skillnaden syns i Excel.

### 3.4 Etappindelning

Sträckorna bildar en graf via brunnsparen (samma som flerinspekterade par). Sammanhängande sträckor med åtgärd och samma metod blir en etapp. En C- eller D-sträcka mellan två åtgärdssträckor tas med om den är kortare än `ETAPP_OVERBRYGGA_M` (standard 60 m, beslutat F5) – en sammanhängande infodring är oftast billigare än två etableringar; sådana sträckor markeras "medtagen för sammanhang". Etapperna numreras efter högsta konstruktionsindex i etappen, alternativt efter konsekvens (antal anslutna uppströms från steg 2) om `ETAPP_ORDNING = "konsekvens"`.

### 3.5 Kostnadsfil `kostnader.csv`

Semikolonseparerad, decimalkomma, användaren äger den. Schablonvärden läggs in som start, tydligt märkta "schablon" (F4). Bara metoderna strumpa och schakt (F7); fler metoder kan läggas till som nya poster senare.

```
post;dimension_fran;dimension_till;enhet;kr;kommentar
strumpa;0;200;m;…;
strumpa;201;300;m;…;
strumpa;301;400;m;…;
strumpa;401;600;m;…;
schakt;0;300;m;…;
schakt;301;600;m;…;
hatt;;;st;…;anslutning som öppnas och tätas med hatt
lagning;;;m;…;punktschakt före strumpning, meter enligt manuell bedömning
framschaktning;;;st;…;schakta fram brunn när bara tillsynsbrunnar finns
ny_brunn;;;st;…;ny nedstigningsbrunn
etablering;;;etapp;…;fast kostnad per etapp
```

Kostnad per sträcka (strumpa) = strumpa kr/m × längd (dimensionsintervall) + hattar × antal anslutningar + lagning kr/m × meter enligt manuell bedömning + framschaktning + ny brunn (när båda ändarna är tillsyns-/rensbrunnar). Kostnad per sträcka (schakt) = schakt kr/m × längd. Båda redovisas; vald metod avgör vad som summeras i etappen. Kostnad per etapp = summan + etablering. Dimensionsintervallet väljs på sträckans dimension; saknas intervall flaggas sträckan.

### 3.6 Utdata

- Ny flik **Etapper** i Excel: etapp, sträckor, brunnar från/till, längd, metod, dimensioner, antal anslutningar (hattar), antal punktlagningar, framschaktning/ny brunn, kostnad per post och totalt, högsta klass, medtagna för sammanhang.
- Fliken Prioritering: kolumnerna Etapp, Metod, Kostnad (synliga), Lagning (m) (synlig, fylls i manuellt), Kostnad strumpa, Kostnad schakt, Brunnstyp start/slut (dolda).
- Fliken Sammanfattning: total kostnad per metod och klass.
- `kartunderlag.json`: `etapp`, `metod`, `kostnad`, `lagning_m`, `brunnstyp_start`, `brunnstyp_slut`; Skapa ledningslager: fälten `ETAPP`, `METOD`, `KOSTNAD`.
- PDF-protokollet: ingen ändring (kostnad hör till Excel och kartan, som klassen).

**Test.** DUF 701: kontroll att brunnstyper tolkas rätt (KNBL/KRB/KTB/SRB…), att etapperna blir sammanhängande, och att en sträcka med bara TB får framschaktning. Syntetisk kostnadsfil.

**Omfattning.** Två arbetspass.

---

## 4. PowerPoint ur analysen

**Syfte.** En färdig presentation per uppdrag ur utdatamappen, i användarens mall.

**Var.** Nytt skript `tv3_pptx.py` som läser en färdig utdatamapp (kan köras om efter justeringar i Excel). Startas även med `--pptx` från `tv3_analys.py`. Beroende: `python-pptx`.

**Mall.** `mall.pptx` (utkast i repots rot, justeras av användaren) med layouterna TITEL, AVDELARE, RUBRIK, RUBRIK_TEXT, RUBRIK_BILD, RUBRIK_TABELL, STRACKA och AVSLUT och namngivna platshållare (title, undertitel, meta, body, bild, tabell, oversikt, foto1–foto3, fakta). Skriptet fyller platshållarna via namnen och ritar inget eget utöver nyckeltalsrutor på layouten RUBRIK. Layoutnamnen och platshållarnamnen måste behållas när mallen anpassas.

**Bilder.** Klassfördelning, topplista, skador per kod och klass per material finns som PNG. Protokollens översikt och profil sparas undan i en bildmapp när `--pptx` är på. Fotografier ur mediamapparna. Kartbild: nytt litet verktyg **Exportera kartbild** i toolboxen sparar aktuell vy som PNG i utdatamappen; saknas bilden hoppas sidan över.

**Innehåll** (lista i KONFIG så att sidor kan strykas): titelsida (uppdrag, område, period, entreprenör); nyckeltal; klassfördelning; topplista; skador per kod; material; etapper och kostnad (steg 3); en sida per sträcka i topp tio med översikt, tre foton och nyckeluppgifter; driftåtgärder; svackor (steg 1); metodsida med parametrarna ur Sammanfattning.

**Test.** Generera för DUF 701 och rendera sidorna till bilder för kontroll.

**Omfattning.** Två arbetspass, plus mallen från användaren.

---

## Beslut (besvarat 2026-10-03)

- **F1** Rensbrunn behandlas som tillsynsbrunn – strumpa kan inte installeras därifrån.
- **F2** Lagning före strumpning är en manuell bedömning: kolumnen "Lagning (m)" i Excel, prissatt kr/m.
- **F3** Ingen automatisk brytpunkt mot schakt; bara manuell bedömning byter metod. Båda kostnaderna visas.
- **F4** Schablonpriser tills vidare, märkta schablon i `kostnader.csv`.
- **F5** C/D-sträckor kortare än 60 m mellan åtgärdssträckor tas med i etappen.
- **F6** Flödesriktning ur ett riktningsattribut i ledningslagret (fält och värdetolkning väljs i dialogen); vattengång och ritad riktning som reserv.
- **F7** Bara strumpa och schakt.

Kvar att få av användaren: namnet på riktningsattributet och dess värden (inför steg 2). `mall.pptx` finns som utkast i repot och justeras av användaren (behåll layout- och platshållarnamnen).
