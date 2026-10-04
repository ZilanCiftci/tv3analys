# tv3_analys – projektminne och överlämning

Detta dokument sammanfattar allt som byggts upp i samarbetet kring analys av TV-inspektioner av
avloppsledningar, så att arbetet kan fortsätta i ett kodprojekt utan att tappa kontext.
Lägg det som `CLAUDE.md` i repots rot (eller läs in det som första prompt).

## 1. Vem, vad och varför

- Användaren (Zilan) är extern projektledare inom svensk kommunal VA-infrastruktur och arbetar
  bl.a. mot Stockholm Vatten och Avfall (SVOA). Kommunikation sker på svenska; kod, kommentarer,
  kolumnrubriker och dokument ska vara på svenska.
- Målet: läsa TV3-filer från filminspektioner av avloppsledningar, hitta de mest kritiska
  ledningarna som bör renoveras och ta fram underlag som går att använda i presentationer och
  som beslutsunderlag (Excel, diagram, PDF-rapporter per sträcka).
- Beslut som togs tidigt: **Python-skript** (inte webbapp – en Vite/TypeScript-plan skrevs men
  lades ned), **viktad poäng per sträcka** som kritikalitetsmodell, och att **flera TV3-filer
  ska kunna analyseras tillsammans** eftersom fler uppdrag kommer.
- Exempelfil: `testdata/DUF 701.TV3` – SVOA, Alvik/Äppelviken, inspekterad maj–juni 2021, 185 sträckor,
  7 385 m, 1 998 observationsrader, entreprenör Foria Miljö (WinCan-exporter). Använd den som
  testfixtur/facit.

## 2. Nuvarande leverans (allt i ett skript)

`tv3_analys.py` (~1 200 rader, Python 3.10+, beroenden: `openpyxl`, `matplotlib`, `reportlab`).
Övriga filer: `filer.txt` (exempel-listfil), `brunnslittera.csv` (exempel på ersättningslittera:
`fel;ratt;fil;nr;motbrunn;kommentar` – felmärkta brunnar i TV3-filen byts ut vid inläsning;
`fil`/`nr`/`motbrunn` är valfria och begränsar raden till en fil, ett sträcknummer eller ett
brunnspar när det felaktiga litterat även är ett riktigt littera på andra sträckor; mest
specifika raden vinner; rubrikraden styr kolumnordningen, tab/;/, som avgränsare; anges med
`littera: FIL` i listfilen eller `--littera`; Excel-kolumn "Littera rättat", rad i PDF och fält
`littera_rattat` i JSON; användarens sex DUF 701-rättningar ligger i filen),
`Anvandarhandledning tv3_analys.docx`,
`Metodbeskrivning prioritering avloppsledningar.docx` (genereras av `make_docs.js` med npm-paketet
`docx`; filnamn hålls ASCII eftersom Windows zip-hantering förvanskar åäö).

Körning:
```
python tv3_analys.py -l filer.txt [-o utdata] [--topp 15] [--rapporter alla|AB|A|inga] [--behall-rapporter] [--diagram] [--media KATALOG] [--bilder KATALOG] [--littera FIL.CSV]
python tv3_analys.py "testdata/DUF 701.TV3"   # enstaka fil, jokertecken eller katalog fungerar också
```

Listfilens format (relativa sökvägar tolkas relativt listfilen, `#` = kommentar):
```
media: D:\Inspektioner\Filmer          # filmmapp för alla filer i listan (kan upprepas; alias film/video)
bild: D:\Inspektioner\Foton            # bildmapp (valfritt, alias foto; sep 2026) – utan den söks
                                       # bilderna i filmmapparna
littera: brunnslittera.csv             # ersättningslittera (kan upprepas)
DUF 701.TV3                            # media söks alltid även i TV3-filens egen mapp (rekursivt)
DUF 702.TV3 ; D:\Filmer\DUF702 ; bild: E:\Foton   # egen film- och bildmapp för just den filen
C:\Inspektioner\2022\                  # katalog: alla .TV3 i den
markprofil: Karta\markprofil.json      # från ArcMap-verktyget Markprofil (sep 2026), alias mark:
```
Internt: `las_listfil` → `(poster, globala)` med poster `(tv3, filmkataloger, bildkataloger)` och
`globala = {"media", "bild", "littera", "markprofil", "uppstroms", "manuell", "kostnader", "gis"}`; `Stracka.media_kataloger`/`bild_kataloger`;
`koppla_media(strackor, media, bild)` söker bilder i bildkataloger + TV3-mappen om några angetts,
annars i filmkatalogerna.

Utdata i `tv3_resultat/` (eller `-o`):
- `prioritering.xlsx` – flikar **Sammanfattning** (nyckeltal, klassfördelning, metodparametrar,
  inbäddade diagram), **Prioritering** (en rad per sträcka, rankad; klassceller färgade;
  hyperlänkar till rapport-PDF (relativ länk) och videofil (absolut `file:///`-länk);
  **dolda kolumner** sep 2026: `DOLDA_KOLUMNER` i KONFIG per flik, varje dold kolumn får
  `hidden=True` + `outlineLevel=1` (kolumngrupp med plustecken, `summaryRight=False`) – synliga i
  Prioritering: Rang, Klass, brunnar, Material, Dim, Längd, Konstruktionsindex, Konstr. maxgrad,
  Anslutningar, Skador, Avbruten, Relinad, Svackdjup, Manuell bedömning, Kommentar, Rapport,
  Videofil – Nr, Driftåtgärd och Höjdflagga döljs på användarens begäran; Observationer döljer
  Fil, Typ, Löpande, Klocka till, Vattennivå). **Enhetsrad** (sep 2026): `tabell()` delar
  rubriker "Namn (enhet)" i rubrikrad 1 och enhetsrad 2 (grå kursiv), data från rad 3, frys A3,
  autofilter från rad 2; returnerar första dataraden (anropen använder den för talformat).
  `DOLDA_KOLUMNER`/`bredder` anges med rubrik utan enhet. "Antal anslutningar" heter "Anslutningar",
  **Observationer** (alla observationer i klartext, länk till bild och video),
  **Kodstatistik**, **Per fil**, **Material** (relinade sträckor redovisas som eget material
  `Relinad` med ursprungsmaterialet i egen kolumn – styrs av `RELINAD_SOM_MATERIAL` i KONFIG;
  samma indelning i diagram 4, medan fliken Prioritering visar ursprungsmaterial + flaggan Relinad).
- `diagram/1_prioritetsklasser.png`, `2_topplista.png`, `3_observationer_per_kod.png`,
  `4_klass_per_material.png` (200 dpi, för PowerPoint) – **bara med `--diagram`**; annars ritas de
  i en temporär katalog och bäddas enbart in i fliken Sammanfattning (KONFIG: `SPARA_DIAGRAM`).
- `rapporter/<tv3>_<nr>_<klass>_<startbrunn>-<slutbrunn>.pdf` – inspektionsprotokoll per sträcka.
  `--behall-rapporter` (KONFIG `BEHALL_RAPPORTER`) hoppar över PDF:er som redan finns – länken sätts
  ändå; kör utan flaggan när layouten ändrats. `rensa_gamla_rapporter` tar alltid bort PDF:er med
  samma `<fil>_<nr>_`-prefix men annat namn (gammal klassbokstav/littera).
- `kartunderlag.json` – en post per sträcka (brunnspar, maskinell bedömning, index, material,
  flaggor) för ArcMap-skriptet. Stängs av med `--karta nej` (KONFIG: `SKRIV_KARTUNDERLAG`).
- `fel.txt` – bara om TV3-filer eller mediamappar saknades (tas bort vid felfri omkörning).
- Kodgranskning sep 2026 rättade: tom Kodstatistik kraschade `tabell()`; diagram 4 delade med noll
  för material utan längd; hårdkodade radnummer i Sammanfattning (nu söks "Prioritetsklass"/
  "Poängmodell"); `fran_brunn` faller tillbaka på Riktning när utgångsbrunnen inte är en av
  sträckans brunnar (felstavning vände annars profilen); litterafilens rubrikrad matchas på
  nyckelord ("Felaktigt littera" → fel) med varning om fel/ratt saknas; Per fil/Material visar
  konstruktionsindex; två TV3-filer med samma namn får `mapp/namn` som `fil`; mediaindexet
  sorterar `os.walk` för determinism.
- **Ingen** `sammanfattning.md` längre (togs bort på användarens begäran).
- **Rapporthastighet** (okt 2026, användaren: "tar väldigt lång tid"): profilering visade att 90 % av
  tiden var reportlabs `_py_asciiBase85Encode` av fotona (C-tillägget rl_accel saknas → ren Python) och
  att fotona bäddades in i full upplösning (778 MB PDF för DUF 701 med 488 full-HD-foton). Rättat:
  `_snabb_reportlab()` sätter `rl_config.useA85 = 0` (binär bilddata, mindre filer); `forminska_foto`
  skalar till `FOTO_MAX_PX` 960 px (JPEG `draft` ger halv avkodning direkt, `FOTO_JPEG_KVALITET` 82,
  cache per process i `_FOTO_CACHE`, original om redan litet/JPEG eller fel); `skriv_rapporter` kör
  parallellt med `ProcessPoolExecutor` (`RAPPORT_PROCESSER`/`--processer`, standard kärnor − 1,
  `_rapport_jobb` per sträcka med egen tempmapp, seriell reserv vid fel). DUF 701: 3 m 35 s → 35 s
  (4 kärnor), 25 rapporter seriellt 266 s → 19 s; rapportmappen 778 → ~130 MB. Andra omgången (användaren ville
  ha mindre filer och mer fart): utjämningen i `rita_profil` var O(n²) (3,9 av 13,4 s för 25 figurer) →
  glidande fönster O(n) (±0,3 m inklusive med 1e-9-epsilon, gamla formeln uteslöt gränspunkter
  slumpmässigt p.g.a. flyttal); `RAPPORT_BILD_DPI` 150 (var 200) och `_spara_rapportbild` →
  palett-PNG 256 färger (`RAPPORT_BILD_PALETT`, 58→20 kB, ~0,1 s/rapport); `FOTO_JPEG_KVALITET` 76 +
  `optimize=True`. 25 figurpar 13,4 → 5,6 s; DUF 701 full körning 35 → 33 s (palett kostar det dpi
  sparar), rapportmappen 125 → 102 MB med brusfoton (verkliga foton ~30 kB/st → ~25 MB). Kvar: text-
  rendering i matplotlib (~0,15 s/figur) och reportlab-bygget (~0,17 s/rapport). Granskning okt 2026
  (rättat): cachefilnamnet i `forminska_foto` är sha1 av originalsökvägen (löpnumret kolliderade när en
  inaktuell post gjordes om → fel foto på fel observation); `draft` får proportionellt mål (960×960 gav
  skala 1 för full-HD); alfa läggs mot vit bakgrund; EXIF-orientering tillämpas (`exif_transpose`, taggen
  tas bort, gråskala/CMYK → RGB); Ctrl-C i poolen avbryter köade jobb (`cancel_futures`); seriell reserv
  skriver bara sträckor utan `rapport_fil`; `freeze_support()` under main-skyddet; progress med `flush`.
  Palett-PNG:n bäddas in som RGB av reportlab – vinsten (~35 %/figur i PDF) är bättre komprimering.
  KONFIG-ändringar gjorda från ett omslagsskript når inte spawn-arbetarna på Windows (modulen
  importeras om) – CLI-körning påverkas inte.

Kartframställning: `arcmap/tv3_verktyg.pyt` är en Python Toolbox (ArcMap 10.x) med verktygen
**Skapa ledningslager** (dialog: JSON-fil, lager, brunnsfält, utdata, valfritt område/.lyr/CSV,
tolerans, max hopp, extra fält) och **Uppdatera bedömning**. Logiken ligger i
`arcmap/skapa_ledningslager.py` (Python 2.7 + arcpy; funktionerna `skapa()` och `uppdatera()`,
modulen laddas om vid varje verktygskörning). Samma fil kan köras med `execfile` i Python-fönstret
med KONFIG-blocket. Verktyget läser `kartunderlag.json`, letar upp varje brunnspar i
brunnslagret (`A Nedstign och övriga brunnar`, fält `EntityID`) och klipper ut ledningen
mellan brunnarna ur `A Ledning`. Matchningen är **grafbaserad** (`Natverk` i skriptet): bara
JSON-filens brunnar är noder; varje ledningsdel delas där en sökt brunn ligger inom toleransen
(2 m) från *linjen* (vertex eller mitt på ett segment), fria ledningsändar blir noder och ändar
inom toleransen slås ihop. Vägen brunn→brunn söks med BFS (färst bitar, högst `MAX_HOPP−1`
andra sökta brunnar emellan, högst `MAX_DELAR` = 8 bitar; BFS:en håller bästa antal
mellanbrunnar per nod och får återbesöka en nod med färre – annars blockerade en slinga i nätet
den enda tillåtna vägen, rättat sep 2026) – så hittas sträckor uppdelade i flera
ledningsobjekt (fältet `ANT_DELAR` = antal ledningsobjekt, inte bitar) och brunnar utan egen vertex. Geometrin orienteras
startbrunn→slutbrunn. Första körningen med gamla vertex-metoden (steg2-skriptets) gav 136/180;
CSV:n över omatchade har kolumner med avstånd brunn→närmaste ledning samt `diagnos`
(`Natverk.diagnos`: "hoj max hopp till N", "glapp X m vid (x, y)", "annat lager", "brunnen … finns
inte i brunnslagren"). Grafmetoden gav 159/180 på DUF 701; resten är 13 littera som saknas i kartan
(`AG`/`STBEXTRA` är platshållare) och 3 par utan väg.
Fält: `MASK_BED` (alias "Maskinell bedömning", klass A–E från modellen), `MAN_BED`
("Manuell bedömning", fylls i för hand), samt de härledda `BEDOMNING` (manuell om ifylld,
annars maskinell), `BED_TYP` (Maskinell/Manuell) och `STIL` (`A - Maskinell`), samt `RAPPORT`/`VIDEO`
(absoluta sökvägar, PDF:en löses från JSON-filens mapp) för ArcMaps fältbaserade hyperlänkar
(Display > Support Hyperlinks using field, sparas i `.lyr`). Användaren kör analysen lokalt men
ArcMap i Citrix: verktygets kategori "Hyperlänkar" har `rapportmapp`/`filmmapp` (KONFIG:
`RAPPORTMAPP`/`FILMMAPP`) som bygger sökvägarna av filnamnet i en mapp som Citrix-klienten ser;
utdatamappen läggs på nätverket så JSON + `rapporter/` nås därifrån. Utdata skrivs i **2D**
(Z i SVOA:s nätverk är −9999) och verktyget kan skriva **GeoJSON enligt RFC 7946** (WGS84 via
`projectAs(4326)`, 2D, ingen `crs`, UTF-8) för webb-GIS – ArcMaps Features To JSON gav EPSG:3011
+ 4 koordinater per punkt som webb-GIS:et inte läste; shapefil-zippen saknade dessutom `.shx`. Symbologin
(Unique values på `STIL`, tio kategorier: färg efter klass, **streckad** = maskinell,
**heldragen** = manuell) sätts **manuellt en gång i ArcMap** och sparas som
`arcmap/bedomda_ledningar.lyr`, som verktyget använder som standard (handledningen 7.2).
En ArcObjects-generator (`skapa_lyr.py` + vendorerad comtypes) byggdes men togs bort på
användarens begäran – oprövad och onödig; `.lyr` är binär och kan inte skapas utanför ArcMap. "Uppdatera bedömning"
(eller `BARA_UPPDATERA = True`) räknar bara om de härledda fälten efter manuell ifyllnad; vid
full omkörning bevaras manuella bedömningar per brunnspar. Omatchade par listas i CSV.
Lagervalen i verktyget är **rullistor med kartans lagernamn (GPString)**, inte
GPFeatureLayer: ArcMap tolkar `/` i lagernamn (`A Rensbrunn/tillsynsbrunn`) som sökväg när namnet
skickas som text → "does not exist". Namnen slås upp till lagerobjekt via `arcpy.mapping`
(`hitta_lager`, matchar kort eller långt namn `Grupp\Lager`) och `kalla()` ger lagrets
`dataSource` + definitionsfråga till `MakeFeatureLayer`, så namnet aldrig går som text till GP.
Brunnar i SVOA-kartan: nedstigningsbrunnar (xNB/xNBL) i `A Nedstign och övriga brunnar`,
rens-/tillsynsbrunnar (xRB/xTB) i `A Rensbrunn/tillsynsbrunn`, plus `A Platsgjuten brunnspunkt` –
alla tre är standard.
**Svackor och bakfall som karta** (okt 2026, steg 1 i PLAN_verktyg.md): `kartunderlag.json` har
`svackpos_m` (kamerans position), `svacka_andel` och `bakfall_segment` (lista {fran_m, till_m,
langd_m, lutning_promille}; `Stracka._bakfall_segment` slår ihop stigande delsträckor, luckor och
segment < `BAKFALL_SEGMENT_MIN_M` 1 m; `BAKFALL_MIN_PROMILLE` 5). Skapa ledningslager har valfria
utdata `svackor_ut` (punkt, fält SVACKA_CM, SVACKLANGD, ANDEL_DIAM, POS_M) och `bakfall_ut` (linje,
LANGD_M, LUTNING, FRAN_M, TILL_M) + gemensamma UTG_BRUNN, MASK_BED, OSAKER, MATERIAL, DIMENSION, NR,
KALLFIL, RAPPORT (`skriv_svackor_bakfall`; alla poster i filen med matchat par, även syskon; svackor
< `SVACKA_MIN_CM` 2 hoppas över). Positionen läggs ut med `_kartposition` längs vägen a→b: skala
kartlängd/filmlängd, ingen skalning vid avbruten, vänd när utgångsbrunnen är b. Hjälpfunktionerna
`_utdata`/`_radera_utdata`/`_skapa_fc` bröts ut ur skapa() (shapefil/gdb, schemalås). GeoJSON
skrivs som `<stam>_svackor/_bakfall.geojson` när geojson_ut anges (skriv_geojson kan nu punkter).
Symbologi: `arcmap/svackor.lyr`/`bakfall.lyr` när användaren sparat dem. Testat med låtsas-arcpy
(scratchpad/test_svackor.py: skalning 1,25, avbruten utan skalning, motströms vändning).
**PowerPoint** (okt 2026, steg 4, `tv3_pptx.py`, python-pptx): `bygg_presentation(strackor, etapper,
diagram, utdata, mall, ut, topp, sidor, ta)` fyller `mall.pptx` via platshållar-idx per layout (`PH`:
100 = title, 101… i mallens ordning – pptxgenjs ger alla platshållare namnet "Text 0", så namnen går
inte att använda); klassen `Mall` tar bort exempelsidorna, klonar sidnummer-platshållaren från
layouten, passar in bilder i platshållarens ruta (platshållaren tas bort, även när bilden saknas),
bygger tabeller (ACCENT_1-rubrik, varannan rad BACKGROUND_2, klassfärg på klasskolumn, `TABELLRADER`
12 per sida → fortsättningssidor "(2)") och nyckeltalsrutor som i mallens exempel. KONFIG: `SIDOR`
(titel, avdelare, nyckeltal, klasser, topplista, topptabell, koder, material, etapper, strackor,
driftatgarder, svackor, karta, metod, avslut), `TOPP_STRACKOR` 10, `FOTON_PER_STRACKA` 3 (de
allvarligaste skadornas bilder först), `KARTBILD` kartbild.png, `AVSLUT_PUNKTER`. `--pptx` i
tv3_analys (efter Excel/JSON, innan diagramkatalogen städas); fristående `python tv3_pptx.py UTDATA`
läser kartunderlag.json (nya toppfält `mediakataloger`/`bildkataloger`), TV3-filerna igen (rättade
littera tas ur JSON på fil+nr), prioritering.xlsx (manuella), kostnader.csv → planera_atgarder,
rita_diagram i tempmapp. ArcMap: verktyget **Exportera kartbild** (ExportToPNG av aktuell vy,
bredd 2400 px, 200 dpi) i .pyt. Test: DUF 701 ger 26 sidor, validate.py OK, foton via låtsasbilder;
LibreOffice i sandlådan kan inte rendera, så layouten är kontrollerad strukturellt (python-pptx), inte
visuellt – användaren justerar mallen.
**Åtgärdspaket och kostnad** (okt 2026, steg 3, i `tv3_analys.py`): KONFIG `KOSTNADSFIL` (kostnader.csv i
repot, användarens justerade belopp okt 2026: strumpa 2 600/2 750/3 000/4 000/8 100 kr/m per dimensionsintervall
≤200/300/400/600/1200, hatt 25 000, lagning 25 000 kr/m, framschaktning 60 000, ny_brunn 45 000, etablering 45 000;
DUF 701 med bara A + överbryggande B: 10,5 Mkr, varav hattar 2,9 Mkr – hattpriset är den stora posten), `ATGARD_KLASSER` (A; var A, B t.o.m. okt 2026), `SCHAKT_AUTOMATISKT`
False (F3), `ETAPP_OVERBRYGGA_M` 60 (F5), `ETAPP_ORDNING` index|konsekvens, `BRUNNSTYP_ANDE_M` 1,5,
`LAGNINGSKODER`, `FRAMSCHAKTA_BRUNNAR`. Stracka: `manuell_bedomning`/`kommentar`/`lagning_m` (ur tidigare
prioritering.xlsx via `manuell:`/`--manuell`, `las_manuella` läser fliken Prioritering på rubriknamn och
Etapper-kolumnen "Schakta fram (manuellt)"; `koppla_manuella` fil+nr, annars brunnspar), `gallande_klass`
(manuell A–E före maskinell), `metod_auto`/`metod` (text schakt/strump/ingen i manuell styr; strumpa för
gällande A/B ej relinad; `_metod` sätts vid överbryggning), `brunnstyp(brunn)` (TVDAT NB/TB/RB vid änden,
annars `brunnstyp_ur_littera`: prefixet innehåller NB/TB/RB), `grad4_koder`, `lagningsbehov`,
`atgardsflagga` ("relinad men klass A", "grad 4 RBR – går strumpa?", "lagning? 2×YTS4", "brunnstyp okänd",
"bara tillsyns-/rensbrunnar – framschaktning", kostnads-/etappflaggor), `dimension_mm`.
`planera_atgarder(strackor, kostnader, framschakta)`: representant per brunnspar (värsta gällande klass,
syskon får etapp/metod men "kostnad räknad på nr X"), överbryggning, komponenter per metod via gemensamma
brunnar, ordning, F9 girig täckning per etapp (behov = strumpsträckor med TB/RB i båda ändar; manuellt
låsta brunnar först), kostnad per sträcka {strumpa, hattar, lagning, summa} och per etapp + brunnar +
etablering; schakt utan kostnad (F8). Excel: Prioritering får Brunnstyp start/slut (dolda), Etapp, Metod,
Kostnad (kr), Åtgärdsflagga, Manuell bedömning, Kommentar, Lagning (m) (förifyllda ur manuell:); ny flik
**Etapper** (en rad per etapp, "Sträckor (lista)" dold); Sammanfattning-block "Åtgärdspaket". JSON:
manuell_bedomning, gallande_bedomning, etapp, metod, kostnad_kr, lagning_m, brunnstyp_start/slut,
atgardsflagga; kartfält ETAPP/METOD/KOSTNAD och MAN_BED fylls från JSON när kartan saknar egen.
`--etapper nej` stänger av. DUF 701 utan manuellt: 17 etapper, 97 strumpa, 30 framschaktade brunnar,
19,1 Mkr schablon. Rundtursttest: lagning 3 m (+75 000), "schakt", "ingen", C→A och låst brunn fungerar.
**Kartexport** (okt 2026, `arcmap/kartexport.py` + verktyget Exportera kartor i .pyt, oprövat i riktig
ArcMap): `exportera(bedomda, ut_mapp, urval='atgard'|'AB'|'valda'|'alla', skalor, marginal, dpi,
markeringslager, samlad, per_etapp, skriv_falt, kartmapp)`. `dataram_matt(df)` = (bredd, höjd)/skala
läses EN gång (df.extent.width/df.scale – oberoende av sidenheter); `valj_skala` = minsta skala i
serien där utbredning + 2·marginal ryms, annars största + flagga; `_centrera` sätter extent = centrum
± ram·skala/2 och df.scale. Markering via definitionsfråga på ett markeringslager (återställs i
finally) eller SelectLayerByAttribute (CLEAR_SELECTION efteråt). Textelement TITEL/UNDERTITEL/SKALA
fylls om de finns (ListLayoutElements TEXT_ELEMENT, namn skiftlägesoberoende). ExportToPDF
PAGE_LAYOUT, PDFDocumentCreate → kartor_strackor.pdf/kartor_etapper.pdf, fält KARTA (254) med
sökväg (kartmapp för Citrix). Filnamn `<klass>_<fil>_<nr>_<fran>-<till>.pdf` / `etapp_NN.pdf`.
**Liggande/stående** (okt 2026; okt 2026 b: användarens SVOA-mallar UtskriftA4_Liggande/Stående.mxd saknade lagret → `Layout(kopiera_fran=(lyr, mark_lyr))`/`_lagg_in` kopierar lagret och markeringslagret från den öppna kartan in i mallens dataram med `arcpy.mapping.AddLayer` när `_lager_i_doc` inte hittar dem, loggat "kopierades in"; test OK2 i test_kartexport.py; användaren ville sedan ha "alla lager jag har aktiverade" → `kopiera_synliga` (param 13, GPBoolean, True): `Layout._kopiera_synliga(oppen_mxd)` lägger in den öppna kartans synliga toppnivålager (longName == name, grupper följer med hela) som inte redan finns i mallen, nederst först med AddLayer 'TOP' så ordningen bevaras; test OK3): `mall_liggande`/`mall_staende` (.mxd, DEMapDocument i dialogen,
kategori Mallar) öppnas med `arcpy.mapping.MapDocument(path)`; klassen `Layout` (mxd, df, ram,
liggande, lyr via `_lager_i_doc` på lagernamnet från den öppna kartan, mark_lyr, markera via
`setSelectionSet('NEW', oids)` med SelectLayerByAttribute som reserv, aterstall); `valj_layout` tar
layouten med nyckel (ryms inte, skala, −fyllnad) = minsta skala, vid lika den där sträckan fyller
ramen bäst; textelement och export sker i den valda mallens dokument; mallarna sparas inte (del).
Utan mallar en `Layout` för CURRENT. Resultat-tupler (fil, skala, ryms, layoutnamn).
Granskning okt 2026 (rättat): nyckeln i `valj_layout` är (ryms inte, skala, fyll om inte ryms annars
−fyll) – minst överskjutning när inget ryms; användarens urval sparas (`_gammalt_urval`) och
återställs i `aterstall`; `_lager_i_doc` provar långt namn, sedan kortnamn, hoppar grupplager;
`_huvudram` = största dataramen; `activeView` sätts till PAGE_LAYOUT; `kontrollera_kalla` varnar
när mallens lager pekar på annan featureklass; geometrin läses med `spatial_reference` =
dataramens SR; `Layout.stang()` släpper mallarna före AddField; filnamnskollisioner får löpnummer
och loggas; ETAPP som text tål `_filnamn`; undertitel per etapp listar alla metoder; `bedomda` som
sökväg ger tydligt fel; fast dataram ger begripligt fel; roterad dataram loggas.
Test: scratchpad/test_kartexport.py med låtsas-mapping (A3-ram 76×50 m i 1:200: 40 m → 1:200,
100 m → 1:400, 300 m hög → 1:1500; per etapp; valda). Handledning 7.10.
**Granskning okt 2026 av steg 1–4 (rättat, tre delgranskningar + egna fynd):** tv3_analys –
`manuell_klass()` kräver ensam bokstav A–E (fri text som "Bevaka" är inte klass B), `manuell_metod`
(schakt vinner; ingen/inget/ej åtgärd/avvakta = ingen) och manuellt "ingen" överbryggas inte;
representant per brunnspar: sträcka med manuell bedömning först, sedan klass/index, syskon med egen
manuell bedömning flaggas; sträcka med samma brunn i båda ändar får eget "par" (#fil/#nr) och egen
etapp; `las_manuella` tolkar Lagning med regex ("3 m" → 3, varning annars) och hoppar enhetsraden i
Etapper; `las_kostnader` tål "2.500,00"/"9 000 kr" och varnar vid otolkbara belopp; `pris()` använder
dimensionslös rad som reserv; "kostnadsfil saknas" i stället för "dimension saknar pris" när filen
saknas; `brunnstyp_film` tar koden närmast änden och `brunnstyp_konflikt` flaggar "brunnstyp? X: NB
enligt littera, RB i filmen"; bakfallssegment slår bara ihop luckor som inte faller och räknar
lutningen över de stigande bitarna; `--pptx` fångar alla fel (fel.txt) och skickar `a.topp`.
natverk – `Graf` behåller nollbitarna men grupperar noder förbundna med dem (`kanon`/`medlemmar`/
`grannar`, union-find lazy): sökningen passerar fritt, stopp gäller gruppen, brunn snett vid en skarv
bryter inte grafen; bitar identifieras på `id(pts)` (`fram_ids`, `noll`) så ringledningar av samma oid
räknas rätt; stoppbrunnar läggs inte i noder (ingen ANT_ANSL-skattning från dem); numeriska
riktningsvärden (1.0) normaliseras, okända värden räknas (`okant_attribut`) och varnas; saknade fält
varnas per lager; dubblettfält i SearchCursor dedupliceras; brunnar projiceras till ledningslagrets SR;
loggtexten säger "provade ledningar". skapa_ledningslager – GeoJSON för bakfall skrivs i rätt block
(låg inuti symbologitipsen och skrev svackpunkterna), symbologi appliceras före AddLayer, MAN_BED ur
JSON kräver ensam bokstav. tv3_pptx – sidnummer via `clone_placeholder` + `a:fld type=slidenum`
(clone_layout_placeholder finns inte i python-pptx 1.0.2), None-säkra KONFIG-texter (LOPANDE_*,
ATTRIBUTFAKTOR, GRAD4_KODER_A, GRADFAKTOR), `ValueError` i stället för SystemExit, STRACKA-sidor bara
för gällande klass i ATGARD_KLASSER och ett brunnspar en gång (`topp_strackor`, `--strackor`), inga
kapade tabeller, tomma fält utelämnas i fakta, undertitel 16 pt och max två filer/områden, radhöjd ≥
texten, 10 pt i sjukolumnstabeller, trasig bild hoppas över, tempmapp städas i finally, rättade
littera matchas på hel sökväg och utgångsbrunn sätts. Granskarnas testskript: scratchpad/rev_nat2.py
(skarv + ring), rev_geo.py.
**Uppströmsanalys** (okt 2026, steg 2, `arcmap/natverk.py` + verktyget Uppströms i .pyt, oprövat i
riktig ArcMap): `Graf(Natverk)` med ALLA brunnar som noder, `fram` = kanter i ritad riktning, bitar utan
längd (fri ände + brunn på samma ställe) tas bort i `_kant`; `riktning(lager, oid)` → attribut
(med/mot-värden ur dialogen, F6) → vattengång (vg_upp > vg_ned = MED) → ritad (loggas som antagen,
OBS om > 50 %); `uppstroms(start)` BFS bakåt, stoppfält (t.ex. tryckledning) och stoppbrunnar
(pumpstationer); serviser = eget lager eller fält i ledningslagret, ingår inte i grafen utan räknas
när en ände ligger inom toleransen från en uppströmskant (`Segmentindex`); `vald_ledning` ger start
från markerad ledning (nedströmsnod). Utdatalager UPPSTROMS_FALT (AVSTAND_M, NIVA, LANGD_M, SERVIS,
RIKTN_UR). `uppstroms_batch(bedomda, ...)`: start = slutbrunn per sträcka, skriver ANT_SERV_U/L_UPPSTR
till lagret + CSV `fil;nr;startbrunn;slutbrunn;serviser_uppstroms;langd_uppstroms_m;antal_ledningar;
antal_brunnar;serviser_kalla;stoppade`; utan serviser summeras ANT_ANSL (nytt fält i EGNA_FALT) för
inspekterade sträckor uppströms = "skattning (ANT_ANSL)". I `tv3_analys.py`: `uppstroms:`/`--uppstroms`,
`las_uppstroms`/`koppla_uppstroms` (fil+nr, annars brunnspar), `Stracka.serviser_uppstroms`,
`langd_uppstroms`, `serviser_kalla`; Excel-kolumner "Serviser uppströms", "Längd uppströms (m)"
(dolda), JSON `serviser_uppstroms`, `langd_uppstroms_m`. Test: scratchpad/test_natverk.py
(förgrening, slinga, attribut MOT, vattengång, tryckledning, serviser, batch + skattning).
**Markprofil** (sep 2026, `arcmap/markprofil.py` + verktyget Markprofil i .pyt, oprövat i riktig ArcMap):
läser bedömda-lagret (fält `NR` tillagt i EGNA_FALT för kopplingen), hämtar vattengång vid start-/
slutbrunn ur ursprungliga ledningslagret (fält vid ledningens start-/slutvertex, valda i dialogen,
gissas VG_UPP/VG_NED; ändpunkt inom tolerans från sträckans ände och grannvertex på linjen), samplar
markhöjd var 1 m längs kartlinjen ur ett punktlager (Z eller fält; IDW av ≤ 8 punkter inom 5 m,
rutnätsindex `Markpunkter`) och skriver `markprofil.json` (`strackor[]`: fil, nr, brunnar,
langd_karta_m, vg_start, vg_slut, mark [[m från startbrunn, z, avstånd]]). I `tv3_analys.py`:
`markprofil:`/`--markprofil`, `koppla_markprofil` (fil+nr, annars brunnspar), `Stracka.hojdanpassning`
(status: "RH2000 ur filen" |diff| ≤ `HOJD_SAMMA_M` 0,3; "förskjuten till GIS" om fallet stämmer inom
`HOJD_FALL_TOL_M` 0,3; "förskjuten och lutningskorrigerad" annars, offset + k·d från startbrunn;
"okänt nollplan" utan GIS; "GIS-vattengång, rät linje" utan filhöjder), `profil_korrigerad()`
(används av svacka och profilbild), `mark_i_filmens_axel()` (kartmeter skalas med langd_karta/langd
och vänds vid motströms), `tackning` (mark − (vattengång + dimension); rät linje mellan brunnshöjder
om profilen är osäker), `hojdflagga` ("Liten täckning" < `TACKNING_MIN_M` 1,0, "Ledning över mark –
höjdfel"), `lutning_promille` ur korrigerade höjder. Excel: Höjdanpassning, Täckning min/max,
Höjdflagga (före Manuell bedömning) + rad i Sammanfattning; JSON: hojdanpassning, hojd_offset_m,
gis_vg_*, tackning_*, hojdflagga; PDF: rad Höjdläge/Täckning, brun markyta i profilen, täckning
utsatt, höjdläge som text nere till vänster. **Avbrutna/ofullständiga inspektioner**
(`Stracka.ofullstandig`: KAM/HINDE eller langd < `OFULLSTANDIG_ANDEL` 0,85 × langd_karta): upphängning
bara i kamerans startbrunn (status "förskjuten till GIS vid X (avbruten inspektion)"), ingen
lutningskorrigering, ingen skalning (kartmeter från kamerans brunn = filmposition), mark bara
över filmad del, "avbrott" i stället för brunnsnamn i profilbilden; syskon filmade från andra
hållet hängs upp i sin egen brunn. Beräkningar cachas i `Stracka._cache` (rensas i
koppla_markprofil). Granskning okt 2026: vid avbrott tas filens brunnshöjder ur inklinometerns
ändpunkter, inte PROFILADM (kopierade/opålitliga); klass E (langd < 1) får ingen hojdanpassning;
bara en GIS-nivå → förskjutning mot den brunnen ("förskjuten till GIS vid X (bara en brunn har
GIS-nivå)"); statustexten använder `Stracka.hojdsystem` ur markprofil.json; `diameter_m` tolkar
"225/300"/"Ø 400"; `koppla_markprofil` matchar fil+nr, sedan fil+brunnspar, sist brunnspar
(syskon till avbrutna får syskonets post – bedömda-lagret har ett objekt per brunnspar);
hojdflagga "Filmad längd X m mot Y m i kartan – fel sträcka?" när langd > 1,25 × langd_karta;
profilbilden ritas även utan inklinometer när GIS-vattengång och mark finns (rät linje). Utan markprofil är allt oförändrat (regressionstestat). Testad med syntetisk
markprofil för DUF 701 (scratchpad) och låtsas-arcpy för verktyget.
**GIS-export och kontroll mot kartan** (okt 2026, `arcmap/gisexport.py` + verktyget Exportera GIS-data i
.pyt, oprövat i riktig ArcMap; användarens önskemål "exportera brunnarna och ledningssträckorna mellan
brunnar så vi kan kolla vg-nivåerna"): `exportera(ledningslager, brunnslager, brunn_id, json_ut, lock_falt,
typ_falt, vg_fran, vg_till, dim_falt, mat_falt, ledntyp_falt, ar_falt, tolerans, omradeslager, hojdsystem)`
bygger `natverk.Graf` med ALLA brunnar som noder, tar varje bit en gång (`fram_ids`, ej `noll`) och
interpolerar vattengången i delningspunkterna längs originalledningen (`markprofil._station`, `_vg_vid`:
ändvärden vid ≤ 0,05 m från änden). Locknivån ligger på brunnarna, vattengången på ledningarna (vg upp/ned
vid start-/slutvertex) – ingen bottennivå på brunnen (användaren). Skriver `gisdata.json` ({brunnar:
[littera, typ, lockniva, x, y, lager], ledningar: [fran, till, fran_xy, till_xy, langd_m, vg_fran, vg_till,
dimension, material, ledningstyp, anlaggningsar, lager, oid, del, antal_delar]}; fri ände = tomt namn) +
`gisdata_brunnar.csv`/`_ledningar.csv` (utf-8-sig, decimalkomma). Dialogen gissar fält (STANDARD_LOCKNIVA,
_BRUNNSTYP, _DIMENSION, _MATERIAL, _LEDNTYP, _AR, VG_FRAN/TILL). I `tv3_analys.py`: `gis:`/`--gis`,
`las_gis`, `koppla_gis` (brunnspar oavsett riktning, vid flera den med längd närmast filmens; sätter
`Stracka.gis` = {ledning, brunn_start, brunn_slut, vg_min_start/slut}; utan markprofil fylls gis_vg_start/
slut, langd_karta, hojdsystem så hojdanpassning, "fel sträcka?" och ofullstandig fungerar som med
markprofil), `Stracka.gisflagga` ("brunn saknas i GIS: X", "ingen ledning i GIS mellan brunnarna",
"vattengång: fall X m i filmen, Y m i GIS" när |diff| > max(`GIS_FALL_TOL_M` 0,3, `GIS_FALL_TOL_ANDEL`
0,5 × |GIS-fall|), "riktning: GIS-vattengången stiger …" (> `GIS_RIKTNING_MIN_M` 0,05), "vattengång
saknas i GIS", "material: …" via `_materialgrupp`/`GIS_MATERIAL` (BTG=Betong, PVC/PE/PP=Plast …),
"dimension: …"), `gis_lutning_promille`, `djup_start/slut` (locknivå − lägsta vg i brunnen),
`anlaggningsar`. Okända brunnar → `littera_forslag.csv` i utdata (brunnslittera-format med motbrunn, en rad
per sträcka): är den andra brunnen känd prövas dess grannar i GIS (`grannar` ur pa_par) och grannar vars
ledning har längd inom `GIS_LANGD_TOL` 0,15 av filmens föreslås, rankade på längddiff − 0,5·namnlikhet
(användarens idé okt 2026: "enbart en brunn fel – gissa rätt brunn med längden"); annars/utan träff
`difflib.get_close_matches` cutoff `GIS_LITTERA_LIKHET` 0,75 (= två omkastade siffror); kommentaren anger
metod och längder. Excel: GIS-flagga
(synlig), Lutning GIS, Djup start/slut, Anläggningsår (dolda), Sammanfattning-rad; JSON gis_flagga,
gis_lutning_promille, djup_start_m, djup_slut_m, anlaggningsar. Test: scratchpad/test_gisexport.py
(låtsas-arcpy: brunn mitt på ledning + snett 0,5 m, interpolerad vg, fri ände, brunn utan ledning) och
syntetisk gisdata för DUF 701 med inlagda fel (dimension, material, omvänd vg, fel fall, vg saknas,
ledning saknas, omkastat/borttaget littera) – alla hittades; facit utan gis oförändrat. Handledning 7.11.
**Inspektionsgrad per driftområde** (okt 2026, användarens önskemål "DUF-områden … hur mycket ledningsnät
finns och hur många är filmade"): exporten tar valfritt `duf_lager`/`duf_falt` (kategori Driftområden,
gissning STANDARD_DUF/STANDARD_DUF_NAMN); klassen `Omraden` i gisexport gör punkt-i-polygon i ren Python
(ringar ur SHAPE@, None skiljer hål, udda antal ringar = inne, bbox-förfilter) och sätter `omrade` på
brunnar (punkten) och ledningar (bitens mittpunkt via `markprofil._punkt_vid`). I tv3_analys:
`inspektionsgrad(strackor, gisfiler)` → rader per (område, ledningstyp) + "alla"-summor med ledningar,
längd i GIS, filmade (GIS-ledning kopplad till någon sträcka, GIS-längd räknas), andel, samt A/B (st),
AB_m och andel_AB av filmad längd (sämsta gällande klass per GIS-ledning, användarens fråga "hur många
procent A och B fel finns det"); `ej_inspekterat`
(GIS-ledningar utan film, sorterade område/år/material); `utan_gis` (filmade sträckor utan GIS-ledning).
Excel: flikar **Inspektionsgrad** (fetstil på alla-rader, andel i %) och **Ej inspekterat**, rad i
Sammanfattning (`_inspektionsgrad_text`), dold kolumn Driftområde (`Stracka.driftomrade`: ledningens,
annars brunnarnas); JSON toppfält `inspektionsgrad`, `ej_inspekterat`, per sträcka `driftomrade`.
Test: test_gisexport.py (två polygoner, hål, gräns) och gisdata_omr.json för DUF 701 (tre områden, 22
ofilmade ledningar → 83 %, DUF 701 91 %, 702 60 %, 703 0 %).
Idéer kvar: eget kartlager för ej inspekterade ledningar i Skapa ledningslager, uppströmsanalys ur
exporten utan ArcMap, djupklass i schaktlistan, ålder som konsekvensfaktor.
**Omfilmning och sammanslagning** (okt 2026, användarens fråga "ibland filmas ledningen en andra gång efter
spolning, eller från andra hållet – bygg båda"): KONFIG `OMFILMNING` True, `OMFILMNING_MIN_ANDEL` 0,85,
`SAMMANSLAGNING` True. `hantera_omfilmningar(strackor)` grupperar per (fil, brunnspar): hela filmer (langd ≥ 1,
ej ofullstandig) → nyaste gäller, äldre hela "ersatt av nr X (omfilmning datum)" (en kortare nyare än 85 %
av den äldre → den längre gäller, "(längre film)"), ofullständiga "ersatt av nr X (hel film)", tomma (langd
< 1) "ersatt av nr X"; utan hel film: per utgångsbrunn gäller den längsta delfilmen ("längre film från
samma brunn"), och finns en delfilm från vardera brunnen slås de ihop med `_sammanslagen(a, b, nr)`: a från
startbrunn, b från slutbrunn, L = langd_karta (markprofil/GIS) annars a+b (status "kartlängd okänd –
filmerna antas mötas utan överlapp"), överlapp = a+b−L (nyare filmens observationer gäller i
överlappet; lucka flaggas), b:s observationer speglas (lage' = L − lage − löpande längd, klocka
`_spegla_klocka` 3→9, 12/6 oförändrade, från/till byter), KAM och b:s brunnskod vid 0 tas bort, slutmarkör
vid L, `dataclasses.replace` av a med nr = max nr i filen + 1, profil tom (ingen profilbild/lutning), videofil
"a + b", manuell bedömning/gis/langd_karta ur a annars b, `sammanslagen_av=[a.nr, b.nr]`, delfilmerna
"ingår i sammanslagen nr Y tillsammans med nr X (filmad från BRUNN, L m) – se protokollet för sträcka Y" (okt 2026, användaren letade efter tvåfilmsritningen i delfilmens protokoll). `_ersatt()` lägger till "– OBS: den ersatta filmen var klass B, den gällande
är C" när den ersatta hade sämre klass (DUF 701 nr 118, B, 58,8 m avbruten, ersatt av hela 123 efter
litterarättning). `Stracka.filmstatus`, `aktiva(strackor)` = varken "ersatt" eller "ingår".
I main används `rakn = aktiva(...)` för statistik, diagram, kartunderlag, åtgärdspaket, PPTX, inspektionsgrad
och topplistan; rapporter skrivs för alla (även sammanslagna: positioner från startbrunnen, foton från båda).
Excel: Prioritering listar alla, ersatta/delfilmer sist utan rang, kolumn **Filmstatus** (synlig) efter
Inspekterad flera ggr; statistikflikar på aktiva; Observationer hoppar över sammanslagna (dubbletter); rad
"Omfilmningar" i Sammanfattning; JSON `filmstatus`, `sammanslagen_av`; PDF-rad Filmstatus. DUF 701: 62 (E)
ersatt av 74, 179 ersatt av 184 (hel film), 72+73 → nr 186 (C, 42,2 m), 170+171 → nr 187 (B, 48,6 m, 171
ensam var B 78 p/100 m, sammanslagen 44,9) ⇒ **facit med omfilmning 40/55/20/61/5** för filen ensam och
**40/54/20/61/5 med filer.txt** (litterarättningen gör 118/123 till samma par; 40/55/22/62/6 med
`OMFILMNING = SAMMANSLAGNING = False`). Överlapptest: 170/171 med langd_karta 40 → överlapp 8,6 m.
**Granskning okt 2026 av GIS-export, GIS-koppling och omfilmning (tre agenter + egna fynd, rättat):**
gisexport – stubbar (bit ≤ tolerans som slutar i samma brunn eller fri ände, dvs. ledningsänden sticker ut
förbi brunnen) hoppas över och loggas; driftområdena projiceras till ledningslagrets SR; `del` numreras
längs ledningen (station `_st`); ringledning: `st2 < st1` → `st2 = L`; nyckel på `longName`; `_tal` ger
None under `SAKNAS_UNDER` −999; dubbelt fält i brunnslagret läses en gång; .pyt laddar beroendena före
gisexport/markprofil. tv3_analys GIS – `las_gis` tål trasig fil (ValueError), sätter vg/locknivå ≤
`GIS_SAKNAS_UNDER` −999 till None, tolkar dimension/årtal med regex; dubbletter mellan GIS-filer
(`_dubblett` på (lager, oid, del) eller (par, längd)) räknas en gång; vid flera ledningar mellan samma
brunnar väljs samma ledningstyp (första bokstaven) före längd; `GIS_PLATSHALLARE` {AG, STBEXTRA} undantas
från "brunn saknas" och förslagen; `alla.get()` mot KeyError; `|fall| > GIS_FALL_ORIMLIGT_M` 50 →
"vattengång orimlig i GIS" och riktning/fall flaggas inte båda; E-sträckor räknas inte som filmade;
årtal som text kraschar inte sorteringen; `littera_forslag.csv` fyller `ratt` bara när längden ger en
tydlig kandidat (`sakert`: ensam eller ≥ 0,05 bättre poäng), kommentaren utan ;/, och filen tas bort när
inga okända finns; konsolens inspektionsgrad räknas efter omfilmningen (samma `gisstat` som Excel);
Sammanfattningsraden heter "Höjdläge (markprofil/GIS från ArcMap)"; hojdanpassning säger "(ofullständig
film)" när filmen bara är kort utan KAM; GIS_MATERIAL har LERGODS/PEM/PRC. Omfilmning – B-rader speglas
till L − lage + löpande längd (låg på startpositionen), öppen A-rad utan längd får `b.langd − lage`, den
äldre filmens löpande skador kapas vid snittet, utan kartlängd räknas samma observation (kod, grad,
infokod, attribut, löpande) inom 1 m från mötet en gång (nyaste behålls; status "samma observation vid
mötet räknas en gång" – 186 hade INH4 dubbelt), orimlig kartlängd (< längsta delfilm − 0,5) används inte
(status), datum + klockslag från den nyare filmen, "film nr X" i varje observations kommentar, b:s
brunnskod vid 0 behålls (speglas till L), `_tidsnyckel` tolkar datum/klockslag (ISO, d.m.Y, d/m/Y, ÅÅÅÅMMDD)
i stället för strängsortering, gällande hel film väljs först och statusarna sätts efteråt (inga kedjor),
tomma filmer pekar på den sammanslagna, SAMMANSLAGNING fungerar utan OMFILMNING, grupperingen använder
`_normlittera` även i `flerinspekterad`, `tv3_pptx.main` kör `aktiva(hantera_omfilmningar(...))`,
Prioritering blankar Etapp/Metod/Kostnad/Åtgärdsflagga för rader utan rang, `id()`-mängd i stället för
`in`. Kvar/design: samma brunnspar i olika TV3-filer hanteras numera av `_over_filer` (se Ny inspektion över filgränser);
testfiler (gis_test/, mp/, upp*.csv) som hamnat i repot togs bort och .gitignore:ades.
**Ny inspektion över filgränser** (okt 2026, användarens begäran "om en ny inspektion finns på en sträcka så vill
jag att den äldre stryks … lämnas kvar som referens för hur mycket sträckan förvärrats, inte räknas dubbelt"):
KONFIG `OMFILMNING_OVER_FILER` True, `TIDIGARE_INDEX_ANDEL` 0,2. `_over_filer(strackor)` körs sist i
`hantera_omfilmningar` på `aktiva()` (langd ≥ 1) grupperade på normaliserat brunnspar oavsett fil; grupper med
≥ 2 filer: nyaste filmen (`_tidsnyckel`, fil, nr) gäller, alla filmer i andra filer får
`filmstatus = "ersatt av ny inspektion nr X i FIL DATUM"` (+ "(samma datum – dubblett?)" vid lika dag, + "OBS: den
nya filmen är ofullständig, 10,2 m mot 28,9 m" när den nya är ofullständig men den ersatta inte; `_ersatt` tar
nu `text`); gällande (alla kvarvarande i nyaste filen) får `Stracka.tidigare` = {fil, nr, datum, klass, index,
langd, antal_skador, utveckling, antal} från den närmast föregående riktiga inspektionen (inte dubbletter);
`_utveckling`: klassbyte avgör, inom samma klass index ändrat ≥ max(5, 0,2·gammalt) = förvärrad/förbättrad.
Utan tolkbart datum rörs gruppen inte (OBS i konsolen). `tidigare_text(s)` = "2021-05-11 (DUF 701.TV3 nr 1):
B 41 → A 97 p/100 m – förvärrad (N äldre filmer)". Excel: synlig kolumn **Tidigare inspektion** efter
Filmstatus (bredd 44), Sammanfattning-rad "Ny inspektion i senare fil" (`_tidigare_text_summa`); JSON
`tidigare_inspektion` {fil, nr, datum, bedomning, konstruktionsindex, langd_m, antal_skador, utveckling}; PDF
infotabellrad; PPTX fakta-rad; konsolutskrift. Test: scratchpad/ny/ (syntetisk "DUF 701 omg2.TV3" 2024 med
sju sträckor ur DUF 701: förvärrad RBR4, relinad utan skador → D, oförändrad, avbruten vid 10 m, motströms
filmad, samma dag = dubblett, 2019 = äldre än DUF 701) – alla fall rätt; facit DUF 701 ensam och filer.txt
oförändrade (filen har inga par över filgränser).
**Kartlängd utan direkt GIS-ledning** (okt 2026, användarens fråga "varför står det att kartlängden är okänd, jag
har ju GIS-data"): KONFIG `GIS_VAG_MAX_HOPP` 4, `GIS_VAG_MAX_ANDEL` 2,0. I `koppla_gis`, när båda brunnarna finns men
ingen ledning: `_gis_vag(ns, ne, grannar, langd)` (Dijkstra via `grannar`, tak = andel × filmlängd + 20 m) ger en
syntetisk ledningspost {fran, till, langd_m, vg i ändarna, dimension/material/ledningstyp/år om alla delar är lika,
omrade ur första delen, `_syntetisk` "via", `_via` [mellanbrunnar], `_delar` [ledningar]}; annars brunnarnas avstånd
fågelvägen (`_syntetisk` "fagelvag", bara langd_m, ≥ 1 m – (0,0)-koordinater ger ingen). gisflagga: "ingen direkt
ledning i GIS – kartlängd 17,4 m via KTB99999 (2 ledningar)" resp. "ingen ledning i GIS mellan brunnarna – kartlängd
fågelvägen 44,7 m" (fågelvägen ger inga fler flaggor). `inspektionsgrad` räknar `_delar` som filmade, fågelväg
som utan GIS-ledning (`utan_gis`). Effekt: sammanslagningen får kartlängd och överlapp. Test: scratchpad/gisvag/
(gisdata med KTB72028–KTB61192 delad via KTB99999 och SRB63043–SRB62385 borttagen med koordinater → 187 överlapp
3,9 m); gisdata_omr och filer.txt oförändrade.
**Lista fält** (okt 2026, användaren: "GIS-exporten plockar inte upp några fält … finns något skript som dumpar
fälten?"): `arcmap/lista_falt.py` (Py2.7, ASCII; `lista(utfil, bara)` skriver per lager i CURRENT: longName,
datakalla, geometri/SR, antal objekt, och per fält namn/alias/typ/längd + upp till `EXEMPEL` 6 distinkta värden ur
de första `RADER` 2000 raderna via da.SearchCursor; `BARA_LAGER` filtrerar på namn; körbar med execfile + `UTFIL`)
och verktyget **Lista fält** först i .pyt (param: textfil, rader, bara-lager multiValue). Röktestad med låtsas-
arcpy. Avsikt: användaren skickar textfilen så att STANDARD_*-gissningarna i .pyt kan kompletteras med SVOA:s
riktiga fältnamn (ledningslager: vg upp/ned, dimension, material, ledningstyp, anläggningsår; brunnslager:
littera/EntityID, brunnstyp, locknivå; DUF-lager: områdesnamn).
**Lista fält hängde + minne av senaste val** (okt 2026, användaren: "tar superlång tid … 2 min 30 s, avbröt";
"pre-selecta de fält jag valt i senaste körningen"): `lista_falt.py` räknar inte längre objekt (GetCount mot SDE
tog minuter per lager), hoppar lager som inte är featurelager, skriver filen lager för lager med flush och
AddMessage "Laser X … klart pa N s" (kvar efter avbrott). `.pyt`: `MINNESFIL` = `arcmap/senaste_val.json`
(.gitignore), `_minne_fyll(verktyg, params)` i slutet av varje `getParameterInfo` fyller tomma inparametrar
(multiValue via `_satt_varden`) med senaste körningens värden per verktygsklass, `_minne_spara` först i varje
`execute` (valueAsText, citattecken bort för multiValue); gissningarna i `updateParameters` sker bara när
fältet är tomt (`not altered and not valueAsText`), så minnet inte skrivs över. gisexport: `Omraden` behåller
arcpy-geometrin och använder `contains(PointGeometry)` (ren Python-ringtest bara som reserv/test – DUF-polygoner
med tusentals hörn gjorde punkt-i-polygon per brunn/ledning långsamt), `exportera` loggar tid per fas och var
5000:e ledning. Användarens lager: Avlopp\A Rensbrunn/tillsynsbrunn, A Nedstign och övriga brunnar,
A Platsgjuten brunnspunkt, A Ledning, A Servis; Områden\DUF-områden.
**Bara A får åtgärd, B överbryggar** (okt 2026, användaren: "huvudsakligen enbart A-ledningar; en eller två
B-ledningar mellan två A kan tas med – inte strumpa B bara för att vi är där"): KONFIG `ATGARD_KLASSER = ("A",)`,
`OVERBRYGGA_KLASSER = ("B",)`, `OVERBRYGGA_MAX_STRACKOR = 2`. I `planera_atgarder` ersätter en kedjesökning F5-regeln:
`overbryggbar(r)` = metod "ingen", ej relinad, ej manuellt ingen, och (klass i OVERBRYGGA_KLASSER oavsett längd,
eller C/D kortare än ETAPP_OVERBRYGGA_M); per metod (strumpa, schakt) djupet-först från varje åtgärdsbrunn över
kandidatsträckor (högst max i rad, ingen återbesökt brunn/sträcka, stopp vid åtgärdsbrunn) – når vägen en annan
åtgärdsbrunn tas alla sträckor på vägen med (`_metod`, etapp_flagga "medtagen för sammanhang (klass B mellan två
åtgärdssträckor)" resp. "medtagen för sammanhang"). DUF 701: 15 etapper, 40 A + 9 B strumpa, 18 framschaktade
brunnar, 9,5 Mkr (var 17/97/30/19,1 Mkr med A+B). Klassfacit oförändrat. tv3_pptx STRACKA-sidor följer
ATGARD_KLASSER (bara A). Handledning och metodbeskrivning uppdaterade.
**Etapplager** (okt 2026, användarens önskemål "exportera ett lager för etapper så jag ser de sammanhängande
etapperna och vilket etappnummer de har – eget verktyg"): `etapper_for_karta(etapper)` i tv3_analys skriver toppfältet
`etapper` i kartunderlag.json (nr, metod, hogsta_klass, max_konstruktionsindex, antal_strackor, langd_m, schakt_m,
dimensioner, brunnar, framschaktade, framschakt_manuell, serviser_uppstroms, kostnad_kr (None för schakt), kostnad
{strumpa, hattar, lagning, brunnar, etablering, summa}, flaggor, strackor [{fil, nr, startbrunn, slutbrunn, bedomning,
langd_m}]); `skriv_kartunderlag(..., etapper)`. `arcmap/etapplager.py` (Py2.7, ASCII): `las_bedomda(bedomda)` läser
ETAPP/METOD/KOSTNAD/BEDOMNING/INDEX_K/LANGD_M/FRAN_BRUNN/TILL_BRUNN/NR/TV3_FIL + SHAPE@ via `kalla` (kräver ETAPP),
`skapa_etapplager(bedomda, ut_fc, json_fil, brunnar_ut, lyr_fil, lyr_brunnar)` grupperar per ETAPP > 0, bygger en
flerdelad Polyline per etapp (arcpy.Array av delar), fält ETAPP_FALT (ETAPP, METOD, HOGSTA_KL, MAX_IDX, ANT_STR,
LANGD_M, SCHAKT_M, KOSTNAD, DIMENSION, BRUNNAR, FRAMSCHAKT, ANT_FRAM, SERV_UPP, FLAGGOR, STRACKOR, ETIKETT "Etapp 1 -
strumpa, 100 m, 285 tkr, 1 brunn att schakta fram"); uppgifter ur JSON:s etapper (`_etapp_ur_json`, varning om
saknas) annars ur lagret; punktlager BRUNN_FALT för framschaktade brunnar med läge ur `_brunnsposition` (första/sista
punkt på sträcka med litterat); `_lagg_i_kartan` med `arcmap/etapper.lyr`/`framschaktning.lyr` om de finns, annars
tips. Verktyget **Etapplager (åtgärdspaket)** i .pyt (bedomda-rullista förvald `LAGERNAMN_BEDOMDA` 'Bedomda ledningar',
json_in valfri, ut_fc förslag `<bedömda>_etapper`, brunnar_ut valfri, lyr-filer under Symbologi). Test:
scratchpad/test_etapplager.py (två etapper, flerdelad geometri, brunn K2 på (50,0), utan JSON summeras KOSTNAD).
Handledning 7.12.
**SVOA:s fältnamn** (okt 2026, ur användarens kartans_falt.txt, Enkel VA_mall.mxd, sde_geopipe_prod): brunnslagren
(pipeSewerWellP, samma featureklass för Nedstign/Rensbrunn/Platsgjuten, olika definitionsfrågor): EntityID littera
(SNBL63489, KNB2635), CoverLevel locknivå, BottomLevel (tom), WellFunction (SNBL/DNB/KNB …), WellType heltal,
ConstructionYear. Ledningar (pipeSewerPipe): EntityID (DSL127791), PipeType DSL/SSL/KSL/SHL (första bokstaven
D/S/K), PipeMaterial Bt/Seg/PVC, PipeDimension text "300", ConstructionYear, RestorationYear/RestorationMethod,
LevelFrom/LevelTo (vattengång, RH2000-nivåer ~20–40), Slope, Length, SewerFunction/SewerWaterType heltal,
Commentary ("Digitaliseringsriktning flippad"). Servis (pipeSewerServicePipe) samma fält, Serviskopplingspunkt
(pipeSewerServicePoint) EntityID SKP/KKP/DKP, Ledningsände avlopp (pipeSewerPipeEnd). DUF-områden (geodata
Admindelning Dufomr): DUF heltal, DRIFTOMR. STANDARD_* i .pyt har SVOA-namnen först (+ STANDARD_RENOVERINGSAR,
STANDARD_SERVIS 'A Servis' som förval i Uppströms). Nytt exportfält `ren_falt` (param 17) → `renoveringsar` i
gisdata.json/CSV + `lager.renoveringsar`; tv3_analys: `las_gis` sätter `_ren_exporterad` per post (bara när
fältet valdes), `Stracka.gis_renoveringsar`, gisflagga "renoverad 2015 enligt GIS men inte relinad enligt filmen"
/ "relinad enligt filmen men inget renoveringsår i GIS" (16 i DUF 701 med syntetiskt fält), dold kolumn
"Renoveringsår GIS", JSON `renoveringsar_gis`; GIS_MATERIAL har BT/SEG/GJ/ODEF(""). lista_falt läser via
dataSource + definitionQuery (lagernamn med "/" gav "does not exist" i ListFields).
**Skarv utan brunn i GIS** (okt 2026, användaren: "ledningen är bruten i mitten, materialförändring, delad som två
ledningar utan ände – ändarna har nästan exakt samma koordinat"; "kan vara uppdelad i mer än två delar"): exporten
skriver sådana delar som `fran`/`till` = "" (fri ände, `_nodnamn`) med `fran_xy`/`till_xy`. `las_gis` kör
`_sammanfoga_fria_andar(ledningar, tol)` (tol = `GIS_SKARV_TOL_M`, None ⇒ filens `tolerans_m`, annars 1 m): fria
ändar rutnätsindexeras, parvisa ömsesidiga möten inom tol blir `partner` (tre ändar på samma plats = förgrening,
fogas inte), kedjor följs från en brunnsände via partners tills en brunn nås (valfritt antal delar, oberoende av
ritriktning och listordning); den hopfogade posten ersätter delarna: fran/till = brunnarna, langd summerad, vg ur
ändarna, dimension/material/ledningstyp/år bara om lika, `antal_delar`, `delar` [{langd_m, material, dimension,
anlaggningsar, oid}], `skarvar` [xy]; konsolrad "N ledningar hopfogade …". gisflagga "GIS: ledningen består av BTG
225 10,0 m + PVC 225 8,9 m + … (skarv utan brunn)" bara när delarna skiljer sig. Inspektionsgrad räknar den som en
ledning (delarna hoppades förut över som fria ändar). Test: scratchpad/gisvag/gisdata_skarv.json (två delar + en
trevägsförgrening som inte fogas) och gisdata_skarv3.json (tre delar, blandad ordning, mittdelen vänd → 28,86 m,
187 överlapp 19,8 m).
**Delfilmer göms** (okt 2026, användaren: "när vi slagit samman två delsträckor vill jag inte att den individuella
listas separat i Excel eller som egen rapport"): KONFIG `DELFILMER_SEPARAT` False; `delfilm(s)` (filmstatus "ingår…"),
`visade(strackor)` = alla utom delfilmer om inte flaggan. Prioritering listar bara `visade(alla)` utan rang (ersatta
kvar), Observationer tar den sammanslagna sträckans observationer (delfilmerna hoppas över; med flaggan tvärtom som
förut), `skriv_rapporter` skriver bara `visade` och `rensa_gamla_rapporter` tar bort delfilmernas gamla PDF:er
(aktuella = visade, prefix = alla), Sammanfattning-texten anpassad. DUF 701: 187 rader i Prioritering (var 191).
**Höjdfel i TV3-filen** (okt 2026, användarens rapport DUF 700 Ålsten Del 5 nr 26: lutning −26 111,8 ‰ och lodrät
profil – brunnshöjderna i filen skiljer ~2 400 m): KONFIG `HOJD_SAKNAS_UNDER` −999 (PROFILADM/PROFILDAT-höjder
under det → None, som GIS_SAKNAS_UNDER), `HOJD_ORIMLIG_M` 50. `kontrollera_hojder(s)` körs i `las_tv3` efter
profilsorteringen: |zs − ze| > 50 → brunnshöjderna None, "brunnshöjder 29,30 / 2444,33 m"; profilens
max − min > 50 → `profil = []`, "inklinometerprofil 2547,0 till 2864,0 m"; `Stracka.hojdfel` = "höjdfel i filen:
…" (OBS per fil i konsolen). `hojdflagga` returnerar hojdfel först (Excel Höjdflagga, JSON `hojdflagga` +
`hojdfel`), Sammanfattning-rad "Höjdfel i TV3-filen", PDF: rad Höjdfel (när ingen höjdanpassning), Lutning
"okänd (höjdfel i filen)", profiltexten "… – profilen ritas inte". `lutning_promille` faller nu tillbaka på
`_filens_brunnshojder()` (inklinometerns ändpunkter) när PROFILADM saknas/är orimlig – `lutning_ur_inklinometer`
ger "(ur inklinometern)" i PDF:en (10 sträckor i DUF 701 saknar PROFILADM men har inklinometer; klasserna
oförändrade); OBS-texten i profilbilden säger "brunnshöjder saknas eller är orimliga i filen – profilen kan
inte kontrolleras" när PROFILADM saknas. Test: scratchpad/hojdfel/ (DUF 701 med nr 1 sluthöjd 2444,33, nr 2
profil ×100, nr 3 −9999). Den riktiga filen (DUF 700 Ålsten Del 5.TV3) har inte setts – användaren ombedd skicka.
**Sammanslagen sträcka i översikten** (okt 2026, användarens begäran "syns i ritningen att den är sammansatt
av två avbrutna filmningar"): `Stracka.sammanslagning` = {a_nr, b_nr, a_fran, b_fran, a_langd, b_langd, L,
overlapp} sätts av `_sammanslagen`; `rita_schema` ritar film b:s del [L − b_langd, L] i blågrå ton
(#d6dde8) ovanpå röret, överlapp skrafferat, lucka vit, streckade skarvlinjer vid a_slut/b_start, och under
skalan två pilar (blå → för film a från a_fran, mörkblå ← för film b från b_fran) med "film nr X från
BRUNN, 0–17,0 m" samt raden "sammanslagen av två avbrutna filmer – skarv vid/överlapp/lucka …";
figurhöjden får 0,3 extra. Samtidigt: hjässa/botten-etiketter packas tillsammans med höger-etiketterna
(båda under röret; "17,0 m kl 12" låg över "17,1 m kl 1").
**Projektering** (sep 2026, `arcmap/projektering.py`, verktygen Skapa projekteringslager och
Projekteringsprofil i .pyt, oprövat i riktig ArcMap; ej kopplat till TV-inspektionerna):
`skapa_projekteringslager(gdb, prefix, sr/sr_lager)` skapar `<prefix>_Ledning` (LEDN_ID, TYP med
domän PROJ_LEDNTYP S/D/V/K/T/O, DIM, MATERIAL, VG_UPP, VG_NED, KOMMENTAR) och `<prefix>_Brunn`
(BRUNN_ID, TYP, LOCKNIVA, BOTTENNIVA, DIAM, KOMMENTAR); ritas med vanlig redigering i
flödesriktningen. `profil(...)`: ledningar (valda om urval finns – lagerobjektets
`getSelectionSet`), brunnar, `Noder` (ändpunkt → brunn inom tolerans 1 m, annars fri ände),
`_kedjor` per typ (start = angiven brunn, annars nod utan inkommande med högst VG_UPP;
förgrening → längsta grenen, resten egna kedjor), längsta kedjan = referensaxel, andra typer
projiceras med `_station` (båda ändar inom 3×tol+2 m, annars eget stråk), `Markhojd` (punkter
IDW som Markprofil, eller raster via `GetCellValue_management` med cache), täckning per ledning,
CSV (anmärkning: bakfall / ledning ovan mark / vattengång saknas), JSON, och `rita_profil`
(matplotlib i ArcMaps Python, A3 liggande, exakta skalor ur LANGDSKALOR/HOJDSKALOR, typfärger
S röd D grön V blå K lila, hjässa streckad, brunnar som rektanglar lock→botten med lyft etikett
när brunnar ligger < 4 m isär, nivåtabell under diagrammet med sektioner inom 1 m hopslagna).
Användarens val: alla ledningstyper i samma diagram, inga minimikrav. Testad med låtsas-arcpy
(scratchpad/test_projektering.py) och renderad PNG.
Granskning okt 2026 (rättat): **Markprofil** – `Ledningar` ersätter `Ledningsandar`: vattengången
interpoleras linjärt längs hela ledningsobjektet (brunnar mitt på en ledning, Natverk-klippta
sträckor, fick annars ingen nivå; riktningskontroll = punkt 1 m bort längs ledningen ligger på
sträckans linje); OBS-logg när vattengången stiger i ritad riktning på > 50 % av ledningarna
(fält/ritriktning); `hasZ`-kontroll och Multipoint-stöd (SHAPE@) för markhöjder; JSON-förslaget
hamnar inte inuti .gdb-mappen. **Projektering** – `Noder.nod(p, typ)` föredrar brunn av samma typ
(S-ledning → S-brunn) och snappar fria ändar mot redan matchade ändpunkter; `_station(forlang=True)`
låter parallella ledningar sticka ut förbi referensaxeln (station < 0 / > L, mark samplas längs
förlängningen med `_langs_forlangd`); rasterlager slås upp via `isRasterLayer` före featurelager
och läses en gång med `RasterToNumPyArray` över stråkets bbox (bilinjär), `GetCellValue` bara som
reserv; `cursor_kalla` behåller definitionsfrågan (tillfälligt lager utan urval, lagerobjekt med);
brunnar läses alltid utan urval; vg 0,0 räknas som värde; startbrunn utan utgående ledning loggas;
CSV med decimalkomma; matplotlib 1.x: `'_nolegend_'`, `set_title(loc=)` i try, `legend(prop=)`;
gdb-parametern är Output (får skapas), prefix A-Z/0-9/_ .
Testas utan ArcMap med en låtsas-arcpy (se sessionshistorik) – arcpy-körningen i sig är oprövad.
Kodgranskning sep 2026 (rättat, testat med låtsas-arcpy): diagnos skiljer på för många brunnar och
för många bitar; tolerans ≤ 0 stoppas i verktyget och `Natverk`; shapefil/geodatabas avgörs med
`Describe(ut_ws).dataType == 'Folder'` (feature dataset/.mdb/.sde funkade inte; domänen läggs i
geodatabasen ovanför ett feature dataset); befintligt utdatalager tas bort ur kartan
(`RemoveLayer`) och raderas före `CreateFeatureclass` (schemalås vid omkörning); manuella
bedömningar som inte matchas igen loggas och skrivs i CSV:n (kolumn `manuell_bedomning`);
`kopiera_falt` läses per ledningslager (saknat fält → NULL) och utdatanamnet tas från
`ListFields` efter `AddField` (shapefil kortar >10 tecken); DATE-fält behåller NULL i shapefil
och GeoJSON serialiserar datum som text; brunnar i annat koordinatsystem projiceras till
ledningslagrets; brunns-id på flera ställen (> tolerans isär) varnas; `logg`/`txt` tål stdout
utan teckenkodning och undantag med cp1252-bytes; sträckor utan brunnspar räknas i loggen.
Verktygsetiketten säger nu "inkl. slutbrunnen (2 = en mellanbrunn)". Kvar: ingen
datumtransformation vid `projectAs(4326)` (rätt för SWEREF 99, fel ~100 m för RT90).

## 3. TV3-formatet (Svenskt Vatten TV-fil v3.0, P93-koder) – det vi lärt oss

- Textfil, CRLF, **Windows-1252/ISO-8859-1** (åäö i koder som `LÄNGS`, `TVÄRS`, `UTFÄL`).
  Läs som bytes; prova UTF-8, fall tillbaka till cp1252. Decimalkomma överallt.
- Sektioner: `#TVADM` (en rad/sträcka), `#TVDAT` (en rad/observation), `#PROFILADM`,
  `#PROFILDAT` (inklinometer, kan vara >100 000 rader), `#SLUT`. Fält separeras med `;`.
- TVADM-index (0-baserat): 0 sträcknr, 2 startbrunn (uppströms), 3 slutbrunn (nedströms),
  4 **utgångsbrunn (där kameran startade – position 0 m räknas härifrån!)**, 7 ägare, 8 område,
  (Konventionen startbrunn = uppströms är verifierad: i DUF 700/701 gäller Medströms ⇔ utgångsbrunn =
  startbrunn (188 st) och Motströms ⇔ utgångsbrunn = slutbrunn (94 st), inga undantag. Flödes-
  orienteringen tas ur utgångsbrunnen, med Riktning som reserv. Motsäger Riktning och utgångsbrunn
  varandra (`brunnar_omvanda`) antas filen ange brunnarna i kamerans riktning: start/slut och
  PROFILADM-höjderna byts vid inläsning och det loggas med "OBS".)
  9 projekt, 10 riktning (Medströms/Motströms), 12 datum, 13 tid, 14 operatör, 18 videofil,
  22 form, 23 dimension, 24 dimension 2, 25 material, 26 foder, 27 fodermaterial,
  29 ledningstyp, 31 väder. 38 kolumner i exempelfilen.
- TVDAT-index: 0 sträcknr, 1 läge (m), 2 tid i film, 3 löpande-markering (`A1`…/`B1`…),
  4 skadekod, 5 grad 1–4, 6 infokod, 7 attribut, 11 klocka från, 12 klocka till,
  13 vattennivå %, 14 bild A, 15 bild B, 16 videofil (första raden), 17 kommentar.
  Sträckans längd = största läge. Rader med `B` är slutmarkering för löpande skada och ska inte
  räknas igen.
- PROFILADM: index 16/17 = start-/sluthöjd vid **start-/slutbrunn** (flödesriktning), inte vid
  utgångsbrunnen. **Vid avbruten inspektion är "sluthöjden" bara inklinometerns höjd där kameran
  stannade** (DUF 701 nr 72: 16,04/15,75 = PROFILDAT:s ändar), och för det motströms filmade
  syskonet (nr 73) är PROFILADM bara kopierat från nr 72 – lita inte på PROFILADM:s slutvärde
  när inspektionen är ofullständig. PROFILDAT: 1 läge, 2 relativ höjd, 3 absolut höjd; **kan ligga i fallande
  positionsordning – sortera alltid på läge**. Höjder har 2 decimaler (cm), vilket ger
  trappstegsprofil på flacka ledningar → utjämna vid uppritning, inte vid beräkning.
- Video/bild finns bara som filnamn i filen. Skriptet indexerar mediamappar rekursivt
  (första träff på filnamnet vinner).
- Koder (tolkade ur P93 + filen; kan behöva stämmas av mot ledningsägarens lista):
  Konstruktion **SPR** sprickor (KOMPL/CIRK/LÄNGS), **RBR** rörbrott, **DEF** deformation,
  **YTS** ytskada, **FOG** fogförskjutning (RIKTN/TVÄRS/LÄFÖR), **FRF** fog/rörfel (FAFOG),
  **DEA** defekt anslutning (EJÖPP). Drift **ROT** rötter (TUNNA/GROVA/PAKET), **INL** inläckage,
  **UTF** utfällning, **SED** sediment, **INH** inträngande hinder. Info **KAM** (HINDE =
  inspektion avbruten). Infokoder utan grad: AS anslutning (PGREN/INHUG/BORRA/SAHUG), AG,
  RB rensbrunn, NB nedstigningsbrunn, TB tillsynsbrunn, RP reparation (LAGAT), BR böj,
  DF dimensionsförändring, MF materialförändring, ST stalp, PP.
- Klockposition för anslutningar: kl 7–11 = vänster, 1–5 = höger, 12 hjässa, 6 botten –
  **sett i inspektionsriktningen**.
- Samma brunnspar kan förekomma två gånger (filmat från båda håll efter avbrott, eller omfilmat) –
  flaggas; sedan okt 2026 ersätter hel film ofullständiga och två delfilmer slås ihop (se Omfilmning).

## 4. Poängmodell och prioritetsklass (alla parametrar under KONFIG i skriptet)

- Grad → poäng: 1 → 1, 2 → 3, 3 → 10, 4 → 30. Driftskador × 0,5. **Löpande skador viktas med
  längden** (sep 2026, användarens val efter simulering): poäng × längd / `LOPANDE_ENHET_M` (10 m),
  minst 1, högst `LOPANDE_TAK` (5). Tidigare räknades de en gång oavsett längd – en 105 m YTS3 gav
  10 p som en punktspricka. `None` stänger av. 213 av 474 skador i DUF 701 är löpande (median 16 m).
- Index = poäng / max(längd, 20 m) × 100 (poäng per 100 m). Konstruktions-, drift- och totalindex.
  **Bara konstruktionsindex styr klass och rangordning** (sep 2026: "att det är rötter medför inte
  att jag vill renovera ledningen") – drift-/totalindex är information (Excel-kolumnordning
  Konstruktions-, Drift-, Totalindex; topplistan visar K-index; PDF-rad Konstruktionsindex först).
- `KODFAKTOR` (KONFIG) viktar konstruktionskoder inbördes: **SPR/RBR/DEF 1,0, YTS 0,7, FOG/FRF 0,6,
  DEA 0,3** (sep 2026, användarens val). Syfte: hitta ledningar att **strumpinfodra medan det går**
  – sprickor/brott/deformation är vägen mot kollaps, ytskada tunnar väggen men röret bär, fogfel
  tätas av strumpan, defekt anslutning åtgärdas ändå med hatt. Simulerat på DUF 701 (A/B/C/D/E):
  alla 1,0 → 58/40/19/62/6; valt förslag → 41/54/22/62/6 (17 A→B, tio av dem en enda löpande YTS3 =
  exakt 100 → 70 p/100 m; 3 B→C); YTS/DEA 0,5 → 34/59/24/62/6; FOG/FRF 0,5 + YTS/DEA 0,3 →
  26/67/24/62/6. Topp 10 oförändrad i alla varianter.
- `GRADFAKTOR` (KONFIG): **YTS grad 4 × 1,0** (sep 2026) – P93: grad 4 = rörväggen genomfrätt, i
  praktiken rörbrott; ingen egen kod/attribut i filen (47 st i DUF 701, 4 med kommentar
  "utläckning"). Ersätter KODFAKTOR för den kombinationen. Effekt: nr 98 och 124 B→A, facit
  40/55/22/62/6, A = 1 501 m, rang 1 859 p/100 m. YTS4 ger fortfarande inte A ensamt
  (GRAD4_KODER_A oförändrad).
- `ATTRIBUTFAKTOR` (KONFIG): **cirkulära sprickor (SPR, CIRK) × 0,7** (sep 2026) – sättning vid fog,
  mindre allvarligt för bärigheten än KOMPL/LÄNGS. DUF 701: 25 sträckor har CIRK (20 A, 5 B);
  0,8 → 39 A (nr 19, 171 A→B), 0,7 och 0,5 → 38 A (även nr 99). Facit blir 38/57/22/62/6, A = 1 391 m.
  Idé som nämnts: flagga "risk för schakt" vid grad 4 RBR/DEF/FOG.
- **Manuellt facit på gång:** användaren går igenom alla A- och B-sträckor i DUF 701 för hand och
  lämnar sin bedömning som facit att kalibrera modellen mot. Excel-fliken Prioritering har därför
  tomma kolumner **"Manuell bedömning"** och **"Kommentar"** (före Rapport/Videofil).
- Klass **A** Åtgärda (hette "Åtgärd snarast" t.o.m. sep 2026): grad 4 på **RBR eller DEF**
  (`GRAD4_KODER_A`, sep 2026 – tidigare alla konstruktionskoder; grad 4-YTS/FRF gav då 4 sträckor A
  med en enda punktskada på 40–80 m frisk ledning) eller konstruktionsindex ≥ 80.
  **B** Planera renovering: grad 3 eller index ≥ 25. **C** Bevaka: övriga med skador.
  **D** Inga skador. **E** Ej bedömd (längd < 1 m). Rangordning inom klass efter konstruktionsindex
  (t.o.m. sep 2026 totalindex).
- Driftåtgärd flaggas separat (rotskärning, spolning, täta inläckage, ta bort hinder) – ingår
  inte i klassen. Även flaggor: avbruten inspektion, relinad, inspekterad flera ggr.
- Facit DUF 701 (längdviktning + GRAD4_KODER_A + KODFAKTOR + GRADFAKTOR + ATTRIBUTFAKTOR): klasser
  A/B/C/D/E = **40/55/22/62/6**; A = 1 501 m (20 %); utan YTS4-vikt 38/57/22/62/6, A = 1 391 m; utan
  CIRK-vikt 41/54/22/62/6, A = 1 481 m; utan KODFAKTOR 58/40/19/62/6, A = 2 199 m; med grad 4 på alla koder dessutom:
  62/36/19/62/6, A = 2 444 m (nr 41, 44, 118, 123 var A via en enda YTS4/FRF4);
  474 räknade skador; rang 1 = SRB64009 → SRB1016560 (betong 225, 35,1 m, 9×YTS4 + SPR3 +
  löpande YTS3 35 m × 3,5: 9×30 (YTS4 fullt) + 10×0,7 (CIRK) + 10×0,7×3,5 = 301,5 p → 859 p/100 m;
  utan YTS4-vikt 628,6; utan KODFAKTOR 315 p → 898); 12 avbrutna inspektioner; 16 relinade;
  YTS 153 (44 grad 4), SPR 138, ROT 110. Utan viktning (t.o.m. sep 2026): 41/53/23/62/6, A = 1 583 m,
  rang 1 = 290 p → 826,5; 25 sträckor byter klass (21 B→A, 4 C→B), nr 108 (105 m YTS3) stannar i B
  (74,9 p/100 m) tack vare taket.
- Profilanalys (viktig lärdom): avvikelse från rät linje mellan brunnarna var **fel mått** –
  det flaggade lutningsbrott som svackor. Nu: **svackdjup** = största stående vattendjup,
  vattenytan i varje punkt = högsta punkten nedströms (vattnet kan bara lämna nedströms – vid
  bakfall räknas djupet upp till utloppet; den tidigare symmetriska "fill"-metoden med
  min(uppströms, nedströms) cappade vid inloppet, fel enligt användaren sep 2026; beräknat djup
  > `SVACKA_MAX_M` (1 m) = driftande inklinometer → svacka None/"okänd", profil osäker; **redovisas
  i cm** i Excel, PDF, profilbild, JSON `svackdjup_cm` och kartfältet `SVACKA_CM` – beräknas i m),
  **svacklängd** (stående vatten > 1 cm), **bakfall** (längd med lutning mot flödet > 5 ‰),
  **svackdjup/diameter**, samt **profil osäker** när inklinometerns fall avviker > 0,3 m eller
  50 % från brunnshöjderna (26 av 181 profiler i DUF 701 – inklinometrar driftar).
  Lutning = (starthöjd − sluthöjd)/längd i ‰, positiv = fall i flödesriktningen.
  Sträckor med störst svacka i DUF 701: KRB68925 → KNBL62753 (plast 200, 0,12 m = 60 % av
  diametern) och KTB1003603 → KRB1000744 (plast 315, 0,10 m, 35 m stående vatten).

## 5. PDF-rapporten per sträcka (reportlab + matplotlib), önskemål som är implementerade

- A4, sidhuvud "Inspektionsprotokoll (TV-inspektion, P93)" + projekt/område, sidfot med fil,
  sträcka och sida. Typsnitt DejaVu Sans / Segoe UI / Arial om hittat (för åäö och →).
- Rubrik med brunn → brunn (**ingen klassruta**). Infotabell: område, datum + väder, start-/slutbrunn
  (uppströms/nedströms), kamera från (position 0), riktning, längd, ledningstyp, material,
  dimension/form, antal anslutningar, antal skador, konstruktionsindex, avbruten inspektion,
  driftindex/totalindex, svacka (djup/längd), lutning (+ bakfall, + "profil osäker"),
  videofil, TV3-fil, littera rättat (bara om rättat).
  **Borttaget på begäran:** projekt, ägare, operatör, driftåtgärd, **prioritetsklass/rekommendation**
  (sep 2026 – finns bara i Excel och kartunderlaget; klassbokstaven sitter kvar i PDF-filnamnet).
- Schematisk översikt: horisontellt rör, brunnar i ändarna (större diameter än röret, okt 2026),
  skador som romber färgade efter grad (grön/gul/orange/röd), löpande skador som band ovanför
  (radpackning `_packa_band`: första lediga raden, etikettbredden räknas med; figurhöjden växer med
  fler än tre rader och `rita_schema` returnerar (bredd, höjd) i tum som PDF:en använder – tidigare
  rad = löpnummer mod 3 så band fyra hamnade på band ett, okt 2026),
  anslutningar som trianglar ovanför (vänster) / under (höger) röret med etikett "15.6 m kl 9",
  meterskala som egen linje under anslutningarna med siffrorna direkt under ticksen och riktningspil
  under skalan (okt 2026, användarens begäran); förklaringen högst upp. Layoutgenomgång okt 2026
  ("fixa efter eget omdöme"): anslutningsetiketter packas per sida i rader (`_packa_band` på
  etikettbredden, inte varannan rad), punktskadornas romber packas i tre rader innanför röret
  (0, +0,15, −0,15) när de ligger inom samma bredd, banden börjar ovanför anslutningsetiketterna och
  skalan läggs under alla etiketter (figurhöjden följer med; tomt band-utrymme borta), avbruten
  inspektion = rött kryss med "inspektion avbruten" (packas med hjässa/botten-etiketter) och slut-
  brunnen ritas ihålig med "(ej nådd)" (`ofullstandig`). Sektionsordning i PDF:en: info, Översikt,
  **Profil**, Observationer, foton (profilen före tabellen så båda graferna oftast får plats på sida 1;
  tabellen delas bra över sidor, bilder inte). Observationstabellen: Pos 15 mm, Tid 17, Kod 12, Kl 13,
  Nivå 11, Foto 32, Grad 12, Poäng 14 ("110.04", "KAM", "Poäng", "Nivå" bröts i två rader);
  poäng med högst en decimal. **Decimalkomma i hela PDF:en** (okt 2026, användarens val): `dk()` byter
  punkt mellan två siffror mot komma och används på infotabellens värden, översiktens etiketter,
  profilens texter och axlar (FuncFormatter), observationstabellen och bildtexterna; kontrollerat med
  regex över alla 185 rapporter. Excel/JSON/konsol oförändrade. **Designgranskning okt 2026** (rättat):
  observationstabellen delas aldrig med färre än `TABELL_MIN_RADER` 3 rader på någon sida (`_Tabell`
  överlagrar `Table.split`, flyttar delningen uppåt; kort tabell hel till nästa sida) och rubriken
  följer med tabellens första rader (`_RubrikTabell`, KeepTogether-variant som bara flyttar när rubrik +
  3 rader inte får plats – reportlabs KeepTogether flyttade hela tabellen); cellmarginal 3 pt och
  kolumnbredder efter uppmätt textbredd (Pos 13, Tid 15, Kod 10, Kl 12, Nivå 10, Foto 32 med 7 pt och ett
  filnamn per rad – 13-siffriga namn bröts mitt i, Grad 10, Poäng 12, Observation resten ≈ 66 mm så
  "Anslutning (påstick/grenrör), vänster" ryms på en rad); profil: OBS-texten nere till vänster (låg
  över brunnsnamnet nere till höger), svacketiketten ovanför när svackan ligger inom 15 % från en ände,
  `HOJDSKALOR` utökad med 500/1000 (nr 23 med driftande inklinometer ritades utanför diagrammet),
  `_dk_axlar` ger alla ticks lika många decimaler (23,0 i stället för 23 bredvid 22,8); attributet
  HINDE heter "hinder" (gav "Kamera/inspektion avbruten, hinder – inspektion avbruten"). Användaren
  ska återkomma med hur anslutningarna ska visas.
- Observationstabell med radfärg efter grad (ingen förklaringstext – borttagen på begäran).
  Bildnamnen i kolumnen Foto är **interna PDF-länkar** (`<a href="#foto_…">`) till fotografiet
  längre bak; ankaret (`<a name>` i en 1 pt-paragraf) ligger ovanför bilden så den hamnar i vy.
  Bara bilder som hittats länkas; första förekomsten av ett filnamn bär ankaret.
- Profil: exakt **höjd- och längdskala** väljs ur fasta serier (1:1, 1:2, 1:5, 1:10 … resp.
  1:10, 1:20, 1:25, 1:40, 1:50, 1:75, 1:100 …) så att kurvan fyller diagrammet; skalorna skrivs i
  rubriken. **Höga änden alltid till vänster** (profilen speglas vid behov; x-axeln anger
  nollbrunn). Utjämnad linje (±0,3 m) för uppritning. Svacka markerad med röd stapel.
  Ingen hjälptext om flödesriktning (borttagen på begäran).
- Fotografier: alla bilder (A och B) som hittas i mediamapparna, två per rad, med bildtext
  filnamn · position · tid · beskrivning. Ingen lista över saknade bilder (borttagen på begäran).
- Länk till rapporten finns i Excel-fliken Prioritering (kolumn Rapport, relativ sökväg).
- Alla 185 rapporter för DUF 701 tar ca 1 minut att generera.

## 6. Stil och arbetssätt användaren vill ha

- Svenska överallt. Korta, konkreta svar; visa exempel (PDF/bild) när något ändras i layouten.
- Rendera och titta på PDF/diagram innan leverans (pdftoppm → bild) – flera fel hittades så.
- Hjälptexter i rapporter är oönskade ("det kan vi se själva").
- Dokumentation som Word (.docx), inte Markdown. Filnamn utan åäö i zip.
- Behåll ett enda skript med KONFIG-block överst; parametrar ska gå att ändra utan att röra
  koden.

**Planerade verktyg (okt 2026):** `PLAN_verktyg.md` i repots rot – svackkarta, uppströmsanalys,
åtgärdspaket/kostnad, PowerPoint – med användarens beslut F1–F7 (rensbrunn = tillsynsbrunn,
lagning manuell i meter kr/m, ingen automatisk brytpunkt mot schakt, schablonpriser, överbryggning
60 m, flödesriktning ur riktningsattribut, bara strumpa/schakt; **F8: ingen kostnad för schakt** –
beror på djup/spont, kalkyleras separat; schaktsträckor listas med längd/dimension utan kronor; **F9: framschaktning + ny brunn per brunn på etappnivå** – gemensam mittenbrunn mellan två TB-sträckor schaktas fram en gång, girig minsta täckning, valet redovisas och kan låsas). **`mall.pptx`** (utkast, repots rot,
byggd med pptxgenjs i scratchpad/pptx/mall.js, tema "tv3_analys": dk2/accent1 0B5C6B teal, lt2 E6EFF2,
accent2–5 klassfärgerna, Calibri, 16:9 wide) har layouterna **TITEL** (title, undertitel, meta),
**AVDELARE** (title, body), **RUBRIK** (bara title – för fritt komponerade sidor som nyckeltalsrutor),
**RUBRIK_TEXT** (title, body), **RUBRIK_BILD** (title, bild 8,2×5,2", body höger), **RUBRIK_TABELL**
(title, tabell), **STRACKA** (title, oversikt 7,9×2,45", foto1–foto3 2,55", fakta höger), **AVSLUT**
(title, body, mörk). Sidfot och sidnummer ligger i layouterna. `tv3_pptx.py` ska fylla via
`placeholder`-namnen. Exempelsidor med DUF 701-innehåll ingår. LibreOffice i sandlådan kan inte
rendera ("source file could not be loaded"), så mallen är kontrollerad med validate.py + python-pptx
(positioner), inte visuellt.

## 7. Idéer som nämnts men inte byggts

- Kartvy: grundversionen finns (`arcmap/`), inkl. hyperlänk till PDF/film. Kvar: lägga
  `.lyr`-filen i repot när användaren sparat den från ArcMap.
- Jämförelse mellan två inspektioner av samma sträcka.
- Stöd för P111-koder som alternativ kodtabell.
- Kostnadsuppskattning per sträcka (kr/m per metod).
- Konsekvensfaktor (ledningens betydelse) som separat dimension – klassen beskriver bara
  tillstånd, inte risk.
- Sträckor med "profil osäker": linjär korrigering mot GIS-brunnshöjder finns nu via markprofilen
  (hojdanpassning); utan GIS-nivåer flaggas fortfarande bara.
- Markprofil ur raster (höjdmodell) i stället för punktlager; locknivåer från brunnslagret;
  datumtransformation vid GeoJSON-export för RT90-data.
