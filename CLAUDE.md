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
`globala = {"media", "bild", "littera", "markprofil"}`; `Stracka.media_kataloger`/`bild_kataloger`;
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
koppla_markprofil). Utan markprofil är allt oförändrat (regressionstestat). Testad med syntetisk
markprofil för DUF 701 (scratchpad) och låtsas-arcpy för verktyget.
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
- Samma brunnspar kan förekomma två gånger (filmat från båda håll efter avbrott) – flaggas,
  slås inte ihop automatiskt.

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
- Schematisk översikt: horisontellt rör, brunnar i ändarna, skador som romber färgade efter grad
  (grön/gul/orange/röd), löpande skador som band ovanför, anslutningar som trianglar ovanför
  (vänster) / under (höger) röret med etikett "15.6 m kl 9", meterskala, riktningspil.
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
