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
Övriga filer: `filer.txt` (exempel-listfil), `Anvandarhandledning tv3_analys.docx`,
`Metodbeskrivning prioritering avloppsledningar.docx` (genereras av `make_docs.js` med npm-paketet
`docx`; filnamn hålls ASCII eftersom Windows zip-hantering förvanskar åäö).

Körning:
```
python tv3_analys.py -l filer.txt [-o utdata] [--topp 15] [--rapporter alla|AB|A|inga] [--diagram] [--media KATALOG]
python tv3_analys.py "testdata/DUF 701.TV3"   # enstaka fil, jokertecken eller katalog fungerar också
```

Listfilens format (relativa sökvägar tolkas relativt listfilen, `#` = kommentar):
```
media: D:\Inspektioner\Filmer          # mediamapp för alla filer i listan (kan upprepas)
DUF 701.TV3                            # media söks alltid även i TV3-filens egen mapp (rekursivt)
DUF 702.TV3 ; D:\Filmer\DUF702 ; E:\Bilder   # egna mediamappar för just den filen
C:\Inspektioner\2022\                  # katalog: alla .TV3 i den
```

Utdata i `tv3_resultat/` (eller `-o`):
- `prioritering.xlsx` – flikar **Sammanfattning** (nyckeltal, klassfördelning, metodparametrar,
  inbäddade diagram), **Prioritering** (en rad per sträcka, rankad; klassceller färgade;
  hyperlänkar till rapport-PDF (relativ länk) och videofil (absolut `file:///`-länk)),
  **Observationer** (alla observationer i klartext, länk till bild och video),
  **Kodstatistik**, **Per fil**, **Material** (relinade sträckor redovisas som eget material
  `Relinad` med ursprungsmaterialet i egen kolumn – styrs av `RELINAD_SOM_MATERIAL` i KONFIG;
  samma indelning i diagram 4, medan fliken Prioritering visar ursprungsmaterial + flaggan Relinad).
- `diagram/1_prioritetsklasser.png`, `2_topplista.png`, `3_observationer_per_kod.png`,
  `4_klass_per_material.png` (200 dpi, för PowerPoint) – **bara med `--diagram`**; annars ritas de
  i en temporär katalog och bäddas enbart in i fliken Sammanfattning (KONFIG: `SPARA_DIAGRAM`).
- `rapporter/<tv3>_<nr>_<klass>_<startbrunn>-<slutbrunn>.pdf` – inspektionsprotokoll per sträcka.
- `kartunderlag.json` – en post per sträcka (brunnspar, maskinell bedömning, index, material,
  flaggor) för ArcMap-skriptet. Stängs av med `--karta nej` (KONFIG: `SKRIV_KARTUNDERLAG`).
- `fel.txt` – bara om TV3-filer eller mediamappar saknades.
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
andra sökta brunnar emellan, högst 8 bitar) – så hittas sträckor uppdelade i flera
ledningsobjekt (fältet `ANT_DELAR`) och brunnar utan egen vertex. Geometrin orienteras
startbrunn→slutbrunn. Första körningen med gamla vertex-metoden (steg2-skriptets) gav 136/180;
CSV:n över omatchade har kolumner med avstånd brunn→närmaste ledning samt `diagnos`
(`Natverk.diagnos`: "hoj max hopp till N", "glapp X m vid (x, y)", "annat lager", "brunnen … finns
inte i brunnslagren"). Grafmetoden gav 159/180 på DUF 701; resten är 13 littera som saknas i kartan
(`AG`/`STBEXTRA` är platshållare) och 3 par utan väg.
Fält: `MASK_BED` (alias "Maskinell bedömning", klass A–E från modellen), `MAN_BED`
("Manuell bedömning", fylls i för hand), samt de härledda `BEDOMNING` (manuell om ifylld,
annars maskinell), `BED_TYP` (Maskinell/Manuell) och `STIL` (`A - Maskinell`). Symbologin
(Unique values på `STIL`, tio kategorier: färg efter klass, **streckad** = maskinell,
**heldragen** = manuell) byggs av `arcmap/skapa_lyr.py` / verktyget **Skapa symbologi (.lyr)**
via ArcObjects (comtypes 1.1.14 ligger vendorerat i `arcmap/lib/`, MIT – användaren kunde inte
pip-installera; `gen/` är gitignorerad) och sparas som `arcmap/bedomda_ledningar.lyr`,
som verktyget använder som standard. Manuell väg (Unique values + Save As Layer File) finns
kvar i handledningen. `.lyr` är binär och kan inte skapas utanför ArcMap. "Uppdatera bedömning"
(eller `BARA_UPPDATERA = True`) räknar bara om de härledda fälten efter manuell ifyllnad; vid
full omkörning bevaras manuella bedömningar per brunnspar. Omatchade par listas i CSV.
Lagervalen i verktyget är **rullistor med kartans lagernamn (GPString)**, inte
GPFeatureLayer: ArcMap tolkar `/` i lagernamn (`A Rensbrunn/tillsynsbrunn`) som sökväg när namnet
skickas som text → "does not exist". Namnen slås upp till lagerobjekt via `arcpy.mapping`
(`hitta_lager`, matchar kort eller långt namn `Grupp\Lager`) och `kalla()` ger lagrets
`dataSource` + definitionsfråga till `MakeFeatureLayer`, så namnet aldrig går som text till GP.
Brunnar i SVOA-kartan: nedstigningsbrunnar (xNB/xNBL) i `A Nedstign och övriga brunnar`,
rens-/tillsynsbrunnar (xRB/xTB) i `A Rensbrunn/tillsynsbrunn` – båda behövs (standard).
Testas utan ArcMap med en låtsas-arcpy (se sessionshistorik) – arcpy-körningen i sig är oprövad.

## 3. TV3-formatet (Svenskt Vatten TV-fil v3.0, P93-koder) – det vi lärt oss

- Textfil, CRLF, **Windows-1252/ISO-8859-1** (åäö i koder som `LÄNGS`, `TVÄRS`, `UTFÄL`).
  Läs som bytes; prova UTF-8, fall tillbaka till cp1252. Decimalkomma överallt.
- Sektioner: `#TVADM` (en rad/sträcka), `#TVDAT` (en rad/observation), `#PROFILADM`,
  `#PROFILDAT` (inklinometer, kan vara >100 000 rader), `#SLUT`. Fält separeras med `;`.
- TVADM-index (0-baserat): 0 sträcknr, 2 startbrunn (uppströms), 3 slutbrunn (nedströms),
  4 **utgångsbrunn (där kameran startade – position 0 m räknas härifrån!)**, 7 ägare, 8 område,
  9 projekt, 10 riktning (Medströms/Motströms), 12 datum, 13 tid, 14 operatör, 18 videofil,
  22 form, 23 dimension, 24 dimension 2, 25 material, 26 foder, 27 fodermaterial,
  29 ledningstyp, 31 väder. 38 kolumner i exempelfilen.
- TVDAT-index: 0 sträcknr, 1 läge (m), 2 tid i film, 3 löpande-markering (`A1`…/`B1`…),
  4 skadekod, 5 grad 1–4, 6 infokod, 7 attribut, 11 klocka från, 12 klocka till,
  13 vattennivå %, 14 bild A, 15 bild B, 16 videofil (första raden), 17 kommentar.
  Sträckans längd = största läge. Rader med `B` är slutmarkering för löpande skada och ska inte
  räknas igen.
- PROFILADM: index 16/17 = start-/sluthöjd vid **start-/slutbrunn** (flödesriktning), inte vid
  utgångsbrunnen. PROFILDAT: 1 läge, 2 relativ höjd, 3 absolut höjd; **kan ligga i fallande
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

- Grad → poäng: 1 → 1, 2 → 3, 3 → 10, 4 → 30. Driftskador × 0,5. Löpande skador räknas en gång.
- Index = poäng / max(längd, 20 m) × 100 (poäng per 100 m). Konstruktions-, drift- och totalindex.
- Klass **A** Åtgärd snarast: konstruktionsgrad 4 eller konstruktionsindex ≥ 80.
  **B** Planera renovering: grad 3 eller index ≥ 25. **C** Bevaka: övriga med skador.
  **D** Inga skador. **E** Ej bedömd (längd < 1 m). Rangordning inom klass efter totalindex.
- Driftåtgärd flaggas separat (rotskärning, spolning, täta inläckage, ta bort hinder) – ingår
  inte i klassen. Även flaggor: avbruten inspektion, relinad, inspekterad flera ggr.
- Facit DUF 701: klasser A/B/C/D/E = 41/53/23/62/6; A = 1 583 m (21 %); 474 räknade skador;
  rang 1 = SRB64009 → SRB1016560 (betong 225, 35,1 m, 9×YTS4 + SPR3 + löpande YTS3 = 290 p →
  826,5 p/100 m); 12 avbrutna inspektioner; 16 relinade; YTS 153 (44 grad 4), SPR 138, ROT 110.
- Profilanalys (viktig lärdom): avvikelse från rät linje mellan brunnarna var **fel mått** –
  det flaggade lutningsbrott som svackor. Nu: **svackdjup** = största stående vattendjup
  ("fill"-metod i flödesriktningen: punkt lägre än både uppströms- och nedströmskant),
  **svacklängd** (stående vatten > 1 cm), **bakfall** (längd med lutning mot flödet > 5 ‰),
  **svackdjup/diameter**, samt **profil osäker** när inklinometerns fall avviker > 0,3 m eller
  50 % från brunnshöjderna (26 av 181 profiler i DUF 701 – inklinometrar driftar).
  Lutning = (starthöjd − sluthöjd)/längd i ‰, positiv = fall i flödesriktningen.
  Sträckor med störst svacka i DUF 701: KRB68925 → KNBL62753 (plast 200, 0,12 m = 60 % av
  diametern) och KTB1003603 → KRB1000744 (plast 315, 0,10 m, 35 m stående vatten).

## 5. PDF-rapporten per sträcka (reportlab + matplotlib), önskemål som är implementerade

- A4, sidhuvud "Inspektionsprotokoll (TV-inspektion, P93)" + projekt/område, sidfot med fil,
  sträcka och sida. Typsnitt DejaVu Sans / Segoe UI / Arial om hittat (för åäö och →).
- Rubrik med brunn → brunn och färgad klassruta. Infotabell: område, datum, start-/slutbrunn
  (uppströms/nedströms), kamera från (position 0), riktning, längd, ledningstyp, material,
  dimension/form, antal anslutningar, antal skador, klass, index, avbruten inspektion, väder,
  svacka (djup/längd), lutning (+ bakfall, + "profil osäker"), videofil, TV3-fil.
  **Borttaget på begäran:** projekt, ägare, operatör, driftåtgärd.
- Schematisk översikt: horisontellt rör, brunnar i ändarna, skador som romber färgade efter grad
  (grön/gul/orange/röd), löpande skador som band ovanför, anslutningar som trianglar ovanför
  (vänster) / under (höger) röret med etikett "15.6 m kl 9", meterskala, riktningspil.
- Observationstabell med radfärg efter grad (ingen förklaringstext – borttagen på begäran).
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

- Kartvy: grundversionen finns (`arcmap/`). Kvar: koppla rapport-PDF och film som hyperlänk i
  kartan; lägga den genererade `.lyr`-filen i repot när den skapats på användarens dator.
- Jämförelse mellan två inspektioner av samma sträcka.
- Stöd för P111-koder som alternativ kodtabell.
- Kostnadsuppskattning per sträcka (kr/m per metod).
- Konsekvensfaktor (ledningens betydelse) som separat dimension – klassen beskriver bara
  tillstånd, inte risk.
- Sträckor med "profil osäker": eventuellt möjlighet att korrigera inklinometern linjärt mot
  brunnshöjderna i stället för att bara flagga.
