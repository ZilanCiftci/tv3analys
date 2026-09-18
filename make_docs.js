// Skapar Användarhandledning.docx och Metodbeskrivning.docx för tv3_analys
const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, Table, TableRow, TableCell,
  WidthType, ShadingType, BorderStyle, AlignmentType, LevelFormat, PageNumber, Footer, Header,
} = require("docx");

const FONT = "Calibri";
const BLUE = "1F3864";

const numbering = {
  config: [
    { reference: "bullets", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT,
      style: { paragraph: { indent: { left: 560, hanging: 280 } } } }] },
    { reference: "steps", levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT,
      style: { paragraph: { indent: { left: 560, hanging: 280 } } } }] },
  ],
};

const styles = {
  default: { document: { run: { font: FONT, size: 22 } } },
  paragraphStyles: [
    { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
      run: { size: 32, bold: true, color: BLUE, font: FONT }, paragraph: { spacing: { before: 360, after: 160 }, outlineLevel: 0 } },
    { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
      run: { size: 26, bold: true, color: BLUE, font: FONT }, paragraph: { spacing: { before: 240, after: 120 }, outlineLevel: 1 } },
    { id: "Title", name: "Title", basedOn: "Normal", next: "Normal",
      run: { size: 48, bold: true, color: BLUE, font: FONT }, paragraph: { spacing: { after: 120 } } },
  ],
};

const p = (text, opts = {}) => new Paragraph({ spacing: { after: 120 }, ...opts, children: runs(text) });
function runs(text) {
  // **fet** och `kod` i enkel markdown-stil
  const parts = [];
  const re = /(\*\*[^*]+\*\*|`[^`]+`)/g;
  let last = 0, m;
  while ((m = re.exec(text))) {
    if (m.index > last) parts.push(new TextRun(text.slice(last, m.index)));
    const t = m[0];
    if (t.startsWith("**")) parts.push(new TextRun({ text: t.slice(2, -2), bold: true }));
    else parts.push(new TextRun({ text: t.slice(1, -1), font: "Consolas", size: 20, shading: { type: ShadingType.CLEAR, fill: "F2F2F2" } }));
    last = m.index + t.length;
  }
  if (last < text.length) parts.push(new TextRun(text.slice(last)));
  return parts;
}
const h1 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun(t)] });
const h2 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun(t)] });
const bullet = (t) => new Paragraph({ numbering: { reference: "bullets", level: 0 }, spacing: { after: 80 }, children: runs(t) });
const step = (t) => new Paragraph({ numbering: { reference: "steps", level: 0 }, spacing: { after: 80 }, children: runs(t) });
const code = (lines) => lines.map((l) => new Paragraph({
  spacing: { after: 0 }, shading: { type: ShadingType.CLEAR, fill: "F2F2F2" },
  indent: { left: 280 },
  children: [new TextRun({ text: l, font: "Consolas", size: 19 })],
}));
const spacer = () => new Paragraph({ spacing: { after: 120 }, children: [] });

function table(header, rows, widths) {
  const total = widths.reduce((a, b) => a + b, 0);
  const border = { style: BorderStyle.SINGLE, size: 4, color: "BFBFBF" };
  const borders = { top: border, bottom: border, left: border, right: border };
  const cell = (text, w, isHead, fill) => new TableCell({
    width: { size: w, type: WidthType.DXA }, borders,
    shading: fill ? { type: ShadingType.CLEAR, fill } : undefined,
    margins: { top: 60, bottom: 60, left: 100, right: 100 },
    children: [new Paragraph({ spacing: { after: 0 }, children: isHead
      ? [new TextRun({ text, bold: true, color: "FFFFFF", size: 20 })]
      : runs(text).map((r) => r) })],
  });
  return new Table({
    width: { size: total, type: WidthType.DXA }, columnWidths: widths,
    rows: [
      new TableRow({ tableHeader: true, children: header.map((h, i) => cell(h, widths[i], true, BLUE)) }),
      ...rows.map((r, ri) => new TableRow({ children: r.map((c, i) => cell(String(c), widths[i], false, ri % 2 ? "F7F9FC" : undefined)) })),
    ],
  });
}

function doc(title, subtitle, children) {
  return new Document({
    creator: "tv3_analys", title, styles, numbering,
    sections: [{
      properties: { page: { margin: { top: 1440, bottom: 1440, left: 1440, right: 1440 } } },
      headers: { default: new Header({ children: [new Paragraph({ alignment: AlignmentType.RIGHT,
        children: [new TextRun({ text: title, color: "898781", size: 18 })] })] }) },
      footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER,
        children: [new TextRun({ children: ["Sida ", PageNumber.CURRENT, " av ", PageNumber.TOTAL_PAGES], color: "898781", size: 18 })] })] }) },
      children: [
        new Paragraph({ style: "Title", children: [new TextRun(title)] }),
        new Paragraph({ spacing: { after: 360 }, children: [new TextRun({ text: subtitle, color: "52514E", size: 24 })] }),
        ...children,
      ],
    }],
  });
}

// ---------------------------------------------------------------------------
// 1. Användarhandledning
// ---------------------------------------------------------------------------
const handledning = doc("Användarhandledning – tv3_analys", "Analys av TV-inspektioner av avloppsledningar från TV3-filer", [
  h1("1. Vad skriptet gör"),
  p("`tv3_analys.py` läser en eller flera TV3-filer (Svenskt Vatten TV-fil version 3.0 med P93-koder), poängsätter varje ledningssträcka utifrån de registrerade observationerna och tar fram ett underlag för prioritering av renovering: en Excel-arbetsbok med klickbara länkar till film och bilder, diagram för presentationer och ett inspektionsprotokoll i PDF per sträcka (i stil med WinCan-rapporter) med sträckdata, schematisk översikt, observationslista, inklinometerprofil och fotografier. All bearbetning sker lokalt på din dator – inga filer skickas någonstans."),
  p("Poängmodellen och hur prioritetsklasserna sätts beskrivs i det separata dokumentet **Metodbeskrivning – prioritering av avloppsledningar**."),

  h1("2. Installation"),
  step("Installera Python 3.10 eller senare (python.org, eller via Microsoft Store på Windows)."),
  step("Öppna en terminal (Kommandotolken/PowerShell på Windows) och installera de två biblioteken som behövs:"),
  ...code(["pip install openpyxl matplotlib reportlab"]),
  spacer(),
  step("Lägg `tv3_analys.py` i en lämplig mapp, till exempel samma mapp som TV3-filerna."),

  h1("3. Körning"),
  h2("3.1 Listfil med flera TV3-filer (rekommenderat)"),
  p("Skapa en vanlig textfil, till exempel `filer.txt`. Varje rad anger en TV3-fil (eller en mapp med TV3-filer) och kan dessutom ange var filmerna och bilderna till just den filen ligger:"),
  ...code([
    "# Inspektioner Äppelviken 2021",
    "media: D:\\Inspektioner\\Filmer            # gäller alla filer i listan",
    "",
    "DUF 701.TV3                                # media söks i TV3-filens egen mapp",
    "DUF 702.TV3 ; D:\\Filmer\\DUF702            # egen mediamapp för denna fil",
    "\"C:\\Inspektioner\\2022\\DUF 810.TV3\" ; E:\\Film ; E:\\Bilder",
    "../andra_omradet/                          # mapp: alla TV3-filer i den",
  ]),
  spacer(),
  bullet("Tomma rader och rader som börjar med `#` hoppas över."),
  bullet("Relativa sökvägar tolkas relativt listfilens mapp; absoluta sökvägar fungerar som de är. Sökvägar med mellanslag kan skrivas med eller utan citattecken."),
  bullet("En mapp på en rad expanderas till alla `.TV3`-filer i mappen, inklusive undermappar."),
  bullet("`media:` (eller `film:`) på en egen rad anger en mapp där filmer och bilder söks för **alla** TV3-filer i listan. Kan upprepas."),
  bullet("Allt efter `;` på en TV3-rad är mediamappar för **just den filen**. Flera mappar separeras med `;`."),
  bullet("Mediafiler söks alltid rekursivt (inklusive undermappar), i ordningen: TV3-filens egen mapp, filens egna mediamappar, de gemensamma `media:`-mapparna och sist `--media` från kommandoraden. Första träffen på filnamnet används."),
  p("Kör sedan:"),
  ...code(["python tv3_analys.py -l filer.txt"]),
  spacer(),
  p("Alla filer i listan analyseras **tillsammans** till ett gemensamt underlag. Filer som saknas eller inte går att läsa stoppar inte körningen; de och mediamappar som inte finns listas i `fel.txt` i utdatamappen."),

  h2("3.2 Andra sätt att ange filer"),
  ...code([
    'python tv3_analys.py "DUF 701.TV3"                # en fil',
    "python tv3_analys.py DUF*.TV3 -o resultat --topp 20   # jokertecken",
    "python tv3_analys.py inspektioner/                # hel mapp",
    "python tv3_analys.py filer.txt                    # listfil utan -l",
  ]),
  spacer(),
  h2("3.3 Alternativ"),
  table(["Alternativ", "Betydelse", "Standard"], [
    ["`-l FIL.TXT`, `--lista`", "Listfil med en TV3-sökväg per rad. Kan anges flera gånger.", "–"],
    ["`-o MAPP`, `--utdata`", "Mapp där resultatet skrivs.", "`tv3_resultat`"],
    ["`--topp N`", "Antal sträckor i topplistan (diagram och Sammanfattning-fliken).", "15"],
    ["`--rapporter alla|AB|A|inga`", "Vilka sträckor som får PDF-rapport. `alla` är standard; `AB` bara klass A och B; `inga` hoppar över (snabbare).", "`alla`"],
    ["`--diagram`", "Sparar diagrammen som PNG-filer i `diagram`-mappen, t.ex. för PowerPoint. Utan flaggan bäddas de bara in i Excel-filen.", "av"],
    ["`--karta ja|nej`", "Skriver `kartunderlag.json` för kartframställning i ArcMap (se avsnitt 7).", "`ja`"],
    ["`--media KATALOG`", "Extra mapp att söka video- och bildfiler i (kan anges flera gånger). TV3-filens egen mapp söks alltid, inklusive undermappar.", "–"],
  ], [2600, 5000, 1760]),
  spacer(),
  p("Skriptet skriver en kort rapport i terminalen: vilka filer som lästes, antal sträckor per prioritetsklass och topplistan."),

  h1("4. Resultatfiler"),
  p("Allt hamnar i utdatamappen (`tv3_resultat` om inget annat anges):"),
  table(["Fil", "Innehåll"], [
    ["`prioritering.xlsx`", "Excel-arbetsbok med flikarna Sammanfattning, Prioritering, Observationer, Kodstatistik, Per fil och Material (se avsnitt 5)."],
    ["`diagram/*.png`", "Skrivs bara om du kör med `--diagram`: 1_prioritetsklasser (inspekterad ledningslängd per prioritetsklass), 2_topplista (de N mest kritiska sträckorna, uppdelat på konstruktions- och driftskador), 3_observationer_per_kod (antal skadeobservationer per kod och grad) och 4_klass_per_material (prioritetsklass per material, andel av inspekterad längd)."],
    ["`rapporter/*.pdf`", "Ett inspektionsprotokoll per sträcka, namngivet `<tv3-fil>_<nr>_<klass>_<startbrunn>-<slutbrunn>.pdf`. Innehåller sträckdata (material, dimension, längd, antal anslutningar, prioritetsklass, index m.m.), schematisk översikt med skador och anslutningar, observationstabell färgad efter grad, inklinometerprofil med lutning och svacka, samt de fotografier som hittats i mediamapparna."],
    ["`kartunderlag.json`", "En post per sträcka (brunnspar, bedömning, index, material) som ArcMap-skriptet läser för att skapa ett ledningslager. Stängs av med `--karta nej`."],
    ["`fel.txt`", "Skrivs bara om någon TV3-fil eller mediamapp saknades eller inte kunde läsas."],
  ], [3600, 5760]),
  spacer(),
  p("De två första diagrammen bäddas alltid in i fliken Sammanfattning. Med `--diagram` sparas alla fyra dessutom som PNG i 200 dpi och kan läggas direkt i PowerPoint. Klassfärgerna är desamma i diagram och Excel: rött A, orange B, gult C, grönt D, grått E."),

  h1("5. Flikarna i Excel"),
  table(["Flik", "Innehåll och användning"], [
    ["Sammanfattning", "Nyckeltal för hela underlaget, fördelning per prioritetsklass (antal och längd), poängmodellens parametrar och två inbäddade diagram."],
    ["Prioritering", "En rad per sträcka, rankad efter prioritetsklass och index. Använd autofiltret för att filtrera på klass, material, ledningstyp, område eller fil. Kolumnen Skador (kod+grad) visar t.ex. 9×YTS4, 1×SPR3. Kolumnerna Avbruten inspektion, Inspekterad flera ggr och Relinad ger extra kontext. Ur inklinometerprofilen beräknas Svackdjup (största stående vattendjup i en svacka), Svackdjup/diameter, Svacklängd, Bakfall längd (meter med lutning mot flödesriktningen) och Lutning; Profil osäker markerar sträckor där inklinometerns fall avviker kraftigt från brunnshöjderna och profilen därför inte bör användas. Antal anslutningar räknar registrerade anslutningar (AS/AG) på sträckan. Kolumnen Rapport öppnar sträckans PDF-rapport (länken är relativ, så Excel-filen och mappen rapporter måste ligga kvar bredvid varandra). Kolumnen Videofil är en klickbar länk som öppnar filmen (se avsnitt 7 om länken inte fungerar)."],
    ["Observationer", "Varje observation i klartext med läge (m), tid i filmen, kod, grad, poäng, klockposition, vattennivå, bild och kommentar. Bild och Videofil är klickbara länkar när filerna hittats. Sorterad i samma ordning som prioriteringslistan så att man snabbt hittar detaljerna för en kritisk sträcka."],
    ["Kodstatistik", "Antal observationer per skadekod fördelat på grad 1–4, samt hur många sträckor som berörs."],
    ["Per fil", "Nyckeltal per TV3-fil: projekt, område, period, sträckor, längd, index, klassfördelning och avbrutna inspektioner. Praktiskt när flera uppdrag analyseras samtidigt."],
    ["Material", "Sträckor, längd, index och andel klass A+B per material och ledningstyp. Relinade (infodrade) sträckor redovisas som ett eget material \"Relinad\" med ursprungsmaterialet i egen kolumn, så att de inte räknas in i betong- eller plaststatistiken. Samma indelning används i diagrammet Prioritetsklass per material."],
  ], [2400, 6960]),

  h1("6. Anpassa poängmodellen"),
  p("Alla parametrar ligger samlade överst i `tv3_analys.py` under rubriken **KONFIG** och kan ändras med en vanlig textredigerare utan att röra resten av koden:"),
  table(["Parameter", "Betydelse", "Standard"], [
    ["`GRADPOANG`", "Poäng per grad 1–4.", "1 / 3 / 10 / 30"],
    ["`DRIFTFAKTOR`", "Faktor för driftskador (rötter, inläckage, sediment m.m.).", "0,5"],
    ["`MINLANGD`", "Minsta längd (m) som används som nämnare vid normering per 100 m.", "20"],
    ["`TROSKEL_A`", "Konstruktionsindex (p/100 m) som ger klass A.", "80"],
    ["`TROSKEL_B`", "Konstruktionsindex (p/100 m) som ger klass B.", "25"],
    ["`KODER`", "Skadekoder med klartext och typ (K = konstruktion, D = drift, I = information).", "P93"],
    ["`ATTRIBUT`, `INFOKODER`", "Klartext för attribut (KOMPL, TUNNA …) och koder utan grad (AS, NB …).", "P93"],
  ], [2600, 5000, 1760]),
  spacer(),
  p("Okända koder passerar igenom och visas som de är i Excel, men ger inga poäng förrän de lagts in i `KODER`. Om ni använder en annan standard än P93 (t.ex. P111) läggs koderna till i tabellerna."),

  h1("7. Kartlager i ArcMap"),
  p("Mappen `arcmap` innehåller en verktygslåda för ArcMap 10.x. Lägg till den en gång: högerklicka i ArcToolbox-fönstret, välj **Add Toolbox** och peka på `arcmap/tv3_verktyg.pyt`. Verktygslådan **tv3_analys** får två verktyg:"),
  bullet("**Skapa ledningslager** – välj `kartunderlag.json`, ledningslager, brunnslager, fältet med brunnsbeteckning (standard `EntityID`) och var utdata ska sparas. Lagervalen är rullistor med kartans lager; `A Ledning`, `A Nedstign och övriga brunnar` och `A Rensbrunn/tillsynsbrunn` är förifyllda när de finns. Ta med **alla** lager där brunnar kan ligga – nedstigningsbrunnar och rens-/tillsynsbrunnar ligger i olika lager, och loggen visar hur många av inspektionens brunnar som hittades och vilka brunnstyper som saknas. Valfritt: ett polygonlager som begränsar sökningen, en `.lyr`-fil med symbologi och en CSV-rapport över brunnspar som inte hittades. Under **Matchning** finns tolerans (2 m), max antal brunnar en sträcka får passera (2) och fält från ledningslagret som ska följa med."),
  bullet("**Uppdatera bedömning** – välj lagret när du fyllt i manuella bedömningar, så räknas Gällande bedömning, Bedömningstyp och STIL om utan att geometrin rörs."),
  bullet("**Skapa symbologi (.lyr)** – bygger symbologin (färg per klass, heldragen/streckad) och sparar den som `.lyr`, se 7.2."),
  p("Verktyget läser `kartunderlag.json`, letar upp varje brunnspar i brunnslagret och klipper ut ledningen mellan brunnarna som ett eget objekt. Ledningar som passerar flera brunnar utan att vara uppdelade hanteras: varje vertexpunkt jämförs med närmaste brunn inom toleransen, och max-hopp styr hur många brunnar en sträcka får passera. Utdata i en filgeodatabas rekommenderas; en shapefil fungerar men visar inte fältalias."),
  p("Samma sak går att köra utan dialog i ArcMaps Python-fönster, med inställningarna i KONFIG överst i `skapa_ledningslager.py`:"),
  ...code(["execfile(r'H:\\PY\\tv3analys\\arcmap\\skapa_ledningslager.py')"]),
  spacer(),
  h2("7.1 Bedömningsfälten"),
  table(["Fält", "Alias", "Innehåll"], [
    ["`MASK_BED`", "Maskinell bedömning", "Prioritetsklass A–E från poängmodellen. Skrivs av skriptet."],
    ["`MAN_BED`", "Manuell bedömning", "Tomt från början. Fyll i A–E för hand i ArcMap när du gjort en egen bedömning."],
    ["`BEDOMNING`", "Gällande bedömning", "Manuell bedömning om den är ifylld, annars den maskinella. Härleds av skriptet."],
    ["`BED_TYP`", "Bedömningstyp", "`Maskinell` eller `Manuell` – styr om linjen ritas streckad eller heldragen."],
    ["`STIL`", "Symbologi", "Klass och typ i ett fält, t.ex. `A - Maskinell`. Symbologin bygger på detta fält."],
  ], [1800, 2600, 4960]),
  spacer(),
  p("Fältnamn i ArcGIS får inte innehålla mellanslag eller åäö, därför heter fälten `MASK_BED` och `MAN_BED` med alias **Maskinell bedömning** och **Manuell bedömning**. Alias visas bara om utdata är en filgeodatabas; i en shapefil syns fältnamnen. I en filgeodatabas får `MAN_BED` dessutom en värdelista med klasserna A–E."),
  p("När du fyllt i manuella bedömningar: kör verktyget **Uppdatera bedömning** (eller skriptet med `BARA_UPPDATERA = True`). Då räknas `BEDOMNING`, `BED_TYP` och `STIL` om utan att geometrin byggs om. Vid en full omkörning läses befintliga manuella bedömningar in per brunnspar och återställs automatiskt."),
  h2("7.2 Symbologi"),
  p("Symbologin är färg efter prioritetsklass (rött A, orange B, gult C, grönt D, grått E – samma som i Excel och diagrammen), **heldragen** linje för manuell bedömning och **streckad** för maskinell. Den sparas som `.lyr`-fil, som verktyget Skapa ledningslager applicerar automatiskt om filen ligger som `arcmap/bedomda_ledningar.lyr` eller anges i fältet Symbologi."),
  p("Enklast är verktyget **Skapa symbologi (.lyr)**: välj ledningslagret i kartan och var `.lyr`-filen ska sparas. Verktyget bygger renderaren, sätter den på lagret och sparar filen. Det kräver Python-paketet `comtypes` i ArcMaps Python (en gång, i Kommandotolken):"),
  ...code(["\"C:\\Python27\\ArcGIS10.8\\Scripts\\pip.exe\" install \"comtypes<1.2\""]),
  spacer(),
  p("Byt `ArcGIS10.8` mot din version. Första körningen tar en stund eftersom comtypes genererar kod för ArcObjects; därefter går det på några sekunder."),
  p("Går inte det, sätt symbologin för hand en gång: i lagrets egenskaper **Symbology > Categories > Unique values**, Value Field `STIL`, **Add All Values**. Det ger tio kategorier – `A - Manuell` … `E - Manuell` heldragna och `A - Maskinell` … `E - Maskinell` streckade, i klassens färg. Spara sedan lagret som `.lyr` (högerklicka på lagret > Save As Layer File) på samma plats."),
  h2("7.3 Sträckor som inte hittas"),
  p("Brunnspar som inte gick att matcha mot en ledning skrivs till `omatchade_par.csv` med uppgift om vilka av brunnarna som finns i kartan. Vanliga orsaker: brunnsbeteckningen skiljer sig mellan TV3-filen och databasen, brunnen saknas i kartan, eller att ledningen passerar fler brunnar än `MAX_HOPP` tillåter. Beteckningar jämförs utan mellanslag, bindestreck och understreck, och utan hänsyn till versaler."),

  h1("8. Vanliga frågor"),
  p("**PDF-rapporterna saknar bilder.** Bilderna hämtas från samma mediamappar som länkarna i Excel. Rapporten anger hur många bilder som hittades och listar de som saknas. Positionerna i rapporten räknas från den brunn kameran startade i (Kamera från), precis som i entreprenörens protokoll, medan Startbrunn/Slutbrunn anger uppströms/nedströms."),
  p("**Körningen tar lång tid.** PDF-rapporterna tar ungefär en halv sekund per sträcka. Använd `--rapporter AB` för att bara skapa rapporter för sträckor i klass A och B, eller `--rapporter inga` när du bara vill uppdatera Excel-filen."),
  p("**Åäö blir fel i resultatet.** TV3-filer är normalt sparade i Windows-1252/ISO-8859-1. Skriptet provar UTF-8 först och faller sedan tillbaka till Windows-1252, så det ska fungera automatiskt. Om en fil ändå blir fel: öppna den i Anteckningar och spara om som UTF-8."),
  p("**Samma sträcka finns två gånger i listan.** Det händer när entreprenören filmat från båda hållen, ofta efter ett avbrott. Kolumnen Inspekterad flera ggr markerar dessa; bedöm dem tillsammans."),
  p("**Sträckor i klass E.** Sträckor utan inspekterad längd (ingen film gjord, bara brunnsregistrering). De behöver inspekteras innan de kan bedömas."),
  p("**Länkarna till film och bilder fungerar inte.** TV3-filen innehåller bara filnamnen (t.ex. `1105202111.mp4`), inte var filerna ligger. Skriptet söker igenom mappen där TV3-filen ligger, inklusive undermappar, samt de mappar du anger i listfilen (`media:` eller `; mapp` efter TV3-filen) eller med `--media`. Terminalen visar hur många filmer och bilder som hittades. Om filmerna ligger på en annan disk eller server: lägg till mappen i listfilen och kör om. Länkarna är absoluta sökvägar, så Excel-filen måste öppnas på en dator som når samma mapp (samma nätverksenhet/enhetsbokstav)."),
  p("**Excel varnar när jag klickar på en länk.** Excel visar en säkerhetsfråga första gången en lokal fil öppnas via länk; svara Ja/OK. Filmen öppnas i datorns standardspelare från början – spola till tiden i kolumnen Tid i film."),
]);

// ---------------------------------------------------------------------------
// 2. Metodbeskrivning
// ---------------------------------------------------------------------------
const metod = doc("Metodbeskrivning – prioritering av avloppsledningar", "Poängmodell och prioritetsklasser för TV-inspekterade ledningar (P93)", [
  h1("1. Syfte"),
  p("Metoden rangordnar TV-inspekterade avloppsledningar efter renoveringsbehov på ett sätt som är enkelt att förklara och att reproducera. Den utgår helt från de observationer entreprenören registrerat enligt Svenskt Vattens P93 och kräver inga andra data. Resultatet är en prioritetsklass A–E per sträcka och ett numeriskt index som gör det möjligt att rangordna sträckor inom en klass."),
  p("Metoden är ett **planeringsunderlag**. Den ersätter inte en teknisk bedömning av den enskilda ledningen, men den pekar ut var den bedömningen bör göras först."),

  h1("2. Indata"),
  p("En TV3-fil innehåller fyra delar som metoden använder:"),
  bullet("**TVADM** – en rad per ledningssträcka: brunnar, material, dimension, ledningstyp, datum, riktning, foder (relining) och videofil."),
  bullet("**TVDAT** – en rad per observation: läge i meter, tid i filmen, skadekod, grad 1–4, attribut, klockposition, vattennivå, bild och kommentar."),
  bullet("**PROFILADM / PROFILDAT** – inklinometerdata (om sådan finns), som används för att beräkna lutning och svacka."),
  p("Sträckans längd sätts till det största lägesvärdet bland dess observationer, det vill säga den faktiskt inspekterade längden."),

  h1("3. Observationskoder"),
  p("Koderna delas in i tre typer. Konstruktionsskador påverkar ledningens bärförmåga och täthet och styr prioritetsklassen. Driftskador påverkar funktionen och kan ofta åtgärdas med underhåll (spolning, rotskärning). Informationskoder ger inga poäng."),
  table(["Kod", "Klartext", "Typ"], [
    ["SPR", "Sprickor (komplexa, cirkulära, längsgående)", "Konstruktion"],
    ["RBR", "Rörbrott", "Konstruktion"],
    ["DEF", "Deformation", "Konstruktion"],
    ["YTS", "Ytskada", "Konstruktion"],
    ["FOG", "Fogförskjutning (riktnings-, tvärs-, längsförskjutning)", "Konstruktion"],
    ["FRF", "Fog, rörfel", "Konstruktion"],
    ["DEA", "Defekt anslutning", "Konstruktion"],
    ["ROT", "Rötter (tunna, grova, rotpaket)", "Drift"],
    ["INL", "Inläckage", "Drift"],
    ["UTF", "Utfällning", "Drift"],
    ["SED", "Sediment", "Drift"],
    ["INH", "Inträngande hinder (t.ex. inträngande servis)", "Drift"],
    ["KAM", "Kamera/inspektion avbruten", "Information"],
  ], [1400, 5560, 2400]),
  spacer(),
  p("Koder utan grad (AS anslutning, RB/NB/TB brunnar, RP reparation, BR böj, DF dimensionsförändring, MF materialförändring, ST stalp) beskriver ledningen och ger inga poäng. Klartexterna för FRF, DEA och ST är tolkade utifrån inspektionsfilerna och bör stämmas av mot ledningsägarens egen kodlista."),

  h1("4. Poängsättning"),
  h2("4.1 Poäng per observation"),
  p("Varje skadeobservation ges poäng efter grad. Skalan är avsiktligt brant så att en allvarlig skada väger tyngre än många lindriga:"),
  table(["Grad", "Poäng", "Innebörd (P93)"], [
    ["1", "1", "Obetydlig skada"],
    ["2", "3", "Mindre skada"],
    ["3", "10", "Allvarlig skada"],
    ["4", "30", "Mycket allvarlig skada"],
  ], [1400, 1400, 6560]),
  spacer(),
  p("Driftskador multipliceras med faktor **0,5**. Löpande skador (markerade A1 … B1 i filen) räknas **en gång**, vid startmarkeringen; längden på den löpande skadan redovisas separat."),
  h2("4.2 Index per sträcka"),
  p("Poängen summeras per sträcka, separat för konstruktion och drift, och normeras till **poäng per 100 m**:"),
  ...code(["index = summa poäng / max(inspekterad längd, 20 m) × 100"]),
  spacer(),
  p("Nämnaren är aldrig kortare än 20 m. Utan den regeln skulle en enstaka skada på en 5 m lång sträcka ge ett orimligt högt index jämfört med samma skada på en 50 m lång sträcka."),
  p("Tre index redovisas: konstruktionsindex, driftindex och totalindex (summan)."),

  h1("5. Prioritetsklass"),
  p("Klassen bestäms av konstruktionsskadorna – den värsta graden på sträckan och konstruktionsindex:"),
  table(["Klass", "Regel", "Tolkning"], [
    ["A – Åtgärd snarast", "Konstruktionsgrad 4 på sträckan, eller konstruktionsindex ≥ 80 p/100 m", "Renovering bör planeras in omgående; teknisk bedömning av metod."],
    ["B – Planera renovering", "Konstruktionsgrad 3, eller konstruktionsindex ≥ 25 p/100 m", "Tas med i den fleråriga förnyelseplanen."],
    ["C – Bevaka", "Övriga sträckor med registrerade skador", "Ingen åtgärd nu; följ upp vid nästa inspektion."],
    ["D – Inga skador", "Inga skadeobservationer", "–"],
    ["E – Ej bedömd", "Ingen inspekterad längd (< 1 m)", "Behöver inspekteras."],
  ], [2400, 3760, 3200]),
  spacer(),
  p("Inom varje klass rangordnas sträckorna efter totalindex, därefter efter högsta konstruktionsgrad. Det innebär att sammanhängande stråk med många skador hamnar överst, vilket ofta också är de sträckor där en samordnad renovering (t.ex. relining av flera sträckor i följd) ger mest nytta."),

  h1("6. Driftåtgärder"),
  p("Driftskador påverkar inte prioritetsklassen men flaggas separat, eftersom de kräver åtgärd oavsett om ledningen ska renoveras:"),
  table(["Observation", "Flagga"], [
    ["Rötter grad 3–4", "Rotskärning"],
    ["Rötter grad 2", "Bevaka rötter"],
    ["Sediment eller utfällning grad 3–4", "Spolning"],
    ["Inläckage grad 3–4", "Täta inläckage"],
    ["Inträngande hinder grad 3–4", "Ta bort inträngande hinder"],
  ], [4680, 4680]),

  h1("7. Kompletterande uppgifter per sträcka"),
  bullet("**Avbruten inspektion** – sträckan har koden KAM med attribut HINDE; kameran kom inte fram (inträngande servis, sediment, rötter, stalp). Den inspekterade längden är då kortare än sträckan och bedömningen ofullständig. Kandidat för kompletterande inspektion från andra hållet."),
  bullet("**Inspekterad flera gånger** – samma brunnspar förekommer flera gånger i underlaget, oftast med- och motströms efter ett avbrott. Bedömningarna bör slås ihop manuellt."),
  bullet("**Relinad** – ledningen har foder enligt TVADM eller kommentaren 'relinad'. Skador på fodret bedöms på samma sätt, men åtgärden är en annan."),
  bullet("**Svackor, bakfall och lutning** – beräknas ur inklinometerprofilen när sådan finns. Svackdjupet är det största vattendjup som blir stående i en svacka, det vill säga hur mycket lägre en punkt ligger än både sin uppströms- och nedströmskant (så kallad fill-metod i flödesriktningen). Svacklängd är den sträcka där stående vatten överstiger 1 cm och bakfall den sammanlagda längden där ledningen lutar mot flödesriktningen (mer än 5 ‰). Svackdjupet sätts också i relation till rördiametern. Stora svackor ger sediment- och kapacitetsproblem och kan motivera åtgärd även utan konstruktionsskador. Inklinometerprofiler kan drifta; om inklinometerns fall avviker mer än 0,3 m eller 50 % från brunnshöjderna markeras profilen som osäker."),

  h1("8. Begränsningar"),
  bullet("Metoden bygger på entreprenörens kodning. Olika operatörer graderar olika; jämförelser mellan uppdrag bör göras med det i åtanke."),
  bullet("Konsekvens vid fel (ledningens betydelse, trafiklast, närhet till vattendrag, dimension, ålder) ingår inte. Prioritetsklassen beskriver **tillstånd**, inte risk. Vid budgetprioritering bör konsekvens läggas till som en separat faktor."),
  bullet("Grad 4-ytskada räcker ensamt för klass A. Det är rimligt för betong (armering kan vara synlig) men kan behöva justeras för andra material."),
  bullet("Alla parametrar – poäng per grad, driftfaktor, minsta längd och trösklar – kan ändras i skriptets konfigurationsdel. Ändringar bör dokumenteras så att resultat från olika tillfällen förblir jämförbara."),

  h1("9. Exempel"),
  p("Sträckan SRB64009 → SRB1016560 (betong 225 mm, 35,1 m) har nio ytskador grad 4, en komplex spricka grad 3 och en löpande ytskada grad 3. Poäng: 9 × 30 + 10 + 10 = 290. Konstruktionsindex: 290 / 35,1 × 100 = 826 p/100 m. Klass A (grad 4 finns), rang 1 i uppdraget DUF 701."),
  p("Sträckan KRB68713 → KRB68712 (betong 225 mm, 9,7 m) har en spricka grad 3, en ytskada grad 3 och ett inträngande hinder grad 2. Poäng: 10 + 10 + 3 × 0,5 = 21,5. Eftersom sträckan är kortare än 20 m används 20 m som nämnare: 21,5 / 20 × 100 = 108 p/100 m, varav konstruktion 100. Klass A (konstruktionsindex ≥ 80)."),
]);

(async () => {
  fs.writeFileSync("Användarhandledning tv3_analys.docx", await Packer.toBuffer(handledning));
  fs.writeFileSync("Metodbeskrivning prioritering avloppsledningar.docx", await Packer.toBuffer(metod));
  console.log("klart");
})();
