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
    "media: D:\\Inspektioner\\Filmer            # filmmapp för alla filer i listan",
    "bild: D:\\Inspektioner\\Foton              # bildmapp för alla filer (valfritt)",
    "littera: brunnslittera.csv                # ersättningslittera, se 3.3",
    "",
    "DUF 701.TV3                                # media söks i TV3-filens egen mapp",
    "DUF 702.TV3 ; D:\\Filmer\\DUF702            # egen filmmapp för denna fil",
    "\"C:\\Inspektioner\\2022\\DUF 810.TV3\" ; E:\\Film ; bild: E:\\Foton",
    "../andra_omradet/                          # mapp: alla TV3-filer i den",
    "markprofil: Karta\\markprofil.json           # från ArcMap-verktyget Markprofil, se 7.5",
    "uppstroms: Karta\\uppstroms.csv              # från ArcMap-verktyget Uppströms, se 7.9",
    "gis: Karta\\gisdata.json                    # från ArcMap-verktyget Exportera GIS-data, se 7.11",
    "manuell: forra_korningen\\prioritering.xlsx   # manuella bedömningar och Lagning (m), se 6.1",
    "kostnader: kostnader.csv                   # kostnadsposter för åtgärdspaketet, se 6.1",
  ]),
  spacer(),
  bullet("Tomma rader och rader som börjar med `#` hoppas över."),
  bullet("Relativa sökvägar tolkas relativt listfilens mapp; absoluta sökvägar fungerar som de är. Sökvägar med mellanslag kan skrivas med eller utan citattecken."),
  bullet("En mapp på en rad expanderas till alla `.TV3`-filer i mappen, inklusive undermappar."),
  bullet("`media:` (eller `film:`, `video:`) på en egen rad anger en mapp där filmerna söks för **alla** TV3-filer i listan. Kan upprepas."),
  bullet("`markprofil:` anger en `markprofil.json` från ArcMap-verktyget Markprofil (avsnitt 7.5). Samma sak som `--markprofil` på kommandoraden."),
  bullet("`uppstroms:` anger CSV-filen från ArcMap-verktyget Uppströms i batchläge (avsnitt 7.9): serviser och ledningslängd uppströms per sträcka. Samma sak som `--uppstroms`."),
  bullet("`gis:` anger `gisdata.json` från ArcMap-verktyget Exportera GIS-data (avsnitt 7.11): brunnar och ledningssträckor ur kartan, som filmerna kontrolleras mot. Samma sak som `--gis`."),
  bullet("`manuell:` pekar på en tidigare `prioritering.xlsx` där ni fyllt i Manuell bedömning, Kommentar och Lagning (m) – värdena följer med till nästa körning och styr åtgärdspaketet (avsnitt 6.1). `kostnader:` pekar på kostnadsfilen; utan den används `kostnader.csv` bredvid listfilen eller skriptet."),
  bullet("`bild:` (eller `foto:`) anger en egen mapp för bilderna. Anges ingen bildmapp söks bilderna i filmmapparna. Kan upprepas."),
  bullet("Allt efter `;` på en TV3-rad är mappar för **just den filen**: utan prefix filmmappar, med `bild:` bildmappar. Flera mappar separeras med `;`."),
  bullet("Filmer söks alltid rekursivt (inklusive undermappar), i ordningen: TV3-filens egen mapp, filens egna filmmappar, de gemensamma `media:`-mapparna och sist `--media` från kommandoraden. Bilder söks på samma sätt i bildmapparna (filens egna, sedan `bild:`, sedan `--bilder`, sist TV3-filens egen mapp) om några angetts, annars i filmmapparna. Första träffen på filnamnet används."),
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
    ["`--processer N`", "Antal parallella processer för PDF-rapporterna. Standard är antal kärnor minus en; `1` skriver dem en i taget (t.ex. vid felsökning).", "kärnor − 1"],
    ["`--behall-rapporter`", "Hoppar över PDF-rapporter som redan finns i utdatamappen och skriver bara de som saknas. Länkarna i Excel pekar ändå på dem. Använd vid omkörningar när bara Excel eller kartunderlaget ska uppdateras; Rapporter vars sträcka bytt klass eller littera sedan sist tas alltid bort och skrivs om. Kör utan flaggan när rapportlayouten ändrats.", "av"],
    ["`--diagram`", "Sparar diagrammen som PNG-filer i `diagram`-mappen, t.ex. för PowerPoint. Utan flaggan bäddas de bara in i Excel-filen.", "av"],
    ["`--pptx`", "Skriver `presentation.pptx` i utdatamappen ur mallen `mall.pptx` (se 6.2). Kan också göras efteråt med `python tv3_pptx.py UTDATAMAPP`.", "av"],
    ["`--karta ja|nej`", "Skriver `kartunderlag.json` för kartframställning i ArcMap (se avsnitt 7).", "`ja`"],
    ["`--media KATALOG`", "Extra mapp att söka video- och bildfiler i (kan anges flera gånger). TV3-filens egen mapp söks alltid, inklusive undermappar.", "–"],
    ["`--littera FIL.CSV`", "CSV med ersättningslittera för brunnar som märkts fel vid filmningen (se 3.3). Kan även anges i listfilen.", "–"],
    ["`--manuell FIL.XLSX`", "Tidigare `prioritering.xlsx` med ifyllda manuella bedömningar, kommentarer och Lagning (m) (se 6.1). Kan även anges i listfilen.", "–"],
    ["`--kostnader FIL.CSV`", "Kostnadsfil för åtgärdspaketet (se 6.1).", "`kostnader.csv`"],
    ["`--etapper ja|nej`", "Åtgärdspaket med etappindelning och kostnad (fliken Etapper). `nej` hoppar över.", "`ja`"],
    ["`--uppstroms FIL.CSV`", "CSV från ArcMap-verktyget Uppströms (batchläge) med serviser och ledningslängd uppströms per sträcka (se 7.9). Kolumnerna Serviser uppströms och Längd uppströms fylls i (dolda som standard). Kan även anges i listfilen.", "–"],
    ["`--gis FIL.JSON`", "gisdata.json från ArcMap-verktyget Exportera GIS-data (se 7.11). Filmerna kontrolleras mot GIS (kolumnen GIS-flagga), djup vid brunnarna och anläggningsår fylls i, och okända brunnar får förslag på rätt littera i littera_forslag.csv. Kan även anges i listfilen.", "–"],
  ], [2600, 5000, 1760]),
  spacer(),
  p("Skriptet skriver en kort rapport i terminalen: vilka filer som lästes, antal sträckor per prioritetsklass och topplistan."),

  h2("3.3 Rätta felmärkta brunnar"),
  p("Entreprenören skriver ibland fel littera på en brunn, och då hittar varken Excel-filen eller kartverktyget rätt sträcka. I stället för att redigera TV3-filen anger du rättningarna en gång i en CSV-fil, som används varje gång analysen körs:"),
  ...code([
    "fel;ratt;fil;nr;motbrunn;kommentar",
    "BDNB1005633;BDNB1015633;;;;heter BDNB1015633 i kartan – gäller överallt",
    "KRB65728;KRB68728;DUF 701;60;;bara sträcka 60 i DUF 701.TV3",
    "KRB65728;KRB68728;;;KTB1016017;bara filmningar mellan KRB65728 och KTB1016017",
  ]),
  spacer(),
  bullet("`fel` och `ratt` är felaktigt respektive rätt littera. Avgränsare `;`, `,` eller tab (filen kan redigeras i Excel). Rubrikraden bestämmer kolumnordningen; utan rubrik antas `fel;ratt;kommentar`. `#`-rader hoppas över."),
  bullet("`fil`, `nr` och `motbrunn` är valfria och **begränsar** raden: till en TV3-fil (filnamn, med eller utan `.TV3`), till ett sträcknummer (kolumnen Nr i fliken Prioritering) eller till filmningar där den andra brunnen är `motbrunn`. Använd dem när det felaktiga litterat också är ett riktigt littera på andra sträckor – då får bara den felmärkta filmningen rättningen. Tomma = gäller överallt."),
  bullet("Har flera rader samma felaktiga littera vinner den med flest villkor, så en begränsad rad kan ligga ovanpå en global."),
  bullet("Jämförelsen tar inte hänsyn till versaler, mellanslag eller bindestreck."),
  bullet("Filen anges i listfilen med `littera: brunnslittera.csv` (relativ sökväg tolkas från listfilens mapp) eller på kommandoraden med `--littera`. Flera filer kan anges."),
  p("Rättningen slår igenom i start-, slut- och utgångsbrunn överallt: fliken Prioritering får kolumnen **Littera rättat** med t.ex. `BDNB1005633→BDNB1015633`, PDF-rapporten får en rad med samma text, och `kartunderlag.json` skickar det rättade litterat till kartverktyget. Rapportfilerna får namn efter det rättade litterat."),

  h1("4. Resultatfiler"),
  p("Allt hamnar i utdatamappen (`tv3_resultat` om inget annat anges):"),
  table(["Fil", "Innehåll"], [
    ["`prioritering.xlsx`", "Excel-arbetsbok med flikarna Sammanfattning, Prioritering, Etapper, Observationer, Kodstatistik, Per fil och Material (se avsnitt 5)."],
    ["`diagram/*.png`", "Skrivs bara om du kör med `--diagram`: 1_prioritetsklasser (inspekterad ledningslängd per prioritetsklass), 2_topplista (de N mest kritiska sträckorna, uppdelat på konstruktions- och driftskador), 3_observationer_per_kod (antal skadeobservationer per kod och grad) och 4_klass_per_material (prioritetsklass per material, andel av inspekterad längd)."],
    ["`rapporter/*.pdf`", "Ett inspektionsprotokoll per sträcka, namngivet `<tv3-fil>_<nr>_<klass>_<startbrunn>-<slutbrunn>.pdf`. Innehåller sträckdata (material, dimension, längd, antal anslutningar, index m.m.), schematisk översikt med skador och anslutningar, inklinometerprofil med lutning och svacka, observationstabell färgad efter grad, samt de fotografier som hittats i mediamapparna. Bildnamnen i observationstabellen är klickbara och hoppar till fotografiet längre bak i rapporten."],
    ["`kartunderlag.json`", "En post per sträcka (brunnspar, bedömning, index, material) som ArcMap-skriptet läser för att skapa ett ledningslager. Stängs av med `--karta nej`."],
    ["`presentation.pptx`", "Skrivs bara med `--pptx` (eller `tv3_pptx.py`): PowerPoint ur mallen `mall.pptx` med nyckeltal, diagram, topplista, etapper, en sida per sträcka i toppen, driftåtgärder, svackor, karta och metod (se 6.2)."],
    ["`fel.txt`", "Skrivs bara om någon TV3-fil eller mediamapp saknades eller inte kunde läsas."],
  ], [3600, 5760]),
  spacer(),
  p("De två första diagrammen bäddas alltid in i fliken Sammanfattning. Med `--diagram` sparas alla fyra dessutom som PNG i 200 dpi och kan läggas direkt i PowerPoint. Klassfärgerna är desamma i diagram och Excel: rött A, orange B, gult C, grönt D, grått E."),

  h1("5. Flikarna i Excel"),
  table(["Flik", "Innehåll och användning"], [
    ["Sammanfattning", "Nyckeltal för hela underlaget, fördelning per prioritetsklass (antal och längd), poängmodellens parametrar och två inbäddade diagram."],
    ["Prioritering", "En rad per sträcka, rankad efter prioritetsklass och index. Använd autofiltret för att filtrera på klass, material, ledningstyp, område eller fil. Kolumnen Skador (kod+grad) visar t.ex. 9×YTS4, 1×SPR3. Kolumnerna Avbruten inspektion, Inspekterad flera ggr och Relinad ger extra kontext. Ur inklinometerprofilen beräknas Svackdjup i cm (största stående vattendjup i en svacka), Svackdjup/diameter, Svacklängd, Bakfall längd (meter med lutning mot flödesriktningen) och Lutning; Profil osäker markerar sträckor där inklinometerns fall avviker kraftigt från brunnshöjderna och profilen därför inte bör användas. Anslutningar räknar registrerade anslutningar (AS/AG) på sträckan. Enheten (m, mm, p/100 m …) står på en egen rad under rubrikerna, så rubrikerna är korta. Med markprofil från ArcMap (7.5) fylls även Höjdanpassning, Täckning min/max och Höjdflagga i, och med GIS-data (7.11) kolumnen GIS-flagga (avvikelser mellan film och karta) samt de dolda Lutning GIS, Djup start/slut och Anläggningsår. Bara de viktigaste kolumnerna visas från början; övriga (fil, nr, område, datum, drift- och totalindex, driftåtgärd, profilmått, höjdflagga m.m.) är dolda i grupper och fälls ut med plustecknet ovanför kolumnrubrikerna eller knappen 2 uppe till vänster. Vilka som döljs styrs av `DOLDA_KOLUMNER` i KONFIG. Kolumnerna Manuell bedömning och Kommentar är tomma och avsedda för den egna genomgången av sträckorna. Kolumnen Rapport öppnar sträckans PDF-rapport (länken är relativ, så Excel-filen och mappen rapporter måste ligga kvar bredvid varandra). Kolumnen Videofil är en klickbar länk som öppnar filmen (se avsnitt 7 om länken inte fungerar)."],
    ["Etapper", "Åtgärdspaketet (avsnitt 6.1): en rad per etapp med metod, sträckor, brunnar, längd, dimensioner, anslutningar (hattar), lagning, framschaktade brunnar, kostnad per post och totalt, samt sträckorna i klartext. Kolumnen Schakta fram (manuellt) fylls i om en annan brunn ska schaktas fram än den skriptet valt; den läses tillbaka med `manuell:`."],
    ["Observationer", "Varje observation i klartext med läge (m), tid i filmen, kod, grad, poäng, klockposition, vattennivå, bild och kommentar. Bild och Videofil är klickbara länkar när filerna hittats. Sorterad i samma ordning som prioriteringslistan så att man snabbt hittar detaljerna för en kritisk sträcka."],
    ["Kodstatistik", "Antal observationer per skadekod fördelat på grad 1–4, samt hur många sträckor som berörs."],
    ["Per fil", "Nyckeltal per TV3-fil: projekt, område, period, sträckor, längd, konstruktionsindex, klassfördelning och avbrutna inspektioner. Praktiskt när flera uppdrag analyseras samtidigt."],
    ["Material", "Sträckor, längd, konstruktionsindex och andel klass A+B per material och ledningstyp. Relinade (infodrade) sträckor redovisas som ett eget material \"Relinad\" med ursprungsmaterialet i egen kolumn, så att de inte räknas in i betong- eller plaststatistiken. Samma indelning används i diagrammet Prioritetsklass per material."],
    ["Inspektionsgrad", "Bara med GIS-data (7.11). Hur stor del av ledningsnätet i GIS som är filmat, per driftområde (t.ex. DUF-område) och ledningstyp: antal ledningar och längd i GIS, filmade ledningar och filmad längd, andel, samt av de filmade hur många som är klass A respektive B, deras längd och andelen A+B av den filmade längden (gällande klass, dvs. manuell bedömning före maskinell; är samma ledning filmad från båda håll räknas den sämsta klassen). Raderna 'alla' är summor per område och totalt. Filmade sträckor som inte hittats i GIS redovisas på en egen rad."],
    ["Ej inspekterat", "Bara med GIS-data. En rad per ledningssträcka i GIS som inte finns i någon TV3-fil, med driftområde, brunnar, längd, dimension, material, ledningstyp och anläggningsår – äldst först inom varje område. Underlag för nästa filmningsomgång."],
  ], [2400, 6960]),

  h1("6. Anpassa poängmodellen"),
  p("Alla parametrar ligger samlade överst i `tv3_analys.py` under rubriken **KONFIG** och kan ändras med en vanlig textredigerare utan att röra resten av koden:"),
  table(["Parameter", "Betydelse", "Standard"], [
    ["`GRADPOANG`", "Poäng per grad 1–4.", "1 / 3 / 10 / 30"],
    ["`DRIFTFAKTOR`", "Faktor för driftskador (rötter, inläckage, sediment m.m.) i driftindex.", "0,5"],
    ["`KODFAKTOR`", "Viktning av konstruktionskoder inbördes. `None` = bara graden avgör.", "YTS 0,7; FOG, FRF 0,6; DEA 0,3; övriga 1,0"],
    ["`GRADFAKTOR`", "Vikt för en kod vid en viss grad, ersätter kodvikten för den kombinationen.", "YTS grad 4: 1,0"],
    ["`ATTRIBUTFAKTOR`", "Viktning per attribut, utöver kodvikten. `None` = ingen attributviktning.", "SPR CIRK 0,7"],
    ["`LOPANDE_ENHET_M`", "Längd (m) per poängenhet för löpande skador: poäng × längd / enhet, minst 1. `None` = räkna en gång oavsett längd.", "10"],
    ["`LOPANDE_TAK`", "Högsta faktor en löpande skada kan få. `None` = inget tak.", "5"],
    ["`MINLANGD`", "Minsta längd (m) som används som nämnare vid normering per 100 m.", "20"],
    ["`GRAD4_KODER_A`", "Skadekoder där en enda grad 4-observation räcker för klass A. `None` = alla konstruktionskoder.", "RBR, DEF"],
    ["`TROSKEL_A`", "Konstruktionsindex (p/100 m) som ger klass A.", "80"],
    ["`TROSKEL_B`", "Konstruktionsindex (p/100 m) som ger klass B.", "25"],
    ["`HOJD_SAMMA_M`", "Avviker filens brunnshöjder mindre än så (m) från GIS-vattengången är filen i samma höjdsystem.", "0,3"],
    ["`HOJD_FALL_TOL_M`", "Skiljer sig fallet mellan brunnarna mer än så (m) från GIS korrigeras lutningen, annars bara en förskjutning.", "0,3"],
    ["`TACKNING_MIN_M`", "Mindre täckning (mark minus hjässa) än så (m) flaggas.", "1,0"],
    ["`GIS_FALL_TOL_M`, `GIS_FALL_TOL_ANDEL`", "Fallet mellan brunnarna i filmen får avvika så mycket (m) eller så stor andel av GIS-fallet från GIS innan GIS-flaggan vattengång sätts (det största gäller).", "0,3, 0,5"],
    ["`GIS_RIKTNING_MIN_M`", "Stiger GIS-vattengången mer än så (m) från start- till slutbrunn flaggas riktningen.", "0,05"],
    ["`GIS_LANGD_TOL`, `GIS_LITTERA_LIKHET`", "Förslag på rätt littera för okända brunnar: en GIS-ledning från den kända brunnen med längd inom så stor andel av filmens föreslås; annars littera med minst så stor namnlikhet (0–1).", "0,15, 0,75"],
    ["`GIS_MATERIAL`", "Materialnamn i film och GIS som räknas som samma (BTG = Betong, PVC/PE/PP = Plast …).", "se skriptet"],
    ["`FOTO_MAX_PX`, `FOTO_JPEG_KVALITET`", "Fotona i PDF-rapporterna förminskas till högst så många pixlar på längsta sidan (JPEG-kvalitet enligt den andra; 76 ger ~30 % mindre filer än 82 utan synlig skillnad i A4). `None` = originalstorlek.", "960, 76"],
    ["`RAPPORT_BILD_DPI`, `RAPPORT_BILD_PALETT`", "Upplösning på översikts- och profilbilden i rapporten, och om de sparas som palett-PNG (256 färger, ca 65 % mindre).", "150, `True`"],
    ["`RAPPORT_PROCESSER`", "Antal parallella processer för PDF-rapporterna. `None` = antal kärnor − 1, `1` = en i taget.", "`None`"],
    ["`DOLDA_KOLUMNER`", "Kolumner per flik som döljs som standard i Excel (grupperade, fälls ut med plustecknet ovanför rubrikraden). Rubrik utan enhet. Tom lista = visa allt.", "se skriptet"],
    ["`OFULLSTANDIG_ANDEL`", "Är filmad längd kortare än så gånger kartlängden nådde kameran inte fram; profilen hängs då bara upp i startbrunnen.", "0,85"],
    ["`KODER`", "Skadekoder med klartext och typ (K = konstruktion, D = drift, I = information).", "P93"],
    ["`ATTRIBUT`, `INFOKODER`", "Klartext för attribut (KOMPL, TUNNA …) och koder utan grad (AS, NB …).", "P93"],
  ], [2600, 5000, 1760]),
  spacer(),
  p("Okända koder passerar igenom och visas som de är i Excel, men ger inga poäng förrän de lagts in i `KODER`. Om ni använder en annan standard än P93 (t.ex. P111) läggs koderna till i tabellerna."),
  spacer(),
  h2("6.1 Åtgärdspaket och kostnad"),
  p("Fliken **Etapper** gör prioriteringslistan till en plan. Sträckor med gällande klass A eller B (manuell bedömning om den är ifylld, annars maskinell) får metoden **strumpa**; relinade och klass C–E får ingen åtgärd. Metoden byts aldrig automatiskt till schakt – skriv `schakt` i kolumnen Manuell bedömning för de sträckor som inte går att infodra (kolumnen Åtgärdsflagga pekar ut grad 4-rörbrott/deformation med texten *går strumpa?*). Orden `strumpa` och `ingen` (även `inget`, `ej åtgärd`, `avvakta`) styr på samma sätt, och `schakt` vinner om flera ord förekommer. Klassen styrs av en ensam bokstav A–E först i cellen (`B` eller `B – fogfel`); fri text som *Bevaka* tolkas inte som klass. Vid flerinspekterade brunnspar gäller den sträcka som har manuell bedömning för hela paret. Sammanhängande sträckor med samma metod bildar en etapp, och en C- eller D-sträcka kortare än 60 m mellan två åtgärdssträckor tas med (*medtagen för sammanhang*). Etapperna numreras efter högsta konstruktionsindex, eller efter flest serviser uppströms om `ETAPP_ORDNING = \"konsekvens\"` och uppströmsanalysen (7.9) körts."),
  p("**Brunnstyp.** Strumpan installeras från en nedstigningsbrunn (NB); från tillsyns- och rensbrunnar (TB, RB) går det inte. Typen tas ur TVDAT-koden närmast sträckans ände, annars ur litterats prefix (KNBL → NB, KTB → TB, SRB → RB; okänd typ räknas som NB men flaggas). Säger filmen och litterat olika typ flaggas sträckan med *brunnstyp?* – kontrollera, eftersom varje felbedömd tillsynsbrunn kostar en framschaktning i kalkylen. Har en strumpsträcka bara tillsyns-/rensbrunnar schaktas en brunn fram och en ny nedstigningsbrunn sätts. Det räknas **per brunn i etappen**: ligger flera sådana sträckor efter varandra räcker den gemensamma mittenbrunnen, och skriptet väljer så få brunnar som möjligt. Valet står i kolumnen Framschaktade brunnar; vill ni schakta fram en annan brunn skriver ni dess littera i Schakta fram (manuellt) och kör om med `manuell:`."),
  p("**Lagning.** Skador som brukar kräva punktlagning före strumpning (rörbrott/deformation grad 3–4, fogfel grad 4, ytskada grad 4) listas i Åtgärdsflagga (*lagning? 2×YTS4*). Bedöm i filmen hur många meter som behöver lagas och skriv det i kolumnen **Lagning (m)**; det prissätts per meter. Skriv ett tal (`3` eller `3,5`); text som `3 m` tolkas som 3 och loggas, text utan tal ignoreras med varning."),
  p("**Kostnad** räknas för strumpsträckor ur `kostnader.csv` (semikolon, decimalkomma; posterna `strumpa` per dimensionsintervall kr/m, `hatt` kr/st per anslutning, `lagning` kr/m, `framschaktning` och `ny_brunn` kr/st per framschaktad brunn, `etablering` kr per etapp). Värdena i repot är **schabloner** – byt mot egna erfarenhetsvärden. Belopp får skrivas som `2500`, `2 500` eller `2.500,00`; rader med belopp som inte går att tolka hoppas över med varning. Saknas pris för en dimension flaggas sträckan; en `strumpa`-rad utan dimensionsintervall gäller som reserv för dimensioner utanför intervallen. Sträckor med metod schakt får ingen kostnad (den beror på djup och om spont behövs) utan redovisas med längd under *Schakt – kostnad ej beräknad*. Kostnaderna syns per sträcka i Prioritering (utan brunnsposter), per etapp i Etapper och summerade i Sammanfattning, och följer med till kartan som fälten ETAPP, METOD och KOSTNAD."),
  p("**Arbetsgång.** Kör analysen, fyll i Manuell bedömning, Kommentar, Lagning (m) och eventuellt Schakta fram (manuellt) i Excel, spara filen under nytt namn och ange den med `manuell:` i listfilen. Nästa körning läser värdena (matchning på TV3-fil och sträcknummer, annars brunnspar), räknar om etapper och kostnad och skriver dem i den nya Excel-filen igen, så att genomgången kan fortsätta."),
  table(["Parameter", "Betydelse", "Standard"], [
    ["`KOSTNADSFIL`", "Kostnadsfil som söks bredvid listfilen, annars bredvid skriptet.", "`kostnader.csv`"],
    ["`ATGARD_KLASSER`", "Gällande klasser som får åtgärd (strumpa).", "A, B"],
    ["`SCHAKT_AUTOMATISKT`", "`True` = grad 4 på rörbrott/deformation ger schakt automatiskt. Standard: bara manuell bedömning byter metod.", "`False`"],
    ["`ETAPP_OVERBRYGGA_M`", "C/D-sträcka kortare än så (m) mellan två åtgärdssträckor tas med i etappen.", "60"],
    ["`ETAPP_ORDNING`", "`index` = högsta konstruktionsindex först; `konsekvens` = flest serviser uppströms först.", "`index`"],
    ["`BRUNNSTYP_ANDE_M`", "TVDAT-kod NB/TB/RB så nära (m) sträckans ände gäller för brunnen där.", "1,5"],
    ["`LAGNINGSKODER`", "Kod + grad som listas som lagningsbehov i Åtgärdsflagga.", "RBR 3–4, DEF 3–4, FOG 4, YTS 4"],
    ["`FRAMSCHAKTA_BRUNNAR`", "Brunnar (littera) som alltid schaktas fram; alternativ till kolumnen Schakta fram (manuellt).", "–"],
  ], [2600, 5000, 1760]),

  h2("6.2 PowerPoint ur analysen"),
  p("Med `--pptx` (eller efteråt: `python tv3_pptx.py tv3_resultat`) skrivs `presentation.pptx` i utdatamappen ur mallen `mall.pptx` i repots rot. Sidorna är: titel, avdelare, nyckeltal, längd per prioritetsklass, de mest kritiska sträckorna (diagram och tabell), observationer per skadekod, klass per material, etapper och kostnad, en sida per sträcka i toppen (översiktsbild, upp till tre foton av de allvarligaste skadorna och nyckeluppgifter), driftåtgärder, svackor och bakfall, karta (om `kartbild.png` finns i utdatamappen, från ArcMap-verktyget Exportera kartbild), metod och nästa steg. Vilka sidor som tas med, hur många sträckor som får egen sida och punkterna på sista sidan ändras i KONFIG överst i `tv3_pptx.py`."),
  p("Mallen får anpassas fritt i PowerPoint (färger, typsnitt, logotyp, sidfot) så länge **layoutnamnen** (TITEL, AVDELARE, RUBRIK, RUBRIK_TEXT, RUBRIK_BILD, RUBRIK_TABELL, STRACKA, AVSLUT) och **platshållarnas ordning** i varje layout behålls – skriptet fyller platshållarna i den ordningen. Exempelsidorna i mallen tas bort när presentationen byggs. Körs `tv3_pptx.py` fristående läser den `kartunderlag.json` och TV3-filerna igen (sökvägarna i JSON-filen måste stämma på den datorn), tar manuella bedömningar ur `prioritering.xlsx` i mappen och räknar om etapper och kostnad; så kan presentationen göras om efter genomgången i Excel utan att köra hela analysen."),
  spacer(),
  h1("7. Kartlager i ArcMap"),
  p("Mappen `arcmap` innehåller en verktygslåda för ArcMap 10.x. Lägg till den en gång: högerklicka i ArcToolbox-fönstret, välj **Add Toolbox** och peka på `arcmap/tv3_verktyg.pyt`. Verktygslådan **tv3_analys** får två verktyg:"),
  bullet("**Skapa ledningslager** – välj `kartunderlag.json`, ledningslager, brunnslager, fältet med brunnsbeteckning (standard `EntityID`) och var utdata ska sparas. Lagervalen är rullistor med kartans lager; `A Ledning`, `A Nedstign och övriga brunnar`, `A Rensbrunn/tillsynsbrunn` och `A Platsgjuten brunnspunkt` är förifyllda när de finns. Ta med **alla** lager där brunnar kan ligga – nedstigningsbrunnar och rens-/tillsynsbrunnar ligger i olika lager, och loggen visar hur många av inspektionens brunnar som hittades och vilka brunnstyper som saknas. Valfritt: ett polygonlager som begränsar sökningen, en `.lyr`-fil med symbologi och en CSV-rapport över brunnspar som inte hittades. Under **Matchning** finns tolerans (2 m), max antal brunnar en sträcka får passera (2) och fält från ledningslagret som ska följa med."),
  bullet("**Uppströms** – allt som ligger uppströms om en brunn eller markerad ledning, och i batchläge serviser och längd uppströms per bedömd sträcka (avsnitt 7.9)."),
  bullet("**Exportera kartor (PDF per sträcka)** – en PDF-karta per åtgärdssträcka eller etapp ur layouten, i minsta skala där sträckan ryms (1:200, 1:300, 1:400 …), med automatiskt val mellan liggande och stående mall, samlad PDF och hyperlänkfältet KARTA (avsnitt 7.10)."),
  bullet("**Exportera kartbild** – sparar kartans aktuella vy som `kartbild.png` i utdatamappen, så att presentationen (6.2) får en kartsida. Zooma och tänd rätt lager först."),
  bullet("**Uppdatera bedömning** – välj lagret när du fyllt i manuella bedömningar, så räknas Gällande bedömning, Bedömningstyp och STIL om utan att geometrin rörs."),
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
    ["`RAPPORT`", "Inspektionsprotokoll (PDF)", "Fullständig sökväg till sträckans PDF-rapport, för hyperlänk (se 7.4). Tomt om inga rapporter skapades."],
    ["`VIDEO`", "Videofil", "Fullständig sökväg till filmen, om den hittades vid analysen."],
  ], [1800, 2600, 4960]),
  spacer(),
  p("Fältnamn i ArcGIS får inte innehålla mellanslag eller åäö, därför heter fälten `MASK_BED` och `MAN_BED` med alias **Maskinell bedömning** och **Manuell bedömning**. Alias visas bara om utdata är en filgeodatabas; i en shapefil syns fältnamnen. I en filgeodatabas får `MAN_BED` dessutom en värdelista med klasserna A–E."),
  p("När du fyllt i manuella bedömningar: kör verktyget **Uppdatera bedömning** (eller skriptet med `BARA_UPPDATERA = True`). Då räknas `BEDOMNING`, `BED_TYP` och `STIL` om utan att geometrin byggs om. Vid en full omkörning läses befintliga manuella bedömningar in per brunnspar och återställs automatiskt."),
  h2("7.2 Symbologi"),
  p("Symbologin är färg efter prioritetsklass (rött A, orange B, gult C, grönt D, grått E – samma som i Excel och diagrammen), **heldragen** linje för manuell bedömning och **streckad** för maskinell. Den sätts en gång i ArcMap och sparas som `.lyr`-fil; verktyget Skapa ledningslager applicerar filen automatiskt om den ligger som `arcmap/bedomda_ledningar.lyr` (fältet Symbologi förifylls) eller anges i fältet."),
  p("I lagrets egenskaper: **Symbology > Categories > Unique values**, Value Field `STIL`, **Add All Values**. Det ger tio kategorier – `A - Manuell` … `E - Manuell` heldragna och `A - Maskinell` … `E - Maskinell` streckade, i klassens färg. Spara sedan lagret som `.lyr` (högerklicka på lagret > Save As Layer File) i mappen `arcmap`."),
  h2("7.3 Öppna PDF-rapporten från kartan"),
  p("Objekten har fältet `RAPPORT` med sökvägen till sträckans inspektionsprotokoll. Gör en gång, i lagrets egenskaper: **Display > Support Hyperlinks using field**, välj `RAPPORT` och **Document**. Sedan öppnar Hyperlänk-verktyget (blixten i verktygsfältet Tools) PDF:en när du klickar på en ledning, och i Identify-fönstret blir fältet klickbart. Samma sak går att göra för `VIDEO` om du hellre vill öppna filmen. Inställningen sparas i `.lyr`-filen tillsammans med symbologin, så den följer med automatiskt vid nästa körning."),
  p("Sökvägarna måste fungera **på den dator där ArcMap körs**. Kör du analysen på din egen dator men ArcMap i Citrix: lägg utdatamappen (med `rapporter`) på en nätverksplats som Citrix-klienten når, och läs `kartunderlag.json` därifrån – PDF-sökvägarna härleds från JSON-filens mapp, så då stämmer de. Filmerna ligger på en annan plats än analysen fann dem; ange under **Hyperlänkar** i verktyget mappen med filmerna som Citrix-klienten ser den, så byggs `VIDEO` av filnamnet i den mappen. På samma sätt kan rapportmappen pekas ut om den inte ligger bredvid JSON-filen. Flyttar du mapparna, kör verktyget igen så räknas sökvägarna om."),
  h2("7.4 Webb-GIS och GeoJSON"),
  p("Ska lagret in i ett webb-GIS, fyll i **GeoJSON för webb-GIS** i verktyget. Filen skrivs enligt GeoJSON-standarden (RFC 7946): koordinater i WGS84 (EPSG:4326), två dimensioner, ingen `crs`-uppgift, UTF-8. ArcMaps egen export (Features To JSON) behåller kartans koordinatsystem (SWEREF 99 18 00) och Z/M-värden, vilket de flesta webbkartor inte läser – det är skillnaden."),
  p("Ska shapefilen delas: en shapefil består av minst `.shp`, `.shx`, `.dbf` och `.prj` – alla fyra måste följa med (plus `.cpg` för åäö). Utan `.shx` går filen inte att öppna någonstans. Lagret skrivs sedan version 2 i 2D, eftersom Z-värdena i ledningsnätverket är odefinierade (−9999) och 3D-shapefiler stoppar många webb-GIS."),
  h2("7.5 Markprofil och täckning"),
  p("Verktyget **Markprofil** tar ut markhöjden längs varje sträcka och vattengångsnivåerna i båda brunnarna, så att protokollen kan visa ledningsprofilen mot markytan och täckningen kan beräknas. Kör det efter Skapa ledningslager:"),
  bullet("**Ledningslager från Skapa ledningslager** – lagret med de bedömda sträckorna (behöver fältet NR, som finns i lager skapade från och med september 2026)."),
  bullet("**Ursprungligt ledningslager** och de två fälten med **vattengång vid ledningens start- och slutpunkt**. Fälten föreslås automatiskt om de heter VG_UPP/VG_NED eller liknande."),
  bullet("**Markhöjder** – ett punktlager. Höjden tas ur punkternas Z, eller ur ett fält om du anger ett. Markhöjden i varje punkt längs ledningen är ett avståndsviktat medel av de närmaste mätpunkterna inom sökradien (5 m); saknas punkter inom radien finns ingen markhöjd där."),
  bullet("**Utdata** – `markprofil.json`, som du anger i listfilen (`markprofil:`) eller med `--markprofil` nästa gång du kör `tv3_analys.py`."),
  spacer(),
  p("**Höjdläge.** Filmernas höjder ligger inte alltid i RH2000; ofta sätter entreprenören 0 i start- eller mottagningsbrunnen och bara höjdskillnaden mellan brunnarna är pålitlig. Skriptet jämför därför filens brunnshöjder med GIS-vattengången och hänger upp profilen på GIS: avviker de mindre än 0,3 m används filen som den är (**RH2000 ur filen**); stämmer fallet mellan brunnarna men inte nivån förskjuts hela profilen (**förskjuten till GIS**); stämmer inte heller fallet korrigeras lutningen linjärt så att båda brunnarna hamnar på GIS-nivån (**förskjuten och lutningskorrigerad** – det är också botemedlet mot inklinometerdrift). Har bara en av brunnarna GIS-nivå förskjuts profilen mot den (*förskjuten till GIS vid <brunn>*). Saknas vattengång i GIS helt ritas ledningen med filens höjder och märks **okänt nollplan**; täckning beräknas inte. Saknar filen inklinometerprofil men GIS har vattengång ritas ledningen som rät linje mellan brunnarna, så att markytan och täckningen ändå kommer med. Är den filmade längden mer än 25 % längre än ledningen i kartan flaggas sträckan som trolig felkoppling. Höjdläget står i fliken Prioritering, i protokollets infotabell och i profildiagrammet."),
  p("**Täckning** = markhöjd minus hjässa (vattengång + innerdiameter) i varje markpunkt. Minsta och största täckning finns i Prioritering och kartunderlaget. Sträckor med täckning under 1,0 m får flaggan *Liten täckning*, och sträckor där ledningen hamnar över marken flaggan *Ledning över mark – höjdfel*, vilket nästan alltid betyder fel höjd i någon källa. Är inklinometerprofilen osäker räknas täckningen mot rät linje mellan brunnshöjderna i stället för mot profilen."),
  p("**Avbrutna inspektioner.** När kameran inte nådde fram (koden KAM/HINDE, eller filmad längd under 85 % av ledningens längd i kartan) är filens sluthöjd bara höjden där kameran stannade, inte slutbrunnens. Profilen hängs då upp enbart i den brunn kameran startade i (status *förskjuten till GIS vid <brunn> (avbruten inspektion)*), utan lutningskorrigering, och positionerna skalas inte: en meter i filmen är en meter i kartan räknat från startbrunnen. Markytan och täckningen tas bara med för den filmade delen, och i diagrammet står 'avbrott' i stället för brunnsnamnet i den ände som inte nåddes. Är samma sträcka filmad från båda håll efter ett avbrott hängs varje film upp i sin egen startbrunn."),
  p("I protokollet ritas markytan brun i samma diagram och skala som ledningsprofilen, med den minsta täckningen utsatt."),
  h2("7.6 Projektering av nya stråk"),
  p("Två verktyg för att rita ett nytt VA-stråk och få ut profilen. De har inget med TV-inspektionerna att göra men ligger i samma verktygslåda."),
  p("**Skapa projekteringslager** lägger upp två tomma lager i en filgeodatabas (skapas om den inte finns) och lägger dem i kartan: `<prefix>_Ledning` (linje) med fälten Ledningsnr, Typ (S/D/V/K/T), Dimension (mm), Material, Vattengång uppströms, Vattengång nedströms och Kommentar, samt `<prefix>_Brunn` (punkt) med Brunnsnr, Typ, Locknivå, Bottennivå, Diameter och Kommentar. Rita sedan med ArcMaps vanliga redigering: **rita varje ledning i flödesriktningen**, från uppströms brunn till nedströms, och fyll i vattengången i start- och slutpunkten. Ledningar i samma schakt (spill, dag, vatten) ritas som separata linjer."),
  p("**Projekteringsprofil** ritar profilen längs stråket till PDF och PNG på A3 liggande i exakt skala, och skriver en CSV och en JSON med samma uppgifter. Är några ledningar valda i kartan ritas bara de. I diagrammet: markytan ur markhöjdslagret (punkter med Z eller fält, eller ett raster), vattengång och hjässa per ledning med typ, nummer, dimension, material, längd och fall i promille, brunnarna från botten till lock med locknivå och bottennivå, och en nivåtabell under diagrammet med sektion, vattengång per ledningstyp och markhöjd. Ordningen längs stråket räknas ut av hur ledningarna hänger ihop (ändpunkt mot brunn eller mot annan lednings ände inom toleransen); den längsta kedjan blir referensaxel och övriga ledningstyper projiceras på den. Ledningar som inte följer axeln blir egna stråk med egna filer. Inga minimikrav kontrolleras – bakfall, ledning över mark och saknad vattengång noteras i CSV-kolumnen anmärkning."),
  p("Verktyget ritar med matplotlib i ArcMaps egen Python. Saknas matplotlib skrivs bara CSV och JSON."),
  h2("7.7 Sträckor som inte hittas"),
  p("Brunnspar som inte gick att matcha mot en ledning skrivs till `omatchade_par.csv` med uppgift om vilka av brunnarna som finns i kartan. Vanliga orsaker: brunnsbeteckningen skiljer sig mellan TV3-filen och databasen, brunnen saknas i kartan, eller att ledningen passerar fler brunnar än `MAX_HOPP` tillåter. Beteckningar jämförs utan mellanslag, bindestreck och understreck, och utan hänsyn till versaler."),

  h2("7.8 Svackor och bakfall i kartan"),
  p("Under **Svackor och bakfall** i Skapa ledningslager kan två valfria lager skrivas ur inklinometerprofilerna: **Svackor** (punkter där det stående vattnet är djupast, fält `SVACKA_CM`, `SVACKLANGD`, `ANDEL_DIAM` = svackdjup/diameter, `POS_M`, `OSAKER` och klass) och **Bakfall** (linjer där ledningen lutar mot flödesriktningen, fält `LANGD_M`, `LUTNING` i ‰, `FRAN_M`/`TILL_M`). Båda har `RAPPORT` för hyperlänk till protokollet. Svackor grundare än 2 cm (`SVACKA_MIN_CM` i skriptet) tas inte med, och bakfall redovisas som sammanhängande segment om minst 1 m (`BAKFALL_SEGMENT_MIN_M` i tv3_analys.py). Positionerna är kamerans och läggs ut längs kartlinjen från utgångsbrunnen, skalade med kartlängd/filmlängd; vid avbruten inspektion skalas inte. Sträckor med `OSAKER = Ja` (inklinometern avviker från brunnshöjderna) bör granskas mot protokollet innan de används. Är **GeoJSON för webb-GIS** ifyllt skrivs lagren också som `<namn>_svackor.geojson` och `<namn>_bakfall.geojson`."),
  p("Symbologi sätts en gång i ArcMap och sparas som `arcmap/svackor.lyr` respektive `arcmap/bakfall.lyr`, så appliceras den automatiskt nästa körning. Förslag: graderade symboler på `SVACKA_CM` i tre steg (2–5, 5–10, över 10 cm) med röd färg när `ANDEL_DIAM` överstiger 0,5, och bakfall som bred linje graderad på `LUTNING`. Svackor och bakfall visar var sediment och stopp återkommer; de påverkar inte prioritetsklassen."),

  h2("7.9 Uppströmsanalys"),
  p("Verktyget **Uppströms** samlar allt som ligger uppströms om en punkt i nätet: ledningar, brunnar, serviser, total längd och längsta gren. Startpunkten är en brunn (littera) eller, om fältet lämnas tomt, den ledning som är markerad i kartan (sökningen börjar i dess nedströmsände). Utdata är ett lager med alla uppströmsledningar (fält `AVSTAND_M` till startpunkten längs nätet, `NIVA`, `LANGD_M`, `SERVIS`, `RIKTN_UR`) och en sammanfattning i loggen och valfri CSV."),
  p("**Flödesriktning.** Under Flödesriktning väljs fältet med riktningsattributet i ledningslagret och vilka värden som betyder *med* respektive *mot* ritad riktning (flera värden separeras med `;`). Saknas värde på en ledning används vattengångsfälten (högre vattengång = uppströms), och sist ritad riktning (startvertex = uppströms). Loggen visar hur många ledningar som fick riktning ur attribut, vattengång respektive ritad riktning – är ritad riktning i majoritet är resultatet osäkert."),
  p("**Serviser** anges antingen som ett eget servislager eller som ett fält i ledningslagret med värdet för servis. En servis räknas när någon av dess ändar ligger inom toleransen från en uppströmsledning. **Stopp**: ledningar med ett visst värde i ett fält (t.ex. tryckledning) passeras inte, och brunnar i stopplistan (t.ex. pumpstationer) stoppar sökningen. Under **Begränsning** kan sökningen avgränsas till ett område eller en sökradie kring startbrunnen – hela nätet är stort."),
  p("**Batchläge.** Välj lagret från Skapa ledningslager under Batch, så körs analysen för varje sträcka med sträckans slutbrunn som start (sträckan själv ingår). Fälten `ANT_SERV_U` (serviser uppströms) och `L_UPPSTR` (längd uppströms) skrivs till lagret, och CSV-filen anges sedan i listfilen (`uppstroms:`) så att kolumnerna Serviser uppströms och Längd uppströms fylls i Excel och kartunderlaget. Utan servisuppgifter summeras i stället antalet anslutningar enligt TV-inspektionen för de inspekterade sträckorna uppströms (kolumnen serviser_kalla säger `skattning`). Uppgifterna påverkar inte prioritetsklassen; de är underlag för konsekvens och etappordning."),

  h2("7.10 Kartor per sträcka (PDF)"),
  p("Verktyget **Exportera kartor** skriver en PDF per åtgärdssträcka ur den layout som är öppen. Förbered layouten en gång: välj sidstorlek, lägg in dataramen med de lager som ska synas, skala­stock, norrpil och tre textelement med **Element Name** `TITEL`, `UNDERTITEL` och `SKALA` (Properties > Size and Position). Finns elementen fylls de i per karta (sträcka och brunnar; klass, etapp, metod, längd, material och skala). Kör verktyget från layoutvyn."),
  p("För varje sträcka centreras dataramen på sträckan och skalan sätts till den **minsta i skalserien där sträckan ryms** med marginalen (standard 10 m runt om): först 1:200, sedan 1:300, 1:400, 1:500, 1:750, 1:1000 … Serien och marginalen ändras i dialogen. Ryms sträckan inte i den största skalan används den ändå och loggen säger vilka kartor det gäller. Beräkningen utgår från dataramens storlek på sidan, så en större sida eller liggande format ger fler sträckor i 1:200."),
  bullet("**Vilka sträckor**: sträckor med åtgärd (fältet METOD ifyllt av åtgärdspaketet), klass A och B, markerade i kartan, eller alla. **En karta per etapp** gör i stället en karta per etapp med alla etappens sträckor."),
  bullet("**Markering**: sträckan markeras i kartan under exporten. Standard är ArcMaps urvalsfärg (ert eget urval återställs efteråt); tydligare blir det med ett **markeringslager** – lägg in lagret från Skapa ledningslager en gång till, ge det en bred färgad linje och ange det i dialogen, så sätts dess definitionsfråga till den aktuella sträckan och återställs efteråt."),
  bullet("**Dataramen** ska ha Extent *Automatic* (Data Frame Properties > Data Frame), inte Fixed Scale eller Fixed Extent, och rotation 0. Har layouten flera dataramar (översiktskarta) används den största. I mallarna måste lagret med sträckorna peka på samma featureklass som i den öppna kartan; annars varnar loggen."),
  bullet("**Liggande eller stående**: spara två kopior av kartdokumentet med liggande respektive stående layout (samma lager och textelement) och ange dem under **Mallar**. För varje sträcka räknas skalan för båda och den mall som ger minsta skala används; vid lika skala den där sträckan fyller sidan bäst. En lång sträcka i nord–sydlig riktning får då stående sida i 1:200 i stället för liggande i 1:300. Utan mallar används den öppna kartans layout."),
  bullet("**Utdata**: `<klass>_<fil>_<nr>_<från>-<till>.pdf` (eller `etapp_NN.pdf`) i utdatamappen, en samlad `kartor_strackor.pdf` med alla sidor, och sökvägen i fältet **KARTA** i lagret för hyperlänk (Display > Support Hyperlinks using field), som för RAPPORT. Under Hyperlänkar kan en annan mapp anges om ArcMap körs i Citrix."),

  h2("7.11 Exportera GIS-data och kontrollera filmerna mot kartan"),
  p("Verktyget **Exportera GIS-data** skriver brunnarna och ledningssträckorna mellan brunnar ur kartan till `gisdata.json`, oberoende av TV3-filerna, så att analysen kan stämma av filmerna mot GIS utan att kartan behöver vara öppen vid varje körning. Välj ledningslager och brunnslager (samma som i Skapa ledningslager), fältet med brunnsbeteckning och, under **Fält**, locknivå och brunnstyp i brunnslagret samt vattengång uppströms/nedströms, dimension, material, ledningstyp och anläggningsår i ledningslagret. Fälten gissas när lagret valts; vattengångarna ligger på ledningarna (vid start- och slutpunkten), inte på brunnarna. En ledning i kartan som passerar flera brunnar delas per brunnspar, och vattengången i delningspunkterna interpoleras längs ledningen. Bredvid JSON-filen skrivs `gisdata_brunnar.csv` och `gisdata_ledningar.csv` för granskning i Excel."),
  p("Ange filen i listfilen (`gis:`) eller med `--gis`. Analysen letar upp varje sträckas brunnspar bland GIS-ledningarna (oavsett riktning; finns flera tas den vars längd ligger närmast filmens) och sätter kolumnen **GIS-flagga** i Prioritering när något avviker:"),
  bullet("**brunn saknas i GIS** – litterat i filmen finns inte i kartan. Alla sådana brunnar skrivs till `littera_forslag.csv` i utdatamappen, en rad per sträcka. Är den andra brunnen känd prövas dess grannar i GIS: den granne vars ledning har ungefär filmens längd (inom 15 %, `GIS_LANGD_TOL`) föreslås som rätt brunn, även när namnet inte liknar, och kommentaren anger längderna. Är båda okända, eller finns ingen ledning med passande längd, föreslås de GIS-littera som liknar mest (t.ex. två omkastade siffror). Filen har samma kolumner som `brunnslittera.csv` med motbrunnen ifylld, så rättningen bara gäller det brunnsparet: granska förslagen och ange filen med `littera:` nästa körning."),
  bullet("**ingen ledning i GIS mellan brunnarna** – båda brunnarna finns men ingen ledning förbinder dem (fel sträcka i filmen, eller ledningen saknas i kartan)."),
  bullet("**vattengång** – fallet mellan brunnarna enligt filmens höjder avviker mer än 0,3 m eller 50 % från GIS-fallet. Säger om det är inklinometern eller kartan som bör ifrågasättas; nollplanet spelar ingen roll, bara skillnaden mellan brunnarna."),
  bullet("**riktning** – GIS-vattengången stiger från filmens startbrunn till slutbrunn, dvs. kartan anser att flödet går åt andra hållet."),
  bullet("**material** och **dimension** – filmen och kartan anger olika (BTG och Betong räknas som samma, liksom PVC, PE och PP som Plast; `GIS_MATERIAL` i KONFIG)."),
  p("Dessutom fylls de dolda kolumnerna **Lutning GIS** (‰ enligt kartans vattengångar), **Djup start/slut** (locknivå minus lägsta vattengång i brunnen – schaktdjupet för schaktsträckor), **Anläggningsår** och **Driftområde** i, och saknas markprofil används GIS-vattengång och kartlängd till höjdanpassningen, längdkontrollen och avbrottsbedömningen på samma sätt som markprofilen (7.5). Flaggorna påverkar inte prioritetsklassen."),
  p("**Inspektionsgrad per driftområde.** Ange under **Driftområden** i dialogen ett polygonlager (t.ex. DUF-områdena) och fältet med områdets namn eller nummer, så får varje ledning och brunn i exporten sitt område (ledningens mittpunkt avgör). Analysen skriver då fliken **Inspektionsgrad** – ledningsnätets längd i GIS per område och ledningstyp, hur mycket av det som är filmat och andelen, samt hur stor del av det filmade som är klass A och B – och fliken **Ej inspekterat** med de ledningar som saknar film, äldst först. Filmad längd räknas med kartans längd så att andelen blir jämförbar. Samma uppgifter ligger i kartunderlaget (`inspektionsgrad`, `ej_inspekterat`) och i Sammanfattning. Exporten tar med alla ledningar i de valda lagren (eller inom Begränsa till område), så dagvattenledningar räknas som egen ledningstyp om de ligger i samma lager."),

  h1("8. Vanliga frågor"),
  p("**PDF-rapporterna saknar bilder.** Bilderna hämtas från samma mediamappar som länkarna i Excel. Rapporten anger hur många bilder som hittades och listar de som saknas. Positionerna i rapporten räknas från den brunn kameran startade i (Kamera från), precis som i entreprenörens protokoll, medan Startbrunn/Slutbrunn anger uppströms/nedströms."),
  p("**Körningen tar lång tid.** PDF-rapporterna skrivs parallellt i flera processer (`--processer`), fotona förminskas till rapportstorlek innan de bäddas in (`FOTO_MAX_PX`, standard 960 px på längsta sidan – 290 dpi i A4) och bilddata lagras binärt i stället för textkodat. DUF 701 med 488 foton i full HD tar ungefär en halv minut på en dator med fyra kärnor, mot tre och en halv minut tidigare, och rapportmappen blir mindre än en sjundedel så stor. Vill ni ha ännu mindre filer: sänk `FOTO_MAX_PX` till 800 eller `FOTO_JPEG_KVALITET` till 70 i KONFIG – fotona är det som tar plats, diagrambilderna är små. Använd dessutom `--behall-rapporter` vid omkörningar så skrivs bara rapporter som saknas, `--rapporter AB` för att bara skapa rapporter för sträckor i klass A och B, eller `--rapporter inga` när du bara vill uppdatera Excel-filen. Sätt `FOTO_MAX_PX = None` i KONFIG om fotona ska bäddas in i originalstorlek."),
  p("**Åäö blir fel i resultatet.** TV3-filer är normalt sparade i Windows-1252/ISO-8859-1. Skriptet provar UTF-8 först och faller sedan tillbaka till Windows-1252, så det ska fungera automatiskt. Om en fil ändå blir fel: öppna den i Anteckningar och spara om som UTF-8."),
  p("**Samma sträcka finns två gånger i listan.** Det händer när entreprenören filmat från båda hållen, ofta efter ett avbrott. Kolumnen Inspekterad flera ggr markerar dessa; bedöm dem tillsammans."),
  p("**Sträckor i klass E.** Sträckor utan inspekterad längd (ingen film gjord, bara brunnsregistrering). De behöver inspekteras innan de kan bedömas."),
  p("**Länkarna till film och bilder fungerar inte.** TV3-filen innehåller bara filnamnen (t.ex. `1105202111.mp4`), inte var filerna ligger. Skriptet söker igenom mappen där TV3-filen ligger, inklusive undermappar, samt de mappar du anger i listfilen (`media:`, `bild:` eller `; mapp` efter TV3-filen) eller med `--media` och `--bilder`. Terminalen visar hur många filmer och bilder som hittades. Om filmerna ligger på en annan disk eller server: lägg till mappen i listfilen och kör om. Länkarna är absoluta sökvägar, så Excel-filen måste öppnas på en dator som når samma mapp (samma nätverksenhet/enhetsbokstav)."),
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
  p("Driftskador multipliceras med faktor **0,5**."),
  p("**Löpande skador** (markerade A1 … B1 i filen) räknas vid startmarkeringen och **viktas med längden**: poängen multipliceras med längden delad med 10 m, dock minst 1 och högst 5. En 30 m löpande ytskada grad 3 räknas alltså som tre punktskador (30 p), och ingen löpande skada räknas som mer än fem (max 150 p för grad 4). Utan viktning räknades en 100 m lång skada lika som en enda punkt; med viktning värderas den som fem. Enhet och tak kan ändras i KONFIG (`LOPANDE_ENHET_M`, `LOPANDE_TAK`); `None` stänger av viktningen."),
  h2("4.2 Index per sträcka"),
  p("Poängen summeras per sträcka, separat för konstruktion och drift, och normeras till **poäng per 100 m**:"),
  ...code(["index = summa poäng / max(inspekterad längd, 20 m) × 100"]),
  spacer(),
  p("Nämnaren är aldrig kortare än 20 m. Utan den regeln skulle en enstaka skada på en 5 m lång sträcka ge ett orimligt högt index jämfört med samma skada på en 50 m lång sträcka."),
  p("Tre index redovisas: konstruktionsindex, driftindex och totalindex (summan). Bara konstruktionsindex används för klass och rangordning; driftindex och totalindex är information."),
  p("Konstruktionskoderna viktas dessutom inbördes med `KODFAKTOR` i KONFIG. Syftet är att hitta ledningar som bör strumpinfodras medan det fortfarande går, innan de blivit så dåliga att bara schakt återstår. Sprickor, rörbrott och deformation är de tydligaste tecknen på att röret är på väg att brista och väger fullt (1,0). Ytskada, där väggen tunnas ut men röret fortfarande bär, väger 0,7 – utom grad 4, som enligt P93 betyder att rörväggen är genomfrätt och i praktiken är ett rörbrott; den väger fullt (`GRADFAKTOR`). Fogförskjutning och fog-/rörfel är i första hand läckagepunkter som strumpan tätar och väger 0,6. Defekt anslutning väger 0,3 eftersom den ändå åtgärdas med hatt när anslutningen öppnas efter infodringen. En ytskada grad 3 ger alltså 7 p i stället för 10. Cirkulära sprickor (attribut CIRK) viktas dessutom med 0,7 (`ATTRIBUTFAKTOR`): de beror oftast på en sättning vid en fog och är mindre allvarliga för bärigheten än komplexa och längsgående sprickor, som är de egentliga förvarningarna om brott."),
  p("Utan kodviktning blev 58 av 185 sträckor i DUF 701 klass A, och 43 av dem bara på grund av ytskador. Med viktningen blir 40 sträckor A; de som flyttas till B har i nästan alla fall en enda löpande ytskada grad 3 och inga sprickor."),

  h1("5. Prioritetsklass"),
  p("Klassen bestäms av konstruktionsskadorna – den värsta graden på sträckan och konstruktionsindex:"),
  table(["Klass", "Regel", "Tolkning"], [
    ["A – Åtgärda", "Rörbrott (RBR) eller deformation (DEF) av grad 4 på sträckan, eller konstruktionsindex ≥ 80 p/100 m", "Renovering bör planeras in omgående; teknisk bedömning av metod."],
    ["B – Planera renovering", "Konstruktionsgrad 3, eller konstruktionsindex ≥ 25 p/100 m", "Tas med i den fleråriga förnyelseplanen."],
    ["C – Bevaka", "Övriga sträckor med registrerade skador", "Ingen åtgärd nu; följ upp vid nästa inspektion."],
    ["D – Inga skador", "Inga skadeobservationer", "–"],
    ["E – Ej bedömd", "Ingen inspekterad längd (< 1 m)", "Behöver inspekteras."],
  ], [2400, 3760, 3200]),
  spacer(),
  p("Inom varje klass rangordnas sträckorna efter konstruktionsindex, därefter efter högsta konstruktionsgrad. Driftskador påverkar inte ordningen – att en ledning har rötter är inte ett skäl att renovera den. Det innebär att sammanhängande stråk med många skador hamnar överst, vilket ofta också är de sträckor där en samordnad renovering (t.ex. relining av flera sträckor i följd) ger mest nytta."),

  h1("6. Driftåtgärder"),
  p("Driftskador påverkar varken prioritetsklassen eller rangordningen men flaggas separat, eftersom de kräver åtgärd oavsett om ledningen ska renoveras:"),
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
  bullet("**Svackor, bakfall och lutning** – beräknas ur inklinometerprofilen när sådan finns. Svackdjupet är det största stående vattendjupet: vattnet kan bara lämna ledningen nedströms, så vattenytan i varje punkt ligger på den högsta punkten nedströms om den, och djupet är skillnaden mellan den nivån och punktens höjd. Ligger hela ledningen i bakfall räknas djupet därför upp till utloppets höjd. Svacklängd är den sträcka där stående vatten överstiger 1 cm och bakfall den sammanlagda längden där ledningen lutar mot flödesriktningen (mer än 5 ‰). Svackdjupet sätts också i relation till rördiametern. Stora svackor ger sediment- och kapacitetsproblem och kan motivera åtgärd även utan konstruktionsskador. Inklinometerprofiler kan drifta; om inklinometerns fall avviker mer än 0,3 m eller 50 % från brunnshöjderna markeras profilen som osäker."),

  bullet("**Höjdläge och täckning** – med markhöjder och GIS-vattengång från kartan (ArcMap-verktyget Markprofil) hängs filmens profil upp på GIS-nivåerna i RH2000, eftersom filmerna ofta har ett lokalt nollplan där bara höjdskillnaden mellan brunnarna är pålitlig: avviker filens brunnshöjder mindre än 0,3 m används de som de är, stämmer fallet förskjuts profilen, annars korrigeras även lutningen linjärt mellan brunnarna. Täckningen (mark minus hjässa) beräknas per markpunkt; under 1,0 m eller negativ täckning flaggas. Vid avbruten inspektion hängs profilen bara upp i den brunn kameran startade i, utan lutningskorrigering, och täckningen beräknas för den filmade delen. Täckningen påverkar inte prioritetsklassen men är ett viktigt underlag för åtgärdsval."),

  h1("8. Åtgärdspaket och kostnad"),
  p("Prioriteringen översätts till en plan i fliken Etapper. Sträckor med gällande klass A eller B (manuell bedömning går före maskinell) får metoden strumpinfodring; relinade sträckor och klass C–E får ingen åtgärd. Metoden byts aldrig automatiskt till schakt – det är ett manuellt beslut som skrivs i kolumnen Manuell bedömning, med stöd av flaggan för grad 4-rörbrott och deformation. Sammanhängande sträckor med samma metod bildar en etapp; en C- eller D-sträcka kortare än 60 m mellan två åtgärdssträckor tas med eftersom en sammanhängande infodring oftast är billigare än två etableringar."),
  p("Strumpan installeras från en nedstigningsbrunn. Har en sträcka bara tillsyns- eller rensbrunnar schaktas en brunn fram och ersätts med en nedstigningsbrunn. Framschaktningen räknas per brunn i etappen: en gemensam mittenbrunn mellan två sådana sträckor schaktas fram en gång och ger åtkomst åt båda hållen, och skriptet väljer så få brunnar som möjligt (valet kan låsas manuellt). Punktlagning före strumpning bedöms manuellt i meter; skriptet pekar ut de skador som brukar kräva det."),
  p("Kostnaden per strumpsträcka är foder (kr/m efter dimension) × längd + hattar (kr/st) × antal anslutningar + lagning (kr/m) × bedömda meter. Per etapp tillkommer framschaktning och ny brunn per framschaktad brunn samt en fast etablering. Alla belopp kommer ur kostnadsfilen, som i leveransen innehåller schabloner. Schakt kostnadsberäknas inte, eftersom kostnaden beror på djup och spontbehov; schaktsträckor redovisas med längd och dimension så att de inte glöms i summeringen."),

  h1("9. Begränsningar"),
  bullet("Metoden bygger på entreprenörens kodning. Olika operatörer graderar olika; jämförelser mellan uppdrag bör göras med det i åtanke."),
  bullet("Konsekvens vid fel (ledningens betydelse, trafiklast, närhet till vattendrag, dimension, ålder) ingår inte. Prioritetsklassen beskriver **tillstånd**, inte risk. Vid budgetprioritering bör konsekvens läggas till som en separat faktor."),
  bullet("Bara rörbrott och deformation av grad 4 ger ensamt klass A (`GRAD4_KODER_A` i KONFIG). En grad 4-ytskada eller ett grad 4-fogfel ger 30 poäng och grad 3-regeln ger klass B, men inte automatiskt A – i DUF 701 gällde det fyra sträckor med en enda sådan skada på 40–80 m i övrigt frisk ledning."),
  bullet("Alla parametrar – poäng per grad, driftfaktor, minsta längd och trösklar – kan ändras i skriptets konfigurationsdel. Ändringar bör dokumenteras så att resultat från olika tillfällen förblir jämförbara."),

  h1("10. Exempel"),
  p("Sträckan SRB64009 → SRB1016560 (betong 225 mm, 35,1 m) har nio ytskador grad 4, en komplex spricka grad 3 och en löpande ytskada grad 3 över 35 m. Poäng: 9 × 30 (ytskada grad 4 väger fullt) + 10 × 0,7 (cirkulär spricka) + 10 × 0,7 × 3,5 = 301,5. Konstruktionsindex: 301,5 / 35,1 × 100 ≈ 859 p/100 m. Klass A (konstruktionsindex ≥ 80), rang 1 i uppdraget DUF 701."),
  p("Sträckan KRB68713 → KRB68712 (betong 225 mm, 9,7 m) har en spricka grad 3, en ytskada grad 3 och ett inträngande hinder grad 2. Poäng: 10 + 10 + 3 × 0,5 = 21,5. Eftersom sträckan är kortare än 20 m används 20 m som nämnare: 21,5 / 20 × 100 = 108 p/100 m, varav konstruktion 100. Klass A (konstruktionsindex ≥ 80)."),
]);

(async () => {
  fs.writeFileSync("Användarhandledning tv3_analys.docx", await Packer.toBuffer(handledning));
  fs.writeFileSync("Metodbeskrivning prioritering avloppsledningar.docx", await Packer.toBuffer(metod));
  console.log("klart");
})();
