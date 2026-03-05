using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas;
using iText.Kernel.Pdf.Canvas.Parser;
using iText.Layout;
using iText.Layout.Element;
using Newtonsoft.Json;
using System.Collections.Immutable;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Text;

JsonConvert.DefaultSettings = () => new JsonSerializerSettings { MaxDepth = 128 };
string? pdfPath = string.Empty;
string orderNumber = string.Empty;
int cellheight = 9;
int cellWidth = 37;
int Colcount = 7;
int Rowcount = 14;
bool flagP = false; //done
bool flagU = false; //done
bool flagC = false; //todo
bool flagG = false; //done
bool flagS = false; //done


//##################################################################
//    script    
//##################################################################
if (args.Length == 0)
{
    PutHelp();
}
pdfPath = GetPathFromArgs(args);
if (pdfPath is not null)
{
    GetFlagsFromArgs(args);
}
else
{
    pdfPath = GetAllFromConsole();
}
ArgumentNullException.ThrowIfNull(pdfPath, nameof(pdfPath));

Console.WriteLine("PDF read, outputting special BMK:\n");

var BMKs = ExtractBMK(pdfPath);
BMKs = Sort(BMKs);

//todo also extract bmk for the other pages on a per-page basis, and then compare via material number against the complete block drawing??

//todo add material number per page into file name (and pdf document)

string outputPath = ExportToCSV(pdfPath, BMKs);
AnnounceWrittenFile(flagS, outputPath);
if (flagP)
{
    outputPath = ExportToPdf(pdfPath, orderNumber, cellheight, cellWidth, BMKs);
    AnnounceWrittenFile(flagS, outputPath);
}

Console.WriteLine("Hit any key to exit.");

_ = Console.ReadLine();
return;


//##################################################################
//    methods
//##################################################################

static float mmToPt(float mm)
{
    return mm / 25.4f * 72;
}

#region console
void PutHelp()
{
    Console.WriteLine(
"""
#### BMKE - Tool to generate BMK from a KraussMaffei Hydraulic schematic for all special additions on the platens, fixed and moving####
A csv with all BMK is generated at the same path as the hydraulic schematic every time.

When starting via command line, you can provide the path to the hydraulic schematic as the first argument.
The following flags are also available:
    -p  | generate a pdf version of the resulting BMK directly, alongside the csv.
    -g  | group BMK by schematic page.
    -s  | appends a page to the pdf/continues the csv instead of splitting into a new file once a the table on the current page is full.
    -u  | unsorted, does not alphabetically sort the valve names.
    -c  | output cascade and core pull valve stacks as one single sticker, and not all as single stickers.
    -h  | outputs this help text.
"""
        );
}

void GetFlagsFromArgs(string[] args)
{
    flagP = false;
    flagG = false;
    flagS = false;
    flagU = false;
    flagC = false;
    foreach (string arg in args)
    {
        if (arg.Length == 2 && arg[0] == '-')
        {
            switch (arg[1])
            {
                case 'p':
                case 'P':
                    flagP = true;
                    Console.WriteLine("Pdf generation turned on");
                    continue;
                case 'g':
                case 'G':
                    flagG = true;
                    Console.WriteLine("Grouping per page turned on");
                    continue;
                case 's':
                case 'S':
                    flagS = true;
                    Console.WriteLine("Output to single file");
                    continue;
                case 'u':
                case 'U':
                    flagU = true;
                    Console.WriteLine("Sorting turned off");
                    continue;
                case 'c':
                case 'C':
                    flagC = true;
                    Console.WriteLine("Cascades and so on will be bundled");
                    continue;
                case 'h':
                case 'H':
                    PutHelp();
                    continue;
            }
        }
    }
}

static string? GetPathFromArgs(string[] args)
{
    if (args.Length < 1)
    {
        Console.WriteLine("Please provide a document");
        return null!;
    }
    if (!CheckFileIsValid(args[0]))
    {
        return null;
    }
    return args[0];
}

string? GetAllFromConsole()
{
    while (true)
    {
        List<string> pathOptions = [];
        StringBuilder pathOptBuilder = new();

        Console.WriteLine("Enter Path to hydraulic schematic pdf:");
        string documentPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
        pathOptBuilder.Append(documentPath);
        Console.Write(documentPath);
        UpdateDirectoryCache(pathOptBuilder, ref pathOptions);

        var input = Console.ReadKey(intercept: true);
        while (input.Key != ConsoleKey.Enter)
        {
            if (input.Key == ConsoleKey.Tab)
            {
                HandleTabInput(pathOptBuilder, ref pathOptions);
            }
            else
            {
                HandleKeyInput(pathOptBuilder, ref pathOptions, input);
            }

            input = Console.ReadKey(intercept: true);
        }
        Console.Write(input.KeyChar);
        var candidate = pathOptBuilder.ToString();

        if (CheckFileIsValid(candidate))
        {
            Console.SetCursorPosition(0, Console.GetCursorPosition().Top + 1);
            //now ask for flags
            Console.WriteLine("Do you want to set flags? [y] for yes, [enter] or [n] for no");
            if (Console.ReadKey().KeyChar == 'y')
            {
                Console.WriteLine(string.Empty);
                Console.WriteLine("you can now enter the flags you want and then hit enter, for example: -s -u");
                var flags = Console.ReadLine();
                if (flags?.Length >= 2)
                {
                    GetFlagsFromArgs(flags.Split(' '));
                }
            }
            return candidate;
        }
    }
}
#endregion console

#region fileBrowser
// https://stackoverflow.com/a/8946847/1188513
static void ClearCurrentLine()
{
    var currentLine = Console.CursorTop;
    Console.SetCursorPosition(0, Console.CursorTop);
    Console.Write(new string(' ', Console.WindowWidth));
    Console.SetCursorPosition(0, currentLine);
}

static void HandleTabInput(StringBuilder builder, ref List<string> data)
{
    var allinput = builder.ToString();
    string currentPath = string.Empty;
    if (allinput.Length > 3)
    {
        currentPath = Path.GetDirectoryName(allinput)!;
    }
    else
    {
        currentPath = allinput;
    }
    var currentInput = Path.GetFileName(allinput);
    if (currentInput == string.Empty && currentPath?.Length < 3)
    {
        currentInput = allinput;
    }
    var match = data.FirstOrDefault(item => item.StartsWith(currentInput, true, CultureInfo.InvariantCulture));
    if (string.IsNullOrEmpty(match))
    {
        Console.Beep();
        return;
    }
    if (match == currentInput)
    {
        int newIndex = data.IndexOf(match) + 1;
        if (newIndex >= 0 && newIndex < data.Count)
        {
            match = data[newIndex];
        }
        else
        {
            match = data[0];
        }
    }

    ClearCurrentLine();
    builder.Clear();

    currentPath ??= string.Empty;
    match = Path.Combine(currentPath, match);

    Console.Write(match);
    builder.Append(match);
    UpdateDirectoryCache(builder, ref data);
}

static void HandleKeyInput(StringBuilder builder, ref List<string> data, ConsoleKeyInfo input)
{
    var currentInput = builder.ToString();
    if (input.Key == ConsoleKey.Backspace && currentInput.Length > 0)
    {
        if (input.Modifiers.HasFlag(ConsoleModifiers.Control))
        {
            var allinput = builder.ToString();
            var currentpath = Path.GetDirectoryName(allinput)!;
            ClearCurrentLine();
            builder.Clear();
            builder.Append(currentpath);
            Console.Write(currentpath);
        }
        else
        {
            builder.Remove(builder.Length - 1, 1);
            ClearCurrentLine();

            currentInput = currentInput[..^1];
            Console.Write(currentInput);
        }
    }
    else
    {
        var key = input.KeyChar;
        builder.Append(key);
        Console.Write(key);
    }
    UpdateDirectoryCache(builder, ref data);
}

static bool CheckFileIsValid(string? candidate)
{
    if (candidate is null)
    {
        Console.WriteLine("candidate path was null??");
        return false;
    }
    if (!File.Exists(candidate))
    {
        Console.WriteLine("### " + candidate + " is not a valid Path ###");
        return false;
    }
    if (Path.GetExtension(candidate) != ".pdf")
    {
        Console.WriteLine(candidate + " does not have the .pdf extension");
        return false;
    }

    return true;
}

static void UpdateDirectoryCache(StringBuilder builder, ref List<string> data)
{
    try
    {
        var temp = builder.ToString();
        if (temp.Length > 2 && Directory.Exists(temp))
        {
            if (temp.Length > 3)
            {
                temp = Path.GetDirectoryName(temp)!;
            }
            var names = Directory.GetFileSystemEntries(temp);
            data.Clear();
            data.Capacity = names.Length;
            foreach (var name in names)
            {
                data.Add(Path.GetFileName(name));
            }
        }
        else if (temp.Length < 3)
        {
            data.Clear();
            foreach (var drive in DriveInfo.GetDrives())
            {
                //remove \ from name
                data.Add(drive.Name[..^1]);
            }
        }
    }
    catch { }
}
#endregion fileBrowser

IList<IList<string>> ExtractBMK(string pdfPath)
{
    PdfReader reader = new(pdfPath);
    PdfDocument pdf = new(reader);
    List<List<string>> Alllines = [];
    if (!flagG)
    {
        Alllines.Add([]);
    }
    for (int page = 1; page <= pdf.GetNumberOfPages(); page++)
    {
        //todo maybe lookup material number against sticker list
        var pageText = PdfTextExtractor.GetTextFromPage(pdf.GetPage(page));
        if (pageText.Contains("INHALT BLOCKABRUF"))
        {
            if (flagG)
            {
                Alllines.Add([]);
            }
            foreach (var line in pageText.Split('\n'))
            {
                Alllines[^1].AddRange(line.Split(' '));
            }
        }
    }
    reader.Close();


    List<HashSet<string>> AllBMKs = [];

    //A for some lines
    //P for pressure lines
    //G for pumps
    //T for tank return lines
    //L for leakage return lines
    //M for cylinders/motors
    //N for NG valve size
    char[] disallowedChars = ['A', 'P', 'G', 'T', 'L', 'M', 'N'];

    foreach (var lines in Alllines)
    {
        AllBMKs.Add([]);
        var BMKs = AllBMKs[^1];
        for (int i = 0; i < lines.Count; i++)
        {
            if (lines[i].Length < 3)
            {
                continue;
            }
            var span = lines[i].AsSpan().Trim();

            if (disallowedChars.Contains(span[0]))
            {
                continue;
            }

            if ((!char.IsAsciiLetterUpper(span[0])
                || !char.IsAsciiDigit(span[1]))
                && (!char.IsAsciiLetterUpper(span[0])
                || !char.IsAsciiLetterUpper(span[1])
                || !char.IsAsciiDigit(span[2])))
            {
                if (orderNumber == string.Empty && span.Contains("Blatt".AsSpan(), StringComparison.InvariantCulture))
                {
                    orderNumber = span[..6].ToString();
                    Console.WriteLine("Order number " + orderNumber + " found");
                }
                continue;
            }

            //remove KM machine description
            if (span[0] == 'K' && span[1] == 'M')
            {
                continue;
            }

            bool broken = false;
            //split lines where we have miltiple in one, either with space or without
            for (int ci = 2; ci < span[2..].Length; ci++)
            {
                char c = span[ci];
                if (!char.IsAsciiDigit(c) && c is not ('/' or ' '))
                {
                    if (char.IsAsciiLetterUpper(c))
                    {
                        var temp = lines[i].AsSpan();
                        lines.RemoveAt(i);
                        lines.InsertRange(i, [temp[..ci].ToString(), temp[ci..].ToString()]);
                    }
                    else
                    {

                        broken = true;
                    }
                    break;
                }
            }
            if (!broken)
            {
                try
                {
                    BMKs.Add(lines[i].Trim().Trim("/").ToString());
                }
                catch { }
            }
        }
    }

    IList<IList<string>> returner = [];

    for (int f = 0; f < AllBMKs.Count; f++)
    {
        returner.Add([.. AllBMKs[f]]);
    }

    return returner;
}

void AddBMKAsTable(IEnumerable<string> BMKs, PdfDocument pdf, Document document, PdfPage page)
{
    int counter = 0;
    int pageNumber = 1;
    int tableLeft = 30;
    int tableBottom = 70;
    Table bmkTable = new(Colcount);
    float PageWidth = page.GetPageSizeWithRotation().GetWidth();
    foreach (string key in BMKs)
    {
        counter++;
        if (counter > Rowcount * Colcount)
        {
            counter = 1;
            FinishTable(1);
            pageNumber++;
            _ = pdf.AddNewPage();
            page = pdf.GetPage(pageNumber);
            document.Add(new AreaBreak());
            bmkTable = new(Colcount);
        }

        Cell data = new();
        //1mm = 72pt
        SetUpCell(cellheight, cellWidth, key, data);
        bmkTable.AddCell(data);

    }

    FinishTable(0);

    //set to 1 for full last row
    void FinishTable(int subtract)
    {
        var canvas = new PdfCanvas(page);
        float stroke = bmkTable.GetStrokeWidth() ?? 1;
        float width = bmkTable.GetNumberOfColumns() * mmToPt(cellWidth + stroke);
        float height = (bmkTable.GetNumberOfRows() - subtract) * mmToPt(cellheight + stroke);

        bmkTable.SetMargin(0);
        bmkTable.SetPadding(0);
        bmkTable.SetBorder(null);
        bmkTable.SetFixedPosition(tableLeft, tableBottom, width);
        bmkTable.SetHeight(height);

        canvas.MoveTo(tableLeft, tableBottom - 10);
        canvas.LineTo(width + tableLeft, tableBottom - 10);
        canvas.MoveTo(width + tableLeft + 10, tableBottom);
        canvas.LineTo(width + tableLeft + 10, height + tableBottom);

        document.ShowTextAligned((bmkTable.GetNumberOfColumns() * cellWidth).ToString(), width / 2 + tableLeft, tableBottom - 30, iText.Layout.Properties.TextAlignment.CENTER);
        document.ShowTextAligned((bmkTable.GetNumberOfRows() * cellheight).ToString(), width + tableLeft + 30, height / 2 + tableBottom, iText.Layout.Properties.TextAlignment.CENTER, MathF.Tau / 4);
        document.ShowTextAligned("Seite: " + pageNumber, PageWidth - 15, 10, iText.Layout.Properties.TextAlignment.RIGHT);
        document.Add(bmkTable);
        document.Flush();
    }

    void SetUpCell(int cellheight, int cellWidth, string key, Cell data)
    {
        data.SetPadding(0);
        data.SetMargin(0);
        data.SetHeight(mmToPt(cellheight));
        data.SetMinHeight(mmToPt(cellheight));
        data.SetMaxHeight(mmToPt(cellheight));
        data.SetWidth(mmToPt(cellWidth));
        data.SetMinWidth(mmToPt(cellWidth));
        data.SetMaxWidth(mmToPt(cellWidth));
        data.Add(new Paragraph(key));
        data.SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER);
        data.SetHorizontalAlignment(iText.Layout.Properties.HorizontalAlignment.CENTER);
        data.SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE);
    }
}

string ExportToPdf(string pdfPath, string orderNumber, int cellheight, int cellWidth, IList<IList<string>> BMKs)
{
    string localDir = Path.GetDirectoryName(pdfPath) ?? string.Empty;
    string outputPath = Path.Combine(localDir, orderNumber + "-Hydraulik-BMK-0.pdf");

    int totalFileCount = BMKs.Count;
    if (!flagS)
    {
        foreach (var list in BMKs)
        {
            if (list.Count > Rowcount * Colcount)
            {
                totalFileCount += (int)MathF.Floor(list.Count / (Rowcount * Colcount));
            }
        }
    }
    int pagecounter = 1;
    for (int i = 0; i < BMKs.Count; i++)
    {
        int PerPageFileCount = 1;
        int SpentItemCounter = 0;
        if (!flagS && BMKs[i].Count > Rowcount * Colcount)
        {
            PerPageFileCount += (int)MathF.Floor(BMKs[i].Count / (Rowcount * Colcount));
        }
        do
        {
            IList<string> partlist = new List<string>((int)MathF.Min(BMKs[i].Count, Rowcount * Colcount));
            if (!flagS)
            {
                int pageFullCounter = 0;
                for (int x = SpentItemCounter; x < BMKs[i].Count; x++)
                {
                    if (pageFullCounter >= Rowcount * Colcount)
                    {
                        break;
                    }
                    pageFullCounter++;
                    partlist.Add(BMKs[i][x]);
                    SpentItemCounter++;
                }
            }
            else
            {
                partlist = BMKs[i];
            }
            SetUpPage(pagecounter, totalFileCount, out PdfDocument pdf, out PdfPage page, out Document document);
            AddBMKAsTable(partlist, pdf, document, page);
            document.Close();
            outputPath = Path.Combine(localDir, $"{orderNumber}-Hydraulik-BMK-{pagecounter}.pdf");
            pagecounter++;
        } while (PerPageFileCount-- > 1);
    }
    //undo counter increase for last one. stupid but easy fix
    outputPath = Path.Combine(localDir, $"{orderNumber}-Hydraulik-BMK-{pagecounter - 2}.pdf");
    return outputPath;

    void SetUpPage(int pageNumber, int pageCount, out PdfDocument pdf, out PdfPage page, out Document document)
    {
        PdfWriter writer = new(outputPath);
        pdf = new(writer);
        //A4 landscape
        pdf.SetDefaultPageSize(iText.Kernel.Geom.PageSize.A4.Rotate());
        page = pdf.AddNewPage();
        document = new(pdf);
        document.Add(new Paragraph($"BMK fuer Blockabruf,BWAP und Blockabruf,FWAP [Auftrag: {orderNumber}] {(pageNumber > 0 ? $"{{Seite {pageNumber}/{pageCount}}}" : string.Empty)}"));
        document.Add(new Paragraph($"jeweils {cellWidth}x{cellheight}mm, einzeln austrennen"));
    }
}

string ExportToCSV(string pdfPath, IList<IList<string>> bMKs)
{
    string localDir = Path.GetDirectoryName(pdfPath) ?? string.Empty;
    string outputPath = Path.Combine(localDir, orderNumber + "-Hydraulik-BMK-0.csv");
    StringBuilder sb = new();
    int fileCounter = 1;
    foreach (var file in BMKs)
    {
        int counter = 0;
        foreach (string key in file)
        {
            counter++;

            sb.Append(key + ",");

            //todo split correctly here if more than col*row but less than what should be for file
            if (!flagS && ((counter >= Rowcount * Colcount)))
            {
                counter = 0;
                File.WriteAllText(outputPath, sb.ToString());
                sb.Clear();

                outputPath = Path.Combine(localDir, $"{orderNumber}-Hydraulik-BMK-{fileCounter++}.csv");
            }
            if (counter % Colcount == 0 && counter > 0)
            {
                sb.Append('\n');
            }
        }
        counter = 0;
        File.WriteAllText(outputPath, sb.ToString());
        sb.Clear();

        outputPath = Path.Combine(localDir, $"{orderNumber}-Hydraulik-BMK-{fileCounter++}.csv");
    }
    outputPath = Path.Combine(localDir, $"{orderNumber}-Hydraulik-BMK-{fileCounter - 2}.csv");
    return outputPath;
}

static void AnnounceWrittenFile(bool flagS, string outputPath)
{
    Console.Write($"Done! Exported BMK to {outputPath}");
    if (!flagS)
    {
        Console.Write(" (last file)");
    }
    Console.Write("\n");
}

IList<IList<string>> Sort(IList<IList<string>> BMKs)
{
    if (!flagU)
    {
        if (flagG)
        {
            for (int i = 0; i < BMKs.Count; i++)
            {
                BMKs[i] = BMKs[i].ToImmutableSortedSet();
            }
        }
        else
        {
            BMKs[0] = BMKs[0].ToImmutableSortedSet();
        }
    }

    return BMKs;
}