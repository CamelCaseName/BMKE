using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas;
using iText.Kernel.Pdf.Canvas.Parser;
using iText.Layout;
using iText.Layout.Element;
using iText.Layout.Renderer;
using Newtonsoft.Json;
using System.Collections.Immutable;
using System.Globalization;
using System.Text;

JsonConvert.DefaultSettings = () => new JsonSerializerSettings { MaxDepth = 128 };
string? pdfPath = string.Empty;
string orderNumber = string.Empty;
int cellheight = 9;
int cellWidth = 37;


//##################################################################
//    script    
//##################################################################


pdfPath = GetPathFromArgs(args);
pdfPath ??= GetPathFromConsole();
if (pdfPath is null)
{
    throw new NullReferenceException("issue in path generation, returned null!");
}

Console.WriteLine("PDF read, outputting special BMK:\n");

var BMKs = ExtractBMK(pdfPath).ToImmutableSortedSet();

//todo also extract bmk for the other pages on a per-page basis, and then compare via material number against the complete block drawing

//todo generate drawing pdf here
//todo build into csv, or excel or whatever?

string outputPath = Path.Combine(Path.GetDirectoryName(pdfPath) ?? string.Empty, orderNumber + "-Hydraulik-BMK.pdf");
PdfWriter writer = new(outputPath);
PdfDocument pdf = new(writer);
//A4 landscape
pdf.SetDefaultPageSize(iText.Kernel.Geom.PageSize.A4.Rotate());

var page = pdf.AddNewPage();
Document document = new(pdf);
document.Add(new Paragraph("BMK fuer Blockabruf,BWAP und Blockabruf,FWAP [" + orderNumber + "]"));

document.Add(new Paragraph($"jeweils {cellWidth}x{cellheight}mm, einzeln austrennen"));

AddBMKAsTable(BMKs, pdf, document, page);

document.Close();

//foreach (var bmk in BMKs)
//{
//    Console.WriteLine($"{bmk}");
//}
Console.WriteLine("Done! Exported BMK to " + outputPath);

_ = Console.ReadLine();
return;


//##################################################################
//    methods
//##################################################################

static float mmToPt(float mm)
{
    return mm / 25.4f * 72;
}

HashSet<string> ExtractBMK(string pdfPath)
{
    PdfReader reader = new(pdfPath);
    PdfDocument pdf = new(reader);
    StringBuilder text = new();
    for (int page = 1; page <= pdf.GetNumberOfPages(); page++)
    {
        //todo maybe lookup material number against sticker list
        var pageText = PdfTextExtractor.GetTextFromPage(pdf.GetPage(page));
        if (pageText.Contains("INHALT BLOCKABRUF"))
        {
            text.Append(pageText);
        }
    }
    reader.Close();

    List<string> lines = [.. text.ToString().Split('\n')];
    HashSet<string> BMKs = new(lines.Count / 2);

    //P for pressure lines
    //G for pumps
    //T for tank return lines
    //L for leakage return lines
    //M for cylinders/motors
    //N for NG valve size
    char[] disallowedChars = ['P', 'G', 'T', 'L', 'M', 'N'];

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
            if (c == ' ')
            {
                string temp = lines[i];
                lines.RemoveAt(i);
                lines.InsertRange(i, temp.Split(' '));
                break;
            }
            else if (!char.IsAsciiDigit(c) && c is not ('/' or ' '))
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

    return BMKs;
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

static string? GetPathFromConsole()
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
            return candidate;
        }
    }
}

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

void AddBMKAsTable(ImmutableSortedSet<string> BMKs, PdfDocument pdf, Document document, PdfPage page)
{
    int Colcount = 7;
    int Rowcount = 14;
    int counter = 0;
    int pageNumber = 1;
    int tableLeft = 30;
    int tableBottom = 70;
    Table bmkTable = new(Colcount);
    float PageWidth = page.GetPageSizeWithRotation().GetWidth();
    foreach (string key in BMKs)
    {
        counter++;
        Cell data = new();
        //1mm = 72pt
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
        bmkTable.AddCell(data);

        if (counter >= Rowcount * Colcount)
        {
            counter = 0;
            FinishTable(1);
            pageNumber++;
            _ = pdf.AddNewPage();
            page = pdf.GetPage(pageNumber);
            document.Add(new AreaBreak());
            bmkTable = new(Colcount);
        }
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
        document.ShowTextAligned("Page: " + pageNumber, PageWidth - 15, 10, iText.Layout.Properties.TextAlignment.RIGHT);
        document.Add(bmkTable);
        document.Flush();
    }
}