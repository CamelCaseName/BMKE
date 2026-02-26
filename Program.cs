using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas.Parser;
using Newtonsoft.Json;
using System.Globalization;
using System.Text;

JsonConvert.DefaultSettings = () => new JsonSerializerSettings { MaxDepth = 128 };
string? pdfPath = string.Empty;

//todo allow re-entry of file path
pdfPath = GetPathFromArgs(args);
pdfPath ??= GetPathFromConsole();
if (pdfPath is null)
{
    throw new NullReferenceException("issue in path generation, returned null!");
}

Console.WriteLine("PDF read, outputting special BMK:\n");

HashSet<string> BMKs = ExtractBMK(pdfPath);

foreach (var bmk in BMKs)
{
    Console.WriteLine($"{bmk}");
}

//todo build into csv, or excel or whatever?

_ = Console.ReadLine();
return;

static HashSet<string> ExtractBMK(string pdfPath)
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