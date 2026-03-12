//uses isofonft from this project here under lgplv3+fe https://github.com/hikikomori82/osifont

using iText.IO.Font;
using iText.Kernel.Font;
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas;
using iText.Kernel.Pdf.Canvas.Parser;
using iText.Layout;
using iText.Layout.Element;
using Newtonsoft.Json;
using Org.BouncyCastle.Bcpg;
using System.Collections.Immutable;
using System.Diagnostics;
using System.Globalization;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;

JsonConvert.DefaultSettings = () => new JsonSerializerSettings { MaxDepth = 128 };
bool putHelpOnce = false;
while (true)
{
    string? pdfPath = string.Empty;
    string orderNumber = string.Empty;
    List<string> matNumbers = [];
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
    if (args.Length == 0 && !putHelpOnce)
    {
        putHelpOnce = true;
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

    PdfFont isofont;
    var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream("BMKE.osifont-lgpl3fe.ttf");
    var bytes = new byte[stream?.Length ?? 0];
    stream?.ReadExactly(bytes, 0, (int)stream.Length);

    FontProgram isoFontProgram = FontProgramFactory.CreateFont(bytes);

    Console.WriteLine("Hydraulic schematic read, outputting special BMK:\n");

    var BMKs = ExtractBMK(pdfPath);
    BMKs = Sort(BMKs);

    if (flagC)
    {
        BMKs = Combine(BMKs);
    }

    //todo also extract bmk for the other pages on a per-page basis, and then compare via material number against the complete block drawing??

    string outputPath = ExportToCSV(pdfPath, BMKs);
    AnnounceWrittenFile(flagS, outputPath);
    if (flagP)
    {
        outputPath = ExportToPdf(pdfPath, orderNumber, cellheight, cellWidth, BMKs);
        AnnounceWrittenFile(flagS, outputPath);
    }

    if (args.Length == 0)
    {
        Console.WriteLine("Hit Ctrl+C to closel, or enter a new path:");
    }
    else
    {
        Console.WriteLine("Hit any key to exit.");
        _ = Console.ReadKey();
        return;
    }

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
#### BMKE v1- Tool to generate BMK from a KraussMaffei MX Hydraulic schematic for all special additions on the platens, fixed and moving ("BLOCKABRUF" in site name)####
A csv with all BMK is generated at the same path as the hydraulic schematic every time.

When starting via command line, you can provide the path to the hydraulic schematic as the first argument.
The following flags are also available:
    -p  | generate a pdf version of the resulting BMK directly, alongside the csv.
    -g  | group BMK by schematic page.
    -s  | appends a page to the pdf/continues the csv instead of splitting into a new file once a the table on the current page is full.
    -u  | unsorted, does not alphabetically sort the valve names.
    -c  | output needle valve and core pull valve stacks as one single sticker, and not all as single stickers.
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
                        Console.WriteLine("Cores, needle valves and so on will be bundled");
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
                    Console.WriteLine("You can now enter the flags you want and then hit enter, for example: -s -p");
                    var flags = Console.ReadLine();
                    if (flags?.Length >= 2)
                    {
                        GetFlagsFromArgs(flags.Split(' '));
                    }
                }
                else
                {
                    Console.WriteLine(string.Empty);
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
        candidate = candidate.Trim().Trim('\"').Trim('\'');
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

    List<List<string>> ExtractBMK(string pdfPath)
    {
        PdfReader reader = new(pdfPath);
        PdfDocument pdf = new(reader);
        List<List<string>> Alllines = [];
        for (int page = 1; page <= pdf.GetNumberOfPages(); page++)
        {
            //todo maybe lookup material number against sticker list
            var pageText = PdfTextExtractor.GetTextFromPage(pdf.GetPage(page));
            if (pageText.Contains("INHALT BLOCKABRUF"))
            {
                Alllines.Add([]);
                foreach (var line in pageText.Split('\n'))
                {
                    Alllines[^1].AddRange(line.Split(' '));
                    //todo exclude author short form on revision, need plan the issue happens with
                    if (!line.Contains("BLOCKABRUF") || line.Contains("BEARB") || line.Contains("GEPR") || matNumbers.Count != Alllines.Count - 1)
                    {
                        continue;
                    }

                    var match = MatNumberRegex().Match(line);
                    if (match.Success)
                    {
                        matNumbers.Add(match.Value);
                    }
                    else if (line.Contains("TAKT 1 x"))
                    {
                        matNumbers.Add("Takt 1");
                    }
                    else if (line.Contains("TAKT 2 x"))
                    {
                        matNumbers.Add("Takt 2");
                    }
                }
            }
        }
        reader.Close();

        Dictionary<string, List<string>> GroupedLines = [];
        for (int i = 0; i < matNumbers.Count; i++)
        {
            if (GroupedLines.TryGetValue(matNumbers[i], out var list))
            {
                list.AddRange(Alllines[i]);
            }
            else
            {
                GroupedLines.Add(matNumbers[i], Alllines[i]);
            }
        }
        Alllines = [.. GroupedLines.Values];

        matNumbers = [.. GroupedLines.Keys];


        List<HashSet<string>> AllBMKs = [];

        //A for some lines
        //P for pressure lines
        //G for pumps
        //T for tank return lines
        //L for leakage return lines
        //M for cylinders/motors
        //N for NG valve size
        char[] disallowedChars = ['A', 'P', 'G', 'T', 'L', 'M', 'N', 'H', 'W'];

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

        List<List<string>> returner = [];

        for (int f = 0; f < AllBMKs.Count; f++)
        {
            AllBMKs[f].Remove(string.Empty);
            AllBMKs[f].Remove(" ");
            returner.Add([.. AllBMKs[f]]);
        }

        return returner;
    }

    List<List<string>> Combine(List<List<string>> BMKs)
    {
        for (int x = 0; x < BMKs.Count; x++)
        {
            List<string>? file = BMKs[x];
            bool isMiddlPlaten = matNumbers[x].Contains("Takt");
            //cores fp
            ReplaceCores(file, true);
            //cores mp
            ReplaceCores(file, false);
            //cores prop fp
            ReplacePropCores(file, true, forceMPLayout: isMiddlPlaten);
            //cores prop mp
            ReplacePropCores(file, false);
            //needle valves fp mp
            CheckAndReplaceNeedles(file);
            //parting plane fp
            CheckAndReplacePartingPlane(file, true);
            //parting plane mp
            CheckAndReplacePartingPlane(file, false);
            //tie bar puller stacks
            if (file.Contains("XXXXXXX")) //todo
            {

            }
            //pull off cylinder
            CheckAndReplacePullOffCylinder(file, true);
            //pull off cylinder
            CheckAndReplacePullOffCylinder(file, false);
            //todo maybe possible for mold coupling?
            //todo maybe possible for stäubli mold clamps?
        }
        return BMKs;


        static void CheckAndReplacePullOffCylinder(List<string> file, bool fp)
        {
            int number = 870 + (fp ? 1 : 2);
            int number2 = 880 + (fp ? 1 : 2);
            int i = file.IndexOf($"Q{number}0");
            //todo find version with the other valves
            if (i >= 0)
            {
                string temp = $"Q{number}0";
                if (file.Contains($"Q{number2}0"))
                {
                    temp += $"/{number2}0";
                    file.Remove($"Q{number2}0");
                    file.Remove($"Q{number}0/{number2}0");
                }
                if (file.Contains($"R{number}1"))
                {
                    temp += $"\nR{number}1";
                    file.Remove($"R{number}1");
                }
                if (file.Contains($"F{number}00"))
                {
                    temp += $"\nF{number}00";
                    file.Remove($"F{number}00");
                }
                else if (file.Contains($"F{number}0"))
                {
                    temp += $"\nF{number}0";
                    file.Remove($"F{number}0");
                }
                if (file.Contains($"X{number}2"))
                {
                    temp += $"\nX{number}2";
                    file.Remove($"X{number}2");
                }
                if (file.Contains($"R{number}00"))
                {
                    temp += $"\nR{number}00";
                    file.Remove($"R{number}00");
                }
                else if (file.Contains($"R{number}0"))
                {
                    temp += $"\nR{number}0";
                    file.Remove($"R{number}0");
                }
                if (file.Contains($"R{number}2"))
                {
                    temp += $"\nR{number}2";
                    file.Remove($"R{number}2");
                }
                if (file.Contains($"R{number2}2"))
                {
                    temp += $"/{number2}2";
                    file.Remove($"R{number2}2");
                }
                i = file.IndexOf($"Q{number}0");
                file[i] = temp;
            }
        }

        static void CheckAndReplacePartingPlane(List<string> file, bool fp)
        {
            int number = (97000 + (fp ? 4 : 0));
            int number2 = number + 100;
            int number3 = number + 110;
            for (int x = 0; x < 2; x++)
            {
                int i = file.IndexOf($"K{number}");
                while (i >= 0)
                {
                    string temp = $"K{number}";
                    if (file.Contains($"F{number2}"))
                    {
                        temp += $"\nF{number2}";
                        file.Remove($"F{number2}");
                    }
                    if (file.Contains($"F{number3}"))
                    {
                        temp += $"\nF{number3}";
                        file.Remove($"F{number3}");
                    }
                    if (file.Contains($"F{number}"))
                    {
                        temp += $"\nF{number}";
                        file.Remove($"F{number}");
                    }
                    if (file.Contains($"R{number}"))
                    {
                        temp += $"\nR{number}";
                        file.Remove($"R{number}");
                    }
                    i = file.IndexOf($"K{number}");
                    file[i] = temp;
                    i = file.IndexOf($"K{number}");
                }
                number *= 10;
                number2 *= 10;
                number3 *= 10;
            }
        }

        static void ReplaceCores(List<string> file, bool fp, int coreCounter = 0)
        {
            bool replaced = false;
            int LowerNumber = 700 + coreCounter * 10 + (fp ? 4 : 0);
            int UpperNumber = 710 + coreCounter * 10 + (fp ? 4 : 0);
            if (coreCounter >= 10)
            {
                int halfCounter = coreCounter - 10;
                LowerNumber = 700 + halfCounter * 10 + (fp ? 5 : 1);
                UpperNumber = 710 + halfCounter * 10 + (fp ? 5 : 1);
            }
            int i = file.IndexOf($"Q{LowerNumber}");
            while (i >= 0)
            {
                string temp = $"Q{LowerNumber}";
                if (file.Contains($"Q{UpperNumber}"))
                {
                    temp += $"/{UpperNumber}";
                    file.Remove($"Q{UpperNumber}");
                    file.Remove($"Q{LowerNumber}/{UpperNumber}");
                }
                if (file.Contains($"R{LowerNumber}1"))
                {
                    temp += $"\nR{LowerNumber}1";
                    file.Remove($"R{LowerNumber}1");
                }
                if (file.Contains($"R{LowerNumber}0"))
                {
                    temp += $"\nR{LowerNumber}0";
                    file.Remove($"R{LowerNumber}0");
                }
                if (file.Contains($"R{UpperNumber}0"))
                {
                    temp += $"/{UpperNumber}0";
                    file.Remove($"R{UpperNumber}0");
                    file.Remove($"R{LowerNumber}0/{UpperNumber}0");
                }

                i = file.IndexOf($"Q{LowerNumber}");
                file[i] = temp;
                replaced = true;

                coreCounter += 2;
                LowerNumber = 700 + coreCounter * 10 + (fp ? 4 : 0);
                UpperNumber = 700 + coreCounter * 10 + 10 + (fp ? 4 : 0);
                if (coreCounter >= 10)
                {
                    int halfCounter = coreCounter - 10;
                    LowerNumber = 700 + halfCounter * 10 + (fp ? 5 : 1);
                    UpperNumber = 710 + halfCounter * 10 + (fp ? 5 : 1);
                }
                i = file.IndexOf($"Q{LowerNumber}");
            }
            if (!replaced && coreCounter < 20)
            {
                ReplaceCores(file, fp, coreCounter + 2);
            }
        }

        static void ReplacePropCores(List<string> file, bool fp, int coreCounter = 0, bool forceMPLayout = false)
        {
            bool replaced = false;
            int LowerNumber = 700 + coreCounter * 10 + (fp ? 4 : 0);
            int UpperNumber = 710 + coreCounter * 10 + (fp ? 4 : 0);
            if (coreCounter >= 10)
            {
                int halfCounter = coreCounter - 10;
                LowerNumber = 700 + halfCounter * 10 + (fp ? 5 : 1);
                UpperNumber = 710 + halfCounter * 10 + (fp ? 5 : 1);
            }
            int firstProp = (fp ? 7000 : 6000) + 5 + coreCounter;
            int secondProp = (fp ? 7000 : 6000) + 6 + coreCounter;
            int i = file.IndexOf($"K{firstProp}");
            while (i >= 0)
            {
                string temp = $"K{firstProp}";
                if (fp && !forceMPLayout)
                {
                    if (file.Contains($"K{secondProp}"))
                    {
                        temp += $"\nK{secondProp}";
                        file.Remove($"K{secondProp}");
                    }
                    if (file.Contains($"R{LowerNumber}1"))
                    {
                        temp += $"\nR{LowerNumber}1";
                        file.Remove($"R{LowerNumber}1");
                    }
                    if (file.Contains($"R{LowerNumber}0"))
                    {
                        temp += $"\nR{LowerNumber}0";
                        file.Remove($"R{LowerNumber}0");
                    }
                    if (file.Contains($"R{UpperNumber}0"))
                    {
                        temp += $"/{UpperNumber}0";
                        file.Remove($"R{UpperNumber}0");
                        file.Remove($"R{LowerNumber}0/{UpperNumber}0");
                    }
                }
                else
                {
                    if (file.Contains($"R{LowerNumber}1"))
                    {
                        temp += $"\nR{LowerNumber}1";
                        file.Remove($"R{LowerNumber}1");
                    }
                    if (file.Contains($"K{secondProp}"))
                    {
                        temp += $"\nK{secondProp}";
                        file.Remove($"K{secondProp}");
                    }
                    if (file.Contains($"R{UpperNumber}0"))
                    {
                        temp += $"\nR{UpperNumber}0";
                        file.Remove($"R{UpperNumber}0");
                    }
                    if (file.Contains($"R{LowerNumber}0"))
                    {
                        temp += $"/{LowerNumber}0";
                        file.Remove($"R{LowerNumber}0");
                        file.Remove($"R{LowerNumber}0/{UpperNumber}0");
                    }
                }

                i = file.IndexOf($"K{firstProp}");
                file[i] = temp;
                replaced = true;

                coreCounter += 2;
                LowerNumber = 700 + coreCounter * 10 + (fp ? 4 : 0);
                UpperNumber = 710 + coreCounter * 10 + (fp ? 4 : 0);
                if (coreCounter >= 10)
                {
                    int halfCounter = coreCounter - 10;
                    LowerNumber = 700 + halfCounter * 10 + (fp ? 5 : 1);
                    UpperNumber = 710 + halfCounter * 10 + (fp ? 5 : 1);
                }
                firstProp = (fp ? 7000 : 6000) + 5 + coreCounter;
                secondProp = (fp ? 7000 : 6000) + 6 + coreCounter;
                i = file.IndexOf($"K{firstProp}");
            }
            if (!replaced && coreCounter < 20)
            {
                ReplacePropCores(file, fp, coreCounter + 2, forceMPLayout: forceMPLayout);
            }
        }

        static void ReplaceNeedles(List<string> file, int needleCounter)
        {
            needleCounter *= 2;
            int LowerNumber = 820 + needleCounter;
            int UpperNumber = LowerNumber + 1;
            int i = file.IndexOf($"Q{LowerNumber}");
            while (i >= 0)
            {
                if (needleCounter == 62)
                {
                    UpperNumber = 8839;
                }
                else if (needleCounter > 62)
                {
                    int tempCounter = needleCounter - 64;
                    LowerNumber = 9820 + tempCounter;
                    UpperNumber = tempCounter + 1;
                }
                string temp = $"Q{LowerNumber}";
                if (file.Contains($"Q{UpperNumber}"))
                {
                    temp += $"/{UpperNumber}";
                    file.Remove($"Q{UpperNumber}");
                    file.Remove($"Q{LowerNumber}/{UpperNumber}");
                }
                if (needleCounter == 62)
                {
                    if (file.Contains($"R{UpperNumber}"))
                    {
                        temp += $"\nR{UpperNumber}";
                        file.Remove($"R{UpperNumber}");
                    }
                }
                else
                {
                    if (file.Contains($"R{UpperNumber}0"))
                    {
                        temp += $"\nR{UpperNumber}0";
                        file.Remove($"R{UpperNumber}0");
                    }
                }
                if (file.Contains($"F{LowerNumber}0"))
                {
                    temp += $"\nF{LowerNumber}0";
                    file.Remove($"F{LowerNumber}0");
                }
                if (file.Contains($"R{LowerNumber}0"))
                {
                    temp += $"\nR{LowerNumber}0";
                    file.Remove($"R{LowerNumber}0");
                }
                if (file.Contains($"R{LowerNumber}1"))
                {
                    temp += $"\nR{LowerNumber}1";
                    file.Remove($"R{LowerNumber}1");
                }
                if (file.Contains($"R{UpperNumber}1"))
                {
                    temp += $"/{UpperNumber}1";
                    file.Remove($"R{UpperNumber}1");
                }
                i = file.IndexOf($"Q{LowerNumber}");
                file[i] = temp;

                needleCounter += 2;
                LowerNumber = 820 + needleCounter;
                UpperNumber = LowerNumber + 1;
                i = i = file.IndexOf($"Q{LowerNumber}");
            }
        }

        static void CheckAndReplaceNeedles(List<string> file)
        {
            if (file.Contains("Q820"))
            {
                ReplaceNeedles(file, 0);
            }
            else if (file.Contains("Q828"))
            {
                ReplaceNeedles(file, 4);
            }
            else if (file.Contains("Q836"))
            {
                ReplaceNeedles(file, 8);
            }
            else if (file.Contains("Q844"))
            {
                ReplaceNeedles(file, 12);
            }
            else if (file.Contains("Q852"))
            {
                ReplaceNeedles(file, 16);
            }
            else if (file.Contains("Q860"))
            {
                ReplaceNeedles(file, 20);
            }
            else if (file.Contains("Q868"))
            {
                ReplaceNeedles(file, 24);
            }
            else if (file.Contains("Q876"))
            {
                ReplaceNeedles(file, 28);
            }
            else if (file.Contains("Q9820"))
            {
                ReplaceNeedles(file, 32);
            }
            else if (file.Contains("Q9828"))
            {
                ReplaceNeedles(file, 36);
            }
        }
    }

    string ExportToPdf(string pdfPath, string orderNumber, int cellheight, int cellWidth, List<List<string>> BMKs)
    {
        string localDir = Path.GetDirectoryName(pdfPath) ?? string.Empty;

        int pagecounter = 0;
        if (!flagG)
        {
            List<string> newAllList = [];
            for (int i = 0; i < BMKs.Count; i++)
            {
                newAllList.AddRange(BMKs[i]);
            }
            BMKs.Clear();
            BMKs.Add(newAllList);
        }
        for (int i = 0; i < BMKs.Count; i++)
        {
            int PerPageFileCount = 1;
            int SpentItemCounter = 0;
            do
            {
                outputPath = Path.Combine(localDir, $"{orderNumber}-Hydraulik-{(flagG ? $"{matNumbers[(int)MathF.Min(i, matNumbers.Count - 1)]}-" : string.Empty)}BMK-{pagecounter}.pdf");

                List<string> partlist = new((int)MathF.Min(BMKs[i].Count, Rowcount * Colcount));
                if (!flagS)
                {
                    byte[,] heightTracker = new byte[Colcount, Rowcount];
                    for (int s = SpentItemCounter; s < BMKs[i].Count; s++)
                    {
                        int height = 1 + (int)MathF.Round(BMKs[i][s].Count('\n') / 1.25f);

                        bool foundSpot = false;
                        for (int y = 0; y < heightTracker.GetLength(1); y++)
                        {
                            for (int x = 0; x < heightTracker.GetLength(0); x++)
                            {
                                if (heightTracker[x, y] == 0 && (y + height <= heightTracker.GetLength(1)))
                                {
                                    for (int h = 0; h < height; h++)
                                    {
                                        heightTracker[x, y + h] = 1;
                                    }
                                    SpentItemCounter++;
                                    foundSpot = true;
                                    partlist.Add(BMKs[i][s]);
                                    goto found;
                                }
                            }
                        }
                    found:
                        if (!foundSpot)
                        {
                            PerPageFileCount++;
                            break;
                        }
                    }
                }
                else
                {
                    partlist = BMKs[i];
                }
                SetUpPage(pagecounter + 1, out PdfDocument pdf, out PdfPage page, out Document document, i);
                AddBMKAsTable(partlist, pdf, document, page);
                document.Close();
                pagecounter++;
            } while (PerPageFileCount-- > 1);
        }
        //undo counter increase for last one. stupid but easy fix
        outputPath = Path.Combine(localDir, $"{orderNumber}-Hydraulik-{(flagG ? $"{matNumbers[^1]}-" : string.Empty)}BMK-{pagecounter - 1}.pdf");
        return outputPath;

        void SetUpPage(int pageNumber, out PdfDocument pdf, out PdfPage page, out Document document, int Groupcounter)
        {
            PdfWriter writer = new(outputPath);
            pdf = new(writer);
            //A4 landscape
            pdf.SetDefaultPageSize(iText.Kernel.Geom.PageSize.A4.Rotate());
            page = pdf.AddNewPage();
            isofont = PdfFontFactory.CreateFont(isoFontProgram);
            document = new(pdf);
            document.SetFont(isofont);
            document.Add(new Paragraph($"BMK fuer Blockabruf,BWAP und Blockabruf,FWAP [Auftrag: {orderNumber}] {(flagG ? $"[Material: {matNumbers[Groupcounter]}]" : string.Empty)}"));
            if (flagC)
            {
                document.Add(new Paragraph($"jeweils {cellWidth}x{cellheight}mm, einzeln austrennen. \nMehrfachaufkleber sind ein vielfaches hoch, aber gleich breit. 2 Zeilen: 18mm | 3 Zeilen: 27mm | 4 Zeilen 27mm | 5 Zeilen: 36mm"));
            }
            else
            {
                document.Add(new Paragraph($"jeweils {cellWidth}x{cellheight}mm, einzeln austrennen."));
            }
        }
    }

    void AddBMKAsTable(IEnumerable<string> BMKs, PdfDocument pdf, Document document, PdfPage page)
    {
        int pageNumber = 1;
        int tableLeft = 30;
        int tableBottom = 70;
        Table bmkTable = new(Colcount);
        float PageWidth = page.GetPageSizeWithRotation().GetWidth();
        byte[,] heightTracker = new byte[Colcount, Rowcount];
        foreach (string key in BMKs)
        {
            Cell data = SetUpCell(cellheight, cellWidth, key);
            int height = data.GetRowspan();

            bool foundSpot = false;
            for (int y = 0; y < heightTracker.GetLength(1); y++)
            {
                for (int x = 0; x < heightTracker.GetLength(0); x++)
                {
                    if (heightTracker[x, y] == 0 && (y + height <= heightTracker.GetLength(1)))
                    {
                        for (int h = 0; h < height; h++)
                        {
                            heightTracker[x, y + h] = 1;
                        }
                        bmkTable.AddCell(data);
                        foundSpot = true;
                        goto found;
                    }
                }
            }
        found:
            if (!foundSpot)
            {
                ResetTable();
                height = data.GetRowspan();
                for (int h = 0; h < height; h++)
                {
                    heightTracker[0, h] = 1;
                }
                bmkTable.AddCell(data);
            }

        }
        FinishTable(0);

        //set to 1 for full last row
        void FinishTable(int subtract)
        {
            for (int y = 0; y < heightTracker.GetLength(1); y++)
            {
                int fullRow = 0;
                bool minOneItem = false;
                for (int x = 0; x < heightTracker.GetLength(0); x++)
                {
                    if (heightTracker[x, y] != 0)
                    {
                        fullRow++;
                        minOneItem = true;
                    }
                }
                if ((fullRow != Colcount) && minOneItem)
                {
                    for (int x = 0; x < heightTracker.GetLength(0); x++)
                    {
                        if (heightTracker[x, y] == 0)
                        {
                            heightTracker[x, y] = 1;
                            bmkTable.AddCell(SetUpCell(cellheight, cellWidth, string.Empty));
                        }
                    }
                }
            }

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

        Cell SetUpCell(int cellheight, int cellWidth, string key)
        {
            int sizer = 1;
            if (key.Contains('\n'))
            {
                sizer += (int)MathF.Round(key.Count('\n') / 1.25f);
            }
            Cell data = new(sizer, 1);
            data.SetPadding(0);
            data.SetMargin(0);
            data.SetHeight(mmToPt(sizer * cellheight));
            data.SetMinHeight(mmToPt(sizer * cellheight));
            data.SetMaxHeight(mmToPt(sizer * cellheight));
            data.SetWidth(mmToPt(cellWidth));
            data.SetMinWidth(mmToPt(cellWidth));
            data.SetMaxWidth(mmToPt(cellWidth));
            Paragraph element = new Paragraph(key).SetMultipliedLeading(0.8f);
            element.SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE);
            element.SetHorizontalAlignment(iText.Layout.Properties.HorizontalAlignment.CENTER);
            element.SetMargin(0);
            element.SetPadding(0);
            element.SetMarginTop(-sizer * (flagG && (sizer > 1) ? 3.3f : 2.5f));
            data.Add(element);
            data.SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER);
            data.SetHorizontalAlignment(iText.Layout.Properties.HorizontalAlignment.CENTER);
            data.SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE);
            return data;
        }

        void ResetTable()
        {
            FinishTable(1);
            for (int x = 0; x < heightTracker.GetLength(0); x++)
            {
                for (int y = 0; y < heightTracker.GetLength(1); y++)
                {
                    heightTracker[x, y] = 0;
                }
            }
            pageNumber++;
            _ = pdf.AddNewPage();
            page = pdf.GetPage(pageNumber);
            document.Add(new AreaBreak());
            bmkTable = new(Colcount);
        }
    }

    string ExportToCSV(string pdfPath, List<List<string>> bMKs)
    {
        string localDir = Path.GetDirectoryName(pdfPath) ?? string.Empty;
        string outputPath = Path.Combine(localDir, $"{orderNumber}-Hydraulik-{(flagG ? $"{matNumbers[0]}-" : string.Empty)}BMK-0.csv");
        char seperator = ',';
        StringBuilder sb = new();
        int fileCounter = 1;
        int counter = 0;
        for (int fileIter = 0; fileIter < BMKs.Count; fileIter++)
        {
            List<string>? file = BMKs[fileIter];
            if (flagG)
            {
                counter = 0;
            }
            foreach (string key in file)
            {
                counter++;

                if (key.Contains('\n'))
                {
                    sb.Append($"\"{key}\"");
                }
                else
                {
                    sb.Append($"{key}");
                }

                if (counter % Colcount == 0 && counter > 0)
                {
                    sb.Append('\n');
                }
                else
                {
                    sb.Append(seperator);
                }

                if (!flagS && ((counter >= Rowcount * Colcount)))
                {
                    counter = 0;
                    if (sb[^1] != '\n')
                    {
                        sb.Append('\n');
                    }
                    File.WriteAllText(outputPath, sb.ToString());
                    sb.Clear();

                    outputPath = Path.Combine(localDir, $"{orderNumber}-Hydraulik-{(flagG ? $"{matNumbers[(int)MathF.Min(fileIter, matNumbers.Count - 1)]}-" : string.Empty)}BMK-{fileCounter++}.csv");
                }
            }

            if (flagG)
            {
                if (counter % Colcount != 0)
                {
                    while ((counter + 1) % Colcount != 0)
                    {
                        counter++;
                        sb.Append(seperator);
                    }
                }
                counter = 0;
                if (sb.Length > 0)
                {
                    if (sb[^1] != '\n')
                    {
                        sb.Append('\n');
                    }
                    File.WriteAllText(outputPath, sb.ToString());
                    sb.Clear();

                    outputPath = Path.Combine(localDir, $"{orderNumber}-Hydraulik-{(flagG ? $"{matNumbers[(int)MathF.Min(fileIter + 1, matNumbers.Count - 1)]}-" : string.Empty)}BMK-{fileCounter++}.csv");
                }
            }
        }
        if (flagG)
        {
            outputPath = Path.Combine(localDir, $"{orderNumber}-Hydraulik-{(flagG ? $"{matNumbers[^1]}-" : string.Empty)}BMK-{fileCounter - 2}.csv");
        }
        else
        {
            if (counter % Colcount != 0)
            {
                while ((counter + 1) % Colcount != 0)
                {
                    counter++;
                    sb.Append(seperator);
                }
            }
            if (sb[^1] != '\n')
            {
                sb.Append('\n');
            }

            File.WriteAllText(outputPath, sb.ToString());
        }
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

    List<List<string>> Sort(List<List<string>> BMKs)
    {
        if (!flagU)
        {
            if (flagG)
            {
                for (int i = 0; i < BMKs.Count; i++)
                {
                    BMKs[i] = BMKs[i].ToImmutableSortedSet().ToList();
                }
            }
            else
            {
                BMKs[0] = BMKs[0].ToImmutableSortedSet().ToList();
            }
        }

        return BMKs;
    }
}
partial class Program
{
    [GeneratedRegex("\\d{6,10}")]
    private static partial Regex MatNumberRegex();
}