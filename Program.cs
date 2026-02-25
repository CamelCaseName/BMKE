using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas.Parser;
using Newtonsoft.Json;
using System.Text;



JsonConvert.DefaultSettings = () => new JsonSerializerSettings { MaxDepth = 128 };
string pdfPath = string.Empty;
//todo allow re-entry of file path
if (args.Length < 1)
{
    Console.WriteLine("Please provide a document");
    _ = Console.ReadLine();
    return;
}
if (!File.Exists(args[0]))
{
    Console.WriteLine(args[0] + " is not a valid Path");
    _ = Console.ReadLine();
    return;
}
if (Path.GetExtension(args[0]) != ".pdf")
{
    Console.WriteLine(args[0] + " does not have the .pdf extension");
    _ = Console.ReadLine();
    return;
}
pdfPath = args[0];
Console.WriteLine("PDF read, outputting special BMK:\n");

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
foreach (var bmk in BMKs)
{
    Console.WriteLine($"{bmk}");
}

//todo build into csv, or excel or whatever?

_ = Console.ReadLine();
return;
