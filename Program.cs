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

for (int i = 0; i < lines.Count; i++)
{
    if (lines[i].Length < 3)
    {
        continue;
    }
    var span = lines[i].AsSpan().Trim();
    if (!char.IsAsciiLetterUpper(span[0])
        || !char.IsAsciiDigit(span[1]))
    {
        continue;
    }
    if (span[0] is 'P' or 'G' or 'T' or 'L' or 'M')
    {
        continue;
    }

    bool broken = false;
    foreach (var c in span[2..])
    {
        if (!char.IsAsciiLetterOrDigit(c) && c is not ('/' or ' '))
        {
            broken = true;
            break;
        }
        else if (c == ' ')
        {
            string temp = lines[i];
            lines.RemoveAt(i);
            lines.InsertRange(i, temp.Split(' '));
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
