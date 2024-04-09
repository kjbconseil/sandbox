// See https://aka.ms/new-console-template for more information
using DocumentFormat.OpenXml.Packaging;
using KJBConseil.WordPubliposting;
using System.Reflection;

string outputPath = Path.Combine(
    Path.GetTempPath(),
    Path.GetRandomFileName() + ".docx");

var myTemplateName = "Template.docx";
string resourceName = "KJBConseil.WordPubliposting." + myTemplateName;
var currentAssembly = Assembly.GetExecutingAssembly();

using (Stream? stream = currentAssembly?.GetManifestResourceStream(resourceName))
{
    if (stream == null)
    {
        throw new FileNotFoundException(
            "File seems to not be embedded.", resourceName);
    }

    using (var fileStream = File.Create(outputPath))
    {
        stream.CopyTo(fileStream);
    }
}

using (var documentFromTemplate =
            WordprocessingDocument.Open(outputPath, true))
{
    var body = documentFromTemplate.MainDocumentPart?.Document.Body ??
        throw new InvalidOperationException($"The body of the XML document is null and should not be.");

    Dictionary<string, string> fieldsWithValues = new()
    {
        { "NumDossier", "0012" },
        { "Name", "Harry Potter" },
    };

    ReplaceVariableByValues.Execute(body, fieldsWithValues);

    documentFromTemplate.MainDocumentPart.Document.Save();
}

Console.WriteLine(outputPath);
Console.ReadKey();
