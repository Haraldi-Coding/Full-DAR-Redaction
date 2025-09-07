using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using RedactionLibrary;
using Xunit;

namespace RedactionLibrary.Tests;

public class RedactWordFileTests
{
    [Fact]
    public void JoinWithoutQuality_AllNull_ReturnsNull()
    {
        var redactor = new RedactWordFile();
        var result = redactor.JoinWithoutQuality(null, null, null, null, null, null);
        Assert.Null(result);
    }

    [Fact]
    public void JoinWithoutQuality_SingleDoc_ReturnsSameBytes()
    {
        byte[] doc = CreateSimpleDoc("Hello");
        var redactor = new RedactWordFile();
        var result = redactor.JoinWithoutQuality(doc, null, null, null, null, null);
        Assert.NotNull(result);
        Assert.Equal(doc, result);
    }

    private static byte[] CreateSimpleDoc(string text)
    {
        using MemoryStream mem = new();
        using (WordprocessingDocument wordDoc = WordprocessingDocument.Create(mem, WordprocessingDocumentType.Document))
        {
            MainDocumentPart mainPart = wordDoc.AddMainDocumentPart();
            mainPart.Document = new Document();
            Body body = mainPart.Document.AppendChild(new Body());
            body.AppendChild(new Paragraph(new Run(new Text(text))));
        }
        return mem.ToArray();
    }
}
