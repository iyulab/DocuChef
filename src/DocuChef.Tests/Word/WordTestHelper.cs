using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocuChef.Tests.Word;

/// <summary>
/// Static helper that creates test .docx documents programmatically using OpenXML SDK.
/// </summary>
public static class WordTestHelper
{
    /// <summary>
    /// Creates a docx with one paragraph per string, single run each.
    /// </summary>
    public static MemoryStream CreateDocx(params string[] paragraphTexts)
    {
        var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            var body = new Body();

            foreach (var text in paragraphTexts)
            {
                var run = new Run(new Text(text) { Space = SpaceProcessingModeValues.Preserve });
                var paragraph = new Paragraph(run);
                body.Append(paragraph);
            }

            mainPart.Document = new Document(body);
            mainPart.Document.Save();
        }

        stream.Position = 0;
        return stream;
    }

    /// <summary>
    /// Creates a docx with one paragraph containing multiple runs (simulating Word's split-run behavior).
    /// </summary>
    public static MemoryStream CreateDocxWithSplitRuns(params string[] runTexts)
    {
        var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            var body = new Body();
            var paragraph = new Paragraph();

            foreach (var text in runTexts)
            {
                var run = new Run(new Text(text) { Space = SpaceProcessingModeValues.Preserve });
                paragraph.Append(run);
            }

            body.Append(paragraph);
            mainPart.Document = new Document(body);
            mainPart.Document.Save();
        }

        stream.Position = 0;
        return stream;
    }

    /// <summary>
    /// Creates a docx with a table (header row + template row).
    /// </summary>
    public static MemoryStream CreateDocxWithTable(string[] headerTexts, string[] templateRowTexts)
    {
        var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            var body = new Body();
            var table = new Table();

            // Header row
            var headerRow = new TableRow();
            foreach (var text in headerTexts)
            {
                var cell = new TableCell(
                    new Paragraph(
                        new Run(new Text(text) { Space = SpaceProcessingModeValues.Preserve })));
                headerRow.Append(cell);
            }
            table.Append(headerRow);

            // Template row
            var templateRow = new TableRow();
            foreach (var text in templateRowTexts)
            {
                var cell = new TableCell(
                    new Paragraph(
                        new Run(new Text(text) { Space = SpaceProcessingModeValues.Preserve })));
                templateRow.Append(cell);
            }
            table.Append(templateRow);

            body.Append(table);
            mainPart.Document = new Document(body);
            mainPart.Document.Save();
        }

        stream.Position = 0;
        return stream;
    }

    /// <summary>
    /// Reads all paragraph inner texts from a docx stream.
    /// </summary>
    public static List<string> ReadParagraphTexts(Stream stream)
    {
        stream.Position = 0;
        using var doc = WordprocessingDocument.Open(stream, false);
        var body = doc.MainDocumentPart!.Document.Body!;
        return body.Elements<Paragraph>()
            .Select(p => p.InnerText)
            .ToList();
    }

    /// <summary>
    /// Reads table cell texts by row from a docx stream.
    /// </summary>
    public static List<List<string>> ReadTableRows(Stream stream)
    {
        stream.Position = 0;
        using var doc = WordprocessingDocument.Open(stream, false);
        var body = doc.MainDocumentPart!.Document.Body!;
        var table = body.Elements<Table>().First();

        return table.Elements<TableRow>()
            .Select(row => row.Elements<TableCell>()
                .Select(cell => cell.InnerText)
                .ToList())
            .ToList();
    }
}
