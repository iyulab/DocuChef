using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocuChef.Word.Processors;
using FluentAssertions;
using Xunit;
using WTable = DocumentFormat.OpenXml.Wordprocessing.Table;
using WTableRow = DocumentFormat.OpenXml.Wordprocessing.TableRow;

namespace DocuChef.Tests.Word;

public class TableRepeaterTests
{
    [Fact]
    public void SimpleCollection_ExpandsRows()
    {
        // Table: header row + template row with ${Items[].Name} and ${Items[].Qty}
        // 3 items → 1 header + 3 data rows = 4 rows total
        using var stream = WordTestHelper.CreateDocxWithTable(
            new[] { "Name", "Qty" },
            new[] { "${Items[].Name}", "${Items[].Qty}" });
        using var doc = WordprocessingDocument.Open(stream, true);
        var body = doc.MainDocumentPart!.Document.Body!;

        var data = new Dictionary<string, object>
        {
            {
                "Items", new List<Dictionary<string, object>>
                {
                    new() { { "Name", "Apple" }, { "Qty", "3" } },
                    new() { { "Name", "Banana" }, { "Qty", "5" } },
                    new() { { "Name", "Cherry" }, { "Qty", "2" } }
                }
            }
        };

        TableRepeater.ProcessTables(body, data);

        var table = body.Elements<WTable>().First();
        var rows = table.Elements<WTableRow>().ToList();
        rows.Should().HaveCount(4); // 1 header + 3 data rows

        // Header unchanged
        rows[0].InnerText.Should().Contain("Name");
        rows[0].InnerText.Should().Contain("Qty");

        // Expanded rows should have indexed expressions like ${Items[0].Name}
        rows[1].InnerText.Should().Contain("${Items[0].Name}");
        rows[1].InnerText.Should().Contain("${Items[0].Qty}");
        rows[2].InnerText.Should().Contain("${Items[1].Name}");
        rows[3].InnerText.Should().Contain("${Items[2].Name}");
    }

    [Fact]
    public void EmptyCollection_RemovesTemplateRow()
    {
        // Empty Items → template row removed, only header remains
        using var stream = WordTestHelper.CreateDocxWithTable(
            new[] { "Name", "Qty" },
            new[] { "${Items[].Name}", "${Items[].Qty}" });
        using var doc = WordprocessingDocument.Open(stream, true);
        var body = doc.MainDocumentPart!.Document.Body!;

        var data = new Dictionary<string, object>
        {
            { "Items", new List<Dictionary<string, object>>() }
        };

        TableRepeater.ProcessTables(body, data);

        var table = body.Elements<WTable>().First();
        var rows = table.Elements<WTableRow>().ToList();
        rows.Should().HaveCount(1); // only header
    }

    [Fact]
    public void NoArrayExpression_NoChange()
    {
        // Static expression ${StaticValue} (no []) → table unchanged
        using var stream = WordTestHelper.CreateDocxWithTable(
            new[] { "Header" },
            new[] { "${StaticValue}" });
        using var doc = WordprocessingDocument.Open(stream, true);
        var body = doc.MainDocumentPart!.Document.Body!;

        var data = new Dictionary<string, object>
        {
            { "StaticValue", "Hello" }
        };

        TableRepeater.ProcessTables(body, data);

        var table = body.Elements<WTable>().First();
        var rows = table.Elements<WTableRow>().ToList();
        rows.Should().HaveCount(2); // header + template row unchanged
    }
}
