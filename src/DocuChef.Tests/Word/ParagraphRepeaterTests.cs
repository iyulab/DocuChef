using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocuChef.Word.Processors;
using FluentAssertions;
using Xunit;

namespace DocuChef.Tests.Word;

public class ParagraphRepeaterTests
{
    [Fact]
    public void ForeachBlock_ExpandsForEachItem()
    {
        // Arrange: template with #foreach / #end block and 2 items
        using var stream = WordTestHelper.CreateDocx(
            "Before",
            "#foreach: Items",
            "Name: ${Name}",
            "#end",
            "After");
        using var doc = WordprocessingDocument.Open(stream, true);
        var body = doc.MainDocumentPart!.Document.Body!;

        var data = new Dictionary<string, object>
        {
            {
                "Items", new List<Dictionary<string, object>>
                {
                    new() { { "Name", "Alice" } },
                    new() { { "Name", "Bob" } }
                }
            }
        };

        // Act
        ParagraphRepeater.ProcessParagraphs(body, data);

        // Assert
        var texts = body.Elements<Paragraph>()
            .Select(p => p.InnerText)
            .ToList();

        // No #foreach or #end markers should remain
        texts.Should().NotContain(t => t.Contains("#foreach"));
        texts.Should().NotContain(t => t.Contains("#end"));

        // Should contain indexed variable expressions
        texts.Should().Contain(t => t.Contains("${Items[0].Name}"));
        texts.Should().Contain(t => t.Contains("${Items[1].Name}"));

        // "Before" and "After" still present
        texts.Should().Contain("Before");
        texts.Should().Contain("After");
    }

    [Fact]
    public void EmptyCollection_RemovesBlock()
    {
        // Arrange: same template but empty collection
        using var stream = WordTestHelper.CreateDocx(
            "Before",
            "#foreach: Items",
            "Name: ${Name}",
            "#end",
            "After");
        using var doc = WordprocessingDocument.Open(stream, true);
        var body = doc.MainDocumentPart!.Document.Body!;

        var data = new Dictionary<string, object>
        {
            { "Items", new List<Dictionary<string, object>>() }
        };

        // Act
        ParagraphRepeater.ProcessParagraphs(body, data);

        // Assert
        var texts = body.Elements<Paragraph>()
            .Select(p => p.InnerText)
            .ToList();

        texts.Should().HaveCount(2);
        texts.Should().Contain("Before");
        texts.Should().Contain("After");
    }

    [Fact]
    public void NoForeachDirective_NoChange()
    {
        // Arrange: no #foreach directive at all
        using var stream = WordTestHelper.CreateDocx(
            "Hello ${Name}!",
            "Bye");
        using var doc = WordprocessingDocument.Open(stream, true);
        var body = doc.MainDocumentPart!.Document.Body!;

        var data = new Dictionary<string, object>
        {
            { "Name", "World" }
        };

        // Act
        ParagraphRepeater.ProcessParagraphs(body, data);

        // Assert
        var texts = body.Elements<Paragraph>()
            .Select(p => p.InnerText)
            .ToList();

        texts.Should().HaveCount(2);
        texts[0].Should().Be("Hello ${Name}!");
        texts[1].Should().Be("Bye");
    }
}
