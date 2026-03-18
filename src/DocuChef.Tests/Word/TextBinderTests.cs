using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocuChef.Word.Processors;
using FluentAssertions;
using Xunit;
using WText = DocumentFormat.OpenXml.Wordprocessing.Text;

namespace DocuChef.Tests.Word;

public class TextBinderTests
{
    [Fact]
    public void SimpleVariable_Replaces()
    {
        using var stream = WordTestHelper.CreateDocx("Hello ${Name}!");
        using var doc = WordprocessingDocument.Open(stream, true);
        var body = doc.MainDocumentPart!.Document.Body!;
        var data = new Dictionary<string, object> { { "Name", "World" } };

        TextBinder.Bind(body, data);

        var text = body.Descendants<WText>().First().Text;
        text.Should().Be("Hello World!");
    }

    [Fact]
    public void MultipleVariables_ReplacesAll()
    {
        using var stream = WordTestHelper.CreateDocx("${First} and ${Second}");
        using var doc = WordprocessingDocument.Open(stream, true);
        var body = doc.MainDocumentPart!.Document.Body!;
        var data = new Dictionary<string, object>
        {
            { "First", "A" },
            { "Second", "B" }
        };

        TextBinder.Bind(body, data);

        var text = body.Descendants<WText>().First().Text;
        text.Should().Be("A and B");
    }

    [Fact]
    public void MissingVariable_DoesNotThrow()
    {
        using var stream = WordTestHelper.CreateDocx("Hello ${Unknown}!");
        using var doc = WordprocessingDocument.Open(stream, true);
        var body = doc.MainDocumentPart!.Document.Body!;
        var data = new Dictionary<string, object>();

        var act = () => TextBinder.Bind(body, data);

        act.Should().NotThrow();
    }

    [Fact]
    public void NoExpressions_NoChange()
    {
        using var stream = WordTestHelper.CreateDocx("Plain text");
        using var doc = WordprocessingDocument.Open(stream, true);
        var body = doc.MainDocumentPart!.Document.Body!;
        var data = new Dictionary<string, object> { { "Name", "World" } };

        TextBinder.Bind(body, data);

        var text = body.Descendants<WText>().First().Text;
        text.Should().Be("Plain text");
    }

    [Fact]
    public void NumericValue_Replaces()
    {
        using var stream = WordTestHelper.CreateDocx("Count: ${Count}");
        using var doc = WordprocessingDocument.Open(stream, true);
        var body = doc.MainDocumentPart!.Document.Body!;
        var data = new Dictionary<string, object> { { "Count", 42 } };

        TextBinder.Bind(body, data);

        var text = body.Descendants<WText>().First().Text;
        text.Should().Be("Count: 42");
    }
}
