using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocuChef.Word.Processors;
using FluentAssertions;
using Xunit;

namespace DocuChef.Tests.Word;

public class RunMergerTests
{
    [Fact]
    public void MergeRuns_SingleRunExpression_NoChange()
    {
        // ${Name} in a single run should not be modified
        using var stream = WordTestHelper.CreateDocx("Hello ${Name}!");
        using var doc = WordprocessingDocument.Open(stream, true);
        var body = doc.MainDocumentPart!.Document.Body!;

        RunMerger.MergeExpressionRuns(body);

        var paragraphs = body.Elements<Paragraph>().ToList();
        paragraphs.Should().HaveCount(1);
        paragraphs[0].InnerText.Should().Be("Hello ${Name}!");
        // Single run should remain as single run
        paragraphs[0].Elements<Run>().Should().HaveCount(1);
    }

    [Fact]
    public void MergeRuns_SplitAcrossThreeRuns_MergesIntoFirst()
    {
        // "${", "Name", "}" in 3 runs → merged to "${Name}" in first run
        using var stream = WordTestHelper.CreateDocxWithSplitRuns("${", "Name", "}");
        using var doc = WordprocessingDocument.Open(stream, true);
        var body = doc.MainDocumentPart!.Document.Body!;

        RunMerger.MergeExpressionRuns(body);

        var paragraph = body.Elements<Paragraph>().First();
        paragraph.InnerText.Should().Be("${Name}");
        // First run should contain the full expression, others should be empty
        var runs = paragraph.Elements<Run>().ToList();
        runs[0].InnerText.Should().Be("${Name}");
        for (int i = 1; i < runs.Count; i++)
        {
            runs[i].InnerText.Should().BeEmpty();
        }
    }

    [Fact]
    public void MergeRuns_MultipleExpressions_MergesAll()
    {
        // "${", "First", "} and ${", "Second", "}" → "${First} and ${Second}"
        using var stream = WordTestHelper.CreateDocxWithSplitRuns("${", "First", "} and ${", "Second", "}");
        using var doc = WordprocessingDocument.Open(stream, true);
        var body = doc.MainDocumentPart!.Document.Body!;

        RunMerger.MergeExpressionRuns(body);

        var paragraph = body.Elements<Paragraph>().First();
        paragraph.InnerText.Should().Be("${First} and ${Second}");
    }

    [Fact]
    public void MergeRuns_NoExpressions_NoChange()
    {
        // Plain text should be untouched
        using var stream = WordTestHelper.CreateDocxWithSplitRuns("Hello ", "World");
        using var doc = WordprocessingDocument.Open(stream, true);
        var body = doc.MainDocumentPart!.Document.Body!;

        var originalTexts = body.Elements<Paragraph>().First()
            .Elements<Run>().Select(r => r.InnerText).ToList();

        RunMerger.MergeExpressionRuns(body);

        var afterTexts = body.Elements<Paragraph>().First()
            .Elements<Run>().Select(r => r.InnerText).ToList();
        afterTexts.Should().BeEquivalentTo(originalTexts);
    }

    [Fact]
    public void MergeRuns_UnmatchedDollarBrace_NoChange()
    {
        // "${" without "}" should be preserved
        using var stream = WordTestHelper.CreateDocxWithSplitRuns("${", "NoClose");
        using var doc = WordprocessingDocument.Open(stream, true);
        var body = doc.MainDocumentPart!.Document.Body!;

        var originalTexts = body.Elements<Paragraph>().First()
            .Elements<Run>().Select(r => r.InnerText).ToList();

        RunMerger.MergeExpressionRuns(body);

        var afterTexts = body.Elements<Paragraph>().First()
            .Elements<Run>().Select(r => r.InnerText).ToList();
        afterTexts.Should().BeEquivalentTo(originalTexts);
    }
}
