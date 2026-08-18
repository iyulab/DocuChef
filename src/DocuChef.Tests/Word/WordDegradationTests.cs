using DocuChef.Word;
using FluentAssertions;
using Xunit.Abstractions;

namespace DocuChef.Tests.Word;

/// <summary>
/// Probes Word for the failure shape repeatedly found in the PowerPoint pipeline:
/// when data is absent or a directive cannot be satisfied, does the engine degrade to
/// empty output, or does it ship template syntax to the reader?
/// </summary>
public class WordDegradationTests : TestBase
{
    public WordDegradationTests(ITestOutputHelper output) : base(output) { }

    public record Person(string Name);

    private List<string> Render(string[] paragraphs, string? key = null, object? data = null)
    {
        string tempPath = TempDocx();
        try
        {
            using var stream = WordTestHelper.CreateDocx(paragraphs);
            File.WriteAllBytes(tempPath, stream.ToArray());

            using var chef = CreateNewChef();
            using var recipe = chef.LoadTemplate(tempPath);
            if (key != null) recipe.AddVariable(key, data!);
            using var dish = recipe.Generate();

            var result = new MemoryStream();
            dish.SaveAs(result);

            var texts = WordTestHelper.ReadParagraphTexts(result);
            for (int i = 0; i < texts.Count; i++)
                _output.WriteLine($"para[{i}] = {texts[i]}");
            return texts;
        }
        finally { Cleanup(tempPath); }
    }

    [Fact]
    public void MissingVariable_RendersEmpty_DoesNotLeakExpression()
    {
        var paragraphs = Render(["Hello ${Missing}!"]);

        string.Join("\n", paragraphs).Should().NotContain("${",
            "an unresolvable variable must render empty, not expose the template expression");
    }

    [Fact]
    public void ThrowOnMissingVariable_True_ThrowsInsteadOfDegrading()
    {
        // Regression guard: DollarSignOptions.ErrorHandler was observed to silently suppress
        // ThrowOnError when both are set (verified empirically against DollarSignEngine 1.6.0).
        // TextBinder must only wire ErrorHandler when ThrowOnMissingVariable is false, or this
        // contract silently stops throwing.
        string tempPath = TempDocx();
        try
        {
            using var stream = WordTestHelper.CreateDocx(["Hello ${Missing}!"]);
            File.WriteAllBytes(tempPath, stream.ToArray());

            using var chef = CreateNewChef();
            using var recipe = chef.LoadWordTemplate(tempPath,
                new WordOptions { ThrowOnMissingVariable = true });

            Action act = () => recipe.Generate();

            act.Should().Throw<Exception>(
                "ThrowOnMissingVariable=true must still surface a failure, not silently degrade");
        }
        finally { Cleanup(tempPath); }
    }

    [Fact]
    public void ThrowOnMissingVariable_True_ThrowsInsideForeachToo()
    {
        // cycle-26 verified ThrowOnMissingVariable=true on a plain paragraph; confirm the same
        // contract holds for text reached via ParagraphRepeater's #foreach expansion — the
        // repeater rewrites ${Name} to ${People[i].Name} before TextBinder ever sees it, so this
        // is a distinct code path from the plain-paragraph case.
        string tempPath = TempDocx();
        try
        {
            using var stream = WordTestHelper.CreateDocx(["#foreach: People", "${Missing}", "#end"]);
            File.WriteAllBytes(tempPath, stream.ToArray());

            using var chef = CreateNewChef();
            using var recipe = chef.LoadWordTemplate(tempPath,
                new WordOptions { ThrowOnMissingVariable = true });
            recipe.AddVariable("People", new[] { new Person("Ada") });

            Action act = () => recipe.Generate();

            act.Should().Throw<Exception>(
                "ThrowOnMissingVariable=true must still surface a failure for expressions reached through #foreach expansion");
        }
        finally { Cleanup(tempPath); }
    }

    [Fact]
    public void ForeachOverEmptyCollection_LeavesNoDirectiveText()
    {
        var paragraphs = Render(
            // Inside a #foreach block the property is written bare; the repeater rewrites it
            // to ${People[i].Name} per iteration.
            ["#foreach: People", "${Name}", "#end"],
            "People", Array.Empty<Person>());

        string all = string.Join("\n", paragraphs);
        all.Should().NotContain("#foreach",
            "a consumed directive must be removed from the document even when it expands to nothing");
        all.Should().NotContain("#end");
        all.Should().NotContain("${", "no raw expression may survive");
    }

    [Fact]
    public void ForeachOverPopulatedCollection_ExpandsAndRemovesDirectives()
    {
        var paragraphs = Render(
            // Inside a #foreach block the property is written bare; the repeater rewrites it
            // to ${People[i].Name} per iteration.
            ["#foreach: People", "${Name}", "#end"],
            "People", new[] { new Person("Ada"), new Person("Grace") });

        string all = string.Join("\n", paragraphs);
        all.Should().Contain("Ada").And.Contain("Grace");
        all.Should().NotContain("#foreach").And.NotContain("#end");
        all.Should().NotContain("${");
    }

    private static string TempDocx() =>
        Path.Combine(Path.GetTempPath(), $"docuchef_test_{Guid.NewGuid():N}.docx");

    private static void Cleanup(string path)
    {
        if (File.Exists(path)) File.Delete(path);
    }
}
