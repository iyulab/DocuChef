using DocuChef.Presentation;
using FluentAssertions;
using Xunit.Abstractions;

namespace DocuChef.Tests.PowerPoint;

/// <summary>
/// Probes the PowerPoint pipeline's behavior when an expression references a variable
/// that was never supplied — the failure shape that previously leaked a raw, unroutable
/// console log from DollarSignEngine with no DocuChef-side signal at all.
/// </summary>
public class PowerPointDegradationTests : TestBase
{
    public PowerPointDegradationTests(ITestOutputHelper output) : base(output) { }

    private static string TempPptx() =>
        Path.Combine(Path.GetTempPath(), $"docuchef_test_{Guid.NewGuid():N}.pptx");

    private static void Cleanup(string path)
    {
        if (File.Exists(path)) File.Delete(path);
    }

    [Fact]
    public void MissingVariable_RendersEmpty_DoesNotLeakExpression()
    {
        string tempPath = TempPptx();
        try
        {
            using var stream = PowerPointTestHelper.CreatePptx("Hello ${Missing}!");
            File.WriteAllBytes(tempPath, stream.ToArray());

            using var chef = CreateNewChef();
            using var recipe = chef.LoadTemplate(tempPath);
            using var dish = recipe.Generate();

            var resultStream = new MemoryStream();
            dish.SaveAs(resultStream);

            var texts = PowerPointTestHelper.ReadAllText(resultStream);
            string.Join("\n", texts).Should().NotContain("${",
                "an unresolvable variable must degrade to empty text, not expose the template expression");
        }
        finally { Cleanup(tempPath); }
    }

    [Fact]
    public void ThrowOnMissingVariable_True_ThrowsInsteadOfDegrading()
    {
        // Regression guard: DollarSignOptions.ErrorHandler was observed to silently suppress
        // ThrowOnError when both are set (verified empirically against DollarSignEngine 1.6.0).
        // DataBinder must only wire ErrorHandler when ThrowOnMissingVariable is false, or this
        // contract silently stops throwing.
        string tempPath = TempPptx();
        try
        {
            using var stream = PowerPointTestHelper.CreatePptx("Hello ${Missing}!");
            File.WriteAllBytes(tempPath, stream.ToArray());

            using var chef = CreateNewChef();
            using var recipe = chef.LoadPowerPointTemplate(tempPath,
                new PowerPointOptions { ThrowOnMissingVariable = true });

            Action act = () => recipe.Generate();

            act.Should().Throw<Exception>(
                "ThrowOnMissingVariable=true must still surface a failure, not silently degrade");
        }
        finally { Cleanup(tempPath); }
    }

    [Fact]
    public void ThrowOnMissingVariable_True_ThrowsInsideNestedContextChildSlide()
    {
        // cycle-26 verified the single-slide case. Nested context ('>') runs through a
        // different plan-generation path (SlidePlanGenerator → ProcessNestedRangeSlides,
        // see PPT_DESIGN.md) before DataBinder ever sees the child slide's expressions —
        // confirm the exception still surfaces from there, not just the simple case.
        string tempPath = TempPptx();
        try
        {
            using var stream = PowerPointTestHelper.CreatePptx(new[]
            {
                "${Categories[0].Name}",            // parent slide (no '>')
                "${Categories>Items[0].Bogus}",     // child slide — Item has no 'Bogus' property
            });
            File.WriteAllBytes(tempPath, stream.ToArray());

            using var chef = CreateNewChef();
            using var recipe = chef.LoadPowerPointTemplate(tempPath,
                new PowerPointOptions { ThrowOnMissingVariable = true });
            recipe.AddVariable("Categories", new[]
            {
                new Category("Electronics", [new Item("Smartphone", 999m)])
            });

            Action act = () => recipe.Generate();

            act.Should().Throw<Exception>(
                "an undefined property on a nested-context child slide must still throw under ThrowOnMissingVariable=true");
        }
        finally { Cleanup(tempPath); }
    }

    [Fact]
    public void ThrowOnMissingVariable_True_ThrowsInsideForeachDirectiveSlide()
    {
        // #foreach-driven slides go through the same DataBinder as auto-detected arrays, but
        // via a distinct plan-generation branch (explicit directive vs. inferred pattern) —
        // confirm the exception isn't swallowed somewhere along that separate path either.
        string tempPath = TempPptx();
        try
        {
            using var stream = PowerPointTestHelper.CreatePptxWithNotes(
                ("${Products[0].Bogus}", "#foreach: Products, max: 1"));
            File.WriteAllBytes(tempPath, stream.ToArray());

            using var chef = CreateNewChef();
            using var recipe = chef.LoadPowerPointTemplate(tempPath,
                new PowerPointOptions { ThrowOnMissingVariable = true });
            recipe.AddVariable("Products", new[] { new Product("Widget", 9.99m) });

            Action act = () => recipe.Generate();

            act.Should().Throw<Exception>(
                "an undefined property reached via #foreach must still throw under ThrowOnMissingVariable=true");
        }
        finally { Cleanup(tempPath); }
    }
}
