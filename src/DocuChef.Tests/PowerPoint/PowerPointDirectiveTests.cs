using FluentAssertions;
using Xunit.Abstractions;

namespace DocuChef.Tests.PowerPoint;

/// <summary>
/// Black-box coverage for control directives, which live in slide notes and override the
/// engine's automatic pattern detection.
/// Contract source: SYNTAX_OF_PPT.md "Control Directives (In Slide Notes Only)".
/// </summary>
public class PowerPointDirectiveTests : TestBase
{
    public PowerPointDirectiveTests(ITestOutputHelper output) : base(output) { }

    private static Product[] SixProducts() =>
    [
        new Product("P0", 0m), new Product("P1", 1m), new Product("P2", 2m),
        new Product("P3", 3m), new Product("P4", 4m), new Product("P5", 5m)
    ];

    private List<string> Render(string slideText, string notes, string key, object data)
    {
        string tempPath = TempPptx();
        try
        {
            using var stream = PowerPointTestHelper.CreatePptxWithNotes((slideText, notes));
            File.WriteAllBytes(tempPath, stream.ToArray());

            using var chef = CreateNewChef();
            using var recipe = chef.LoadTemplate(tempPath);
            recipe.AddVariable(key, data);
            using var dish = recipe.Generate();

            var result = new MemoryStream();
            dish.SaveAs(result);

            var perSlide = PowerPointTestHelper.ReadTextBySlide(result);
            for (int i = 0; i < perSlide.Count; i++)
                _output.WriteLine($"slide[{i}] = {perSlide[i]}");
            return perSlide;
        }
        finally { Cleanup(tempPath); }
    }

    [Fact]
    public void ForeachDirective_WithMax_DrivesSlideCount()
    {
        var slides = Render(
            "${Products[0].Name}|${Products[1].Name}|${Products[2].Name}",
            "#foreach: Products, max: 3",
            "Products", SixProducts());

        slides.Should().HaveCount(2, "six items at three per slide need two slides");
        slides.Should().NotContain(s => s.Contains("${"),
            "no raw expression may reach the generated document");
        string all = string.Join("\n", slides);
        all.Should().Contain("P0").And.Contain("P5", "every item must appear somewhere");
    }

    [Fact]
    public void ForeachDirective_WithOffset_SkipsLeadingItems()
    {
        var slides = Render(
            "${Products[0].Name}|${Products[1].Name}|${Products[2].Name}",
            "#foreach: Products, max: 3, offset: 3",
            "Products", SixProducts());

        slides.Should().NotContain(s => s.Contains("${"));
        string all = string.Join("\n", slides);
        all.Should().Contain("P3", "offset 3 starts the window at the fourth item");
        all.Should().NotContain("P0", "items before the offset must not be rendered");
    }

    /// <summary>
    /// #range-begin / #range-end group slides for batch processing. It shares the directive
    /// path-resolution that hand-written #foreach also uses, so it is verified rather than
    /// assumed fixed.
    /// </summary>
    [Fact]
    public void RangeDirective_GroupsSlidesWithoutLeakingExpressions()
    {
        string tempPath = TempPptx();
        try
        {
            using var stream = PowerPointTestHelper.CreatePptxWithNotes(
                ("${Products[0].Name}", "#range-begin: Products"),
                ("${Products[0].Price:N2}", "#range-end: Products"));
            File.WriteAllBytes(tempPath, stream.ToArray());

            using var chef = CreateNewChef();
            using var recipe = chef.LoadTemplate(tempPath);
            recipe.AddVariable("Products", SixProducts());
            using var dish = recipe.Generate();

            var result = new MemoryStream();
            dish.SaveAs(result);

            var slides = PowerPointTestHelper.ReadTextBySlide(result);
            for (int i = 0; i < slides.Count; i++)
                _output.WriteLine($"slide[{i}] = {slides[i]}");

            slides.Should().NotContain(s => s.Contains("${"),
                "no raw expression may reach the generated document");
            slides.Should().NotBeEmpty("a range must still produce slides");
        }
        finally { Cleanup(tempPath); }
    }

    [Fact]
    public void AliasDirective_ShortensPathInExpressions()
    {
        var slides = Render(
            "${Items[0].Name}",
            "#alias: Products as Items",
            "Products", SixProducts());

        slides.Should().NotContain(s => s.Contains("${"),
            "an aliased path must resolve, not leak the raw expression");
        string.Join("\n", slides).Should().Contain("P0",
            "'Items' must resolve through the alias to 'Products'");
    }

    private static string TempPptx() =>
        Path.Combine(Path.GetTempPath(), $"docuchef_test_{Guid.NewGuid():N}.pptx");

    private static void Cleanup(string path)
    {
        if (File.Exists(path)) File.Delete(path);
    }
}
