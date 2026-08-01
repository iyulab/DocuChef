using FluentAssertions;
using Xunit.Abstractions;

namespace DocuChef.Tests.PowerPoint;

/// <summary>
/// Black-box coverage for automatic array processing: a slide whose expressions
/// reference Collection[0..n] defines n+1 items per slide, and the engine clones
/// the slide as many times as the data requires, re-indexing expressions per slide.
/// Contract source: SYNTAX_OF_PPT.md "Automatic Array Processing".
/// </summary>
/// <remarks>
/// The data type must be public: expressions are evaluated by a compiled script,
/// which cannot see members of a private nested type.
/// </remarks>
public record Product(string Name, decimal Price);

public class PowerPointArrayBindingTests : TestBase
{
    public PowerPointArrayBindingTests(ITestOutputHelper output) : base(output) { }

    /// <summary>Runs a template through the engine and returns text per generated slide.</summary>
    private List<string> Render(string slideText, string dataKey, object data)
    {
        string tempPath = TempPptx();
        try
        {
            using var stream = PowerPointTestHelper.CreatePptx(slideText);
            File.WriteAllBytes(tempPath, stream.ToArray());

            using var chef = CreateNewChef();
            using var recipe = chef.LoadTemplate(tempPath);
            recipe.AddVariable(dataKey, data);
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
    public void ArrayBinding_ItemsFitOneSlide_BindsEachIndex()
    {
        var products = new[] { new Product("Alpha", 10m), new Product("Beta", 20m) };

        var slides = Render("${Products[0].Name}|${Products[1].Name}", "Products", products);

        slides.Should().HaveCount(1, "two slots and two items need exactly one slide");
        slides[0].Should().Contain("Alpha").And.Contain("Beta");
    }

    [Fact]
    public void ArrayBinding_MoreItemsThanSlots_ClonesSlides()
    {
        var products = new[]
        {
            new Product("P0", 0m), new Product("P1", 1m), new Product("P2", 2m),
            new Product("P3", 3m), new Product("P4", 4m)
        };

        var slides = Render("${Products[0].Name}|${Products[1].Name}", "Products", products);

        slides.Should().HaveCount(3,
            "2 items per slide over 5 items requires ceil(5/2) = 3 slides");
    }

    [Fact]
    public void ArrayBinding_MoreItemsThanSlots_ReindexesPerSlide()
    {
        var products = new[]
        {
            new Product("P0", 0m), new Product("P1", 1m), new Product("P2", 2m),
            new Product("P3", 3m), new Product("P4", 4m)
        };

        var slides = Render("${Products[0].Name}|${Products[1].Name}", "Products", products);

        slides[0].Should().Contain("P0").And.Contain("P1");
        slides[1].Should().Contain("P2").And.Contain("P3",
            "the second slide must advance the window, not repeat indices 0..1");
        slides[2].Should().Contain("P4");
    }

    [Fact]
    public void ArrayBinding_TrailingSlotWithoutData_RendersEmptyNotExpression()
    {
        var products = new[] { new Product("Only0", 0m), new Product("Only1", 1m), new Product("Only2", 2m) };

        var slides = Render("${Products[0].Name}|${Products[1].Name}", "Products", products);

        slides.Should().HaveCount(2);
        slides[1].Should().Contain("Only2");
        slides[1].Should().NotContain("${",
            "an out-of-range slot must resolve to empty, never leak the raw expression");
    }

    /// <summary>
    /// Control experiment for <see cref="ArrayBinding_TrailingSlotWithoutData_RendersEmptyNotExpression"/>:
    /// the same overflow, but with each expression in its own shape. If this passes while the
    /// single-shape variant fails, the trigger is shape-level granularity, not the overflow itself.
    /// </summary>
    [Fact]
    public void ArrayBinding_TrailingSlotWithoutData_SeparateShapes_KeepsInRangeItem()
    {
        var products = new[] { new Product("Only0", 0m), new Product("Only1", 1m), new Product("Only2", 2m) };

        string tempPath = TempPptx();
        try
        {
            using var stream = PowerPointTestHelper.CreatePptxWithShapes("${Products[0].Name}", "${Products[1].Name}");
            File.WriteAllBytes(tempPath, stream.ToArray());

            using var chef = CreateNewChef();
            using var recipe = chef.LoadTemplate(tempPath);
            recipe.AddVariable("Products", products);
            using var dish = recipe.Generate();

            var result = new MemoryStream();
            dish.SaveAs(result);

            var slides = PowerPointTestHelper.ReadTextBySlide(result);
            for (int i = 0; i < slides.Count; i++)
                _output.WriteLine($"slide[{i}] = {slides[i]}");

            slides.Should().HaveCount(2);
            slides[1].Should().Contain("Only2",
                "the in-range item must survive even though its sibling slot overflowed");
        }
        finally { Cleanup(tempPath); }
    }

    /// <summary>
    /// A spent slot must not survive as a dangling label: when every expression in a shape
    /// overflows, the shape goes away even though it still holds static text.
    /// </summary>
    [Fact]
    public void ArrayBinding_ExhaustedSlotWithStaticLabel_DropsWholeShape()
    {
        var products = new[] { new Product("Only0", 0m), new Product("Only1", 1m), new Product("Only2", 2m) };

        string tempPath = TempPptx();
        try
        {
            using var stream = PowerPointTestHelper.CreatePptxWithShapes(
                "Item 1: ${Products[0].Name}", "Item 2: ${Products[1].Name}");
            File.WriteAllBytes(tempPath, stream.ToArray());

            using var chef = CreateNewChef();
            using var recipe = chef.LoadTemplate(tempPath);
            recipe.AddVariable("Products", products);
            using var dish = recipe.Generate();

            var result = new MemoryStream();
            dish.SaveAs(result);

            var slides = PowerPointTestHelper.ReadTextBySlide(result);
            for (int i = 0; i < slides.Count; i++)
                _output.WriteLine($"slide[{i}] = {slides[i]}");

            slides.Should().HaveCount(2);
            slides[1].Should().Contain("Only2", "the in-range slot keeps its label and value");
            slides[1].Should().NotContain("Item 2:",
                "the slot with no data left must be removed, label included");
        }
        finally { Cleanup(tempPath); }
    }

    [Fact]
    public void ArrayBinding_FormatSpecifier_IsApplied()
    {
        var products = new[] { new Product("Alpha", 1299.5m) };

        var slides = Render("${Products[0].Price:N2}", "Products", products);

        slides[0].Should().Contain("1,299.50", "N2 formats with thousands separator and 2 decimals");
    }

    private static string TempPptx() =>
        Path.Combine(Path.GetTempPath(), $"docuchef_test_{Guid.NewGuid():N}.pptx");

    private static void Cleanup(string path)
    {
        if (File.Exists(path)) File.Delete(path);
    }
}
