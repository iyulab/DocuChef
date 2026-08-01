using FluentAssertions;
using Xunit.Abstractions;

namespace DocuChef.Tests.PowerPoint;

/// <summary>
/// Black-box coverage for the context operator '&gt;'. Unlike dot notation, which always
/// resolves from the root, '&gt;' resolves against the parent item the slide was generated
/// for — so one template slide becomes one slide per parent.
/// Contract source: SYNTAX_OF_PPT.md "Contextual Hierarchy with '&gt;' Operator".
/// </summary>
public class PowerPointNestedContextTests : TestBase
{
    public PowerPointNestedContextTests(ITestOutputHelper output) : base(output) { }

    private static Category[] SampleCategories() =>
    [
        new Category("Electronics",
        [
            new Item("Smartphone", 999m), new Item("Laptop", 1299m), new Item("Tablet", 599m)
        ]),
        new Category("Furniture",
        [
            new Item("Sofa", 799m), new Item("Table", 499m)
        ])
    ];

    private List<string> Render(params string[] slideTexts)
    {
        string tempPath = TempPptx();
        try
        {
            using var stream = PowerPointTestHelper.CreatePptx(slideTexts);
            File.WriteAllBytes(tempPath, stream.ToArray());

            using var chef = CreateNewChef();
            using var recipe = chef.LoadTemplate(tempPath);
            recipe.AddVariable("Categories", SampleCategories());
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
    public void ContextOperator_OneSlidePerParent()
    {
        var slides = Render("${Categories[0].Name}");

        slides.Should().HaveCount(2, "two categories with one slot each need one slide per category");
        slides[0].Should().Contain("Electronics");
        slides[1].Should().Contain("Furniture");
    }

    /// <summary>
    /// The engine's nested model wants two template slides: a parent slide without '&gt;',
    /// and a child slide carrying the '&gt;' expressions.
    /// </summary>
    [Fact]
    public void ContextOperator_TwoSlideTemplate_ResolvesChildAgainstCurrentParent()
    {
        var slides = Render("${Categories[0].Name}", "${Categories>Items[0].Name}");

        slides.Should().NotContain(s => s.Contains("${"),
            "no raw template expression may reach the generated document");
        string all = string.Join("\n", slides);
        all.Should().Contain("Electronics").And.Contain("Furniture");
        all.Should().Contain("Smartphone").And.Contain("Sofa",
            "'>' must resolve against each parent in turn");
    }

    /// <summary>
    /// A template the engine cannot plan must still not ship raw syntax to the reader:
    /// unresolved expressions render empty (SYNTAX_OF_PPT.md "Empty Value Handling"),
    /// they are never passed through verbatim. Degrading must also not invent slides —
    /// an unplannable source slide is rendered once, not multiplied.
    /// </summary>
    [Fact]
    public void ContextOperator_SingleSlideTemplate_DegradesWithoutLeakingOrDuplicating()
    {
        var slides = Render("${Categories[0].Name}:${Categories>Items[0].Name}");

        slides.Should().NotContain(s => s.Contains("${"),
            "an unplannable nested template must degrade to empty output, not leak '${...}'");
        slides.Should().HaveCount(1,
            "one source slide the engine could not plan yields one rendered slide");
        slides[0].Should().Contain("Electronics",
            "the resolvable part of the template still binds");
    }

    private static string TempPptx() =>
        Path.Combine(Path.GetTempPath(), $"docuchef_test_{Guid.NewGuid():N}.pptx");

    private static void Cleanup(string path)
    {
        if (File.Exists(path)) File.Delete(path);
    }
}

/// <remarks>Public so the compiled binding expressions can see the members.</remarks>
public record Category(string Name, Item[] Items);

public record Item(string Name, decimal Price);
