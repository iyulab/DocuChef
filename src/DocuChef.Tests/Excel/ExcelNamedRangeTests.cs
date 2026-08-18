using DocuChef.Excel;
using FluentAssertions;
using Xunit.Abstractions;

namespace DocuChef.Tests.Excel;

/// <summary>
/// Verifies that a named-range template row expands into one row per list item —
/// the core ClosedXML.Report list-binding feature ExcelRecipe delegates to. This surface
/// had zero test coverage even though the test helper (CreateXlsxWithNamedRange) already
/// existed for it.
/// </summary>
public class ExcelNamedRangeTests : TestBase
{
    public ExcelNamedRangeTests(ITestOutputHelper output) : base(output) { }

    public record Product(string Name, decimal Price);

    [Fact]
    public void NamedRange_ListBinding_ExpandsOneRowPerItem()
    {
        string tempPath = TempXlsx();
        try
        {
            using var stream = ExcelTestHelper.CreateXlsxWithNamedRange(
                "Products", templateRow: 1,
                (2, "{{item.Name}}"), (3, "{{item.Price}}"));
            File.WriteAllBytes(tempPath, stream.ToArray());

            var products = new List<Product>
            {
                new("Widget", 9.99m),
                new("Gadget", 19.99m),
                new("Gizmo", 29.99m),
            };

            using var chef = CreateNewChef();
            using var recipe = (ExcelRecipe)chef.LoadTemplate(tempPath);
            recipe.AddVariable("Products", products);
            using var dish = recipe.Generate();

            var resultStream = new MemoryStream();
            dish.SaveAs(resultStream);

            ExcelTestHelper.CountNonEmptyRows(resultStream).Should().Be(products.Count,
                "the named-range template row must expand into one row per list item");

            ExcelTestHelper.ReadCellValue(resultStream, 1, 2).Should().Be("Widget");
            ExcelTestHelper.ReadCellValue(resultStream, 2, 2).Should().Be("Gadget");
            ExcelTestHelper.ReadCellValue(resultStream, 3, 2).Should().Be("Gizmo");
        }
        finally { Cleanup(tempPath); }
    }

    [Fact]
    public void NamedRange_EmptyList_ProducesNoDataRows()
    {
        string tempPath = TempXlsx();
        try
        {
            using var stream = ExcelTestHelper.CreateXlsxWithNamedRange(
                "Products", templateRow: 1,
                (2, "{{item.Name}}"), (3, "{{item.Price}}"));
            File.WriteAllBytes(tempPath, stream.ToArray());

            using var chef = CreateNewChef();
            using var recipe = (ExcelRecipe)chef.LoadTemplate(tempPath);
            recipe.AddVariable("Products", new List<Product>());
            using var dish = recipe.Generate();

            var resultStream = new MemoryStream();
            dish.SaveAs(resultStream);

            ExcelTestHelper.CountNonEmptyRows(resultStream).Should().Be(0,
                "an empty list must not leave the template row or raw expression behind");
        }
        finally { Cleanup(tempPath); }
    }

    private static string TempXlsx() =>
        Path.Combine(Path.GetTempPath(), $"docuchef_test_{Guid.NewGuid():N}.xlsx");

    private static void Cleanup(string path)
    {
        if (File.Exists(path)) File.Delete(path);
    }
}
