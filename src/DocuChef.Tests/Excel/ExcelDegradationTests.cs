using FluentAssertions;
using Xunit.Abstractions;

namespace DocuChef.Tests.Excel;

/// <summary>
/// Probes Excel for the failure shape repeatedly found in the PowerPoint pipeline:
/// when data is absent, does the engine degrade cleanly or ship template syntax?
/// </summary>
public class ExcelDegradationTests : TestBase
{
    public ExcelDegradationTests(ITestOutputHelper output) : base(output) { }

    private string RenderCell(string cellTemplate, string? key = null, object? data = null)
    {
        string tempPath = TempXlsx();
        try
        {
            using var stream = ExcelTestHelper.CreateXlsx((1, 1, cellTemplate));
            File.WriteAllBytes(tempPath, stream.ToArray());

            using var chef = CreateNewChef();
            using var recipe = chef.LoadTemplate(tempPath);
            if (key != null) recipe.AddVariable(key, data!);
            using var dish = recipe.Generate();

            var result = new MemoryStream();
            dish.SaveAs(result);

            string value = ExcelTestHelper.ReadCellValue(result, 1, 1);
            _output.WriteLine($"A1 = '{value}'");
            return value;
        }
        finally { Cleanup(tempPath); }
    }

    /// <summary>
    /// Excel delegates binding to XLCustomTemplate, which currently writes
    /// <c>Unknown identifier 'Missing'</c> into the cell — a diagnostic string that is
    /// indistinguishable from data once the file is delivered. Word and PowerPoint render
    /// empty instead. Only the guarantee DocuChef can make is asserted here; the upstream
    /// behaviour itself is tracked in claudedocs/upstream-issues/.
    /// </summary>
    [Fact]
    public void MissingVariable_DoesNotLeakTemplateSyntax()
    {
        string value = RenderCell("{{Missing}}");

        value.Should().NotContain("{{",
            "an unresolvable variable must not leave template syntax in the cell");
    }

    [Fact]
    public void PresentVariable_BindsNormally()
    {
        string value = RenderCell("{{Name}}", "Name", "Alice");

        value.Should().Be("Alice");
    }

    private static string TempXlsx() =>
        Path.Combine(Path.GetTempPath(), $"docuchef_test_{Guid.NewGuid():N}.xlsx");

    private static void Cleanup(string path)
    {
        if (File.Exists(path)) File.Delete(path);
    }
}
