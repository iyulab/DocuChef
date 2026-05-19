using ClosedXML.Excel;
using DocuChef.Excel;
using FluentAssertions;
using Xunit.Abstractions;

namespace DocuChef.Tests.Excel;

public class ExcelIntegrationTests : TestBase
{
    public ExcelIntegrationTests(ITestOutputHelper output) : base(output) { }

    [Fact]
    public void Chef_LoadTemplate_Xlsx_ReturnsExcelRecipe()
    {
        string tempPath = TempXlsx();
        try
        {
            using var stream = ExcelTestHelper.CreateXlsx((1, 1, "Hello"));
            File.WriteAllBytes(tempPath, stream.ToArray());

            using var chef = CreateNewChef();
            using var recipe = chef.LoadTemplate(tempPath);

            recipe.Should().BeOfType<ExcelRecipe>();
        }
        finally { Cleanup(tempPath); }
    }

    [Fact]
    public void ExcelRecipe_BasicVariable_Replaces()
    {
        string tempPath = TempXlsx();
        try
        {
            using var templateStream = ExcelTestHelper.CreateXlsx((1, 1, "{{Name}}"));
            File.WriteAllBytes(tempPath, templateStream.ToArray());

            using var chef = CreateNewChef();
            using var recipe = (ExcelRecipe)chef.LoadTemplate(tempPath);
            recipe.AddVariable("Name", "Alice");

            using var dish = recipe.Generate();

            var resultStream = new MemoryStream();
            dish.SaveAs(resultStream);

            var value = ExcelTestHelper.ReadCellValue(resultStream, 1, 1);
            value.Should().Be("Alice");
        }
        finally { Cleanup(tempPath); }
    }

    [Fact]
    public void ExcelRecipe_Generate_OutputPath_CreatesFile()
    {
        string tempPath = TempXlsx();
        string outputPath = TempXlsx();
        try
        {
            using var stream = ExcelTestHelper.CreateXlsx((1, 1, "{{Name}}"));
            File.WriteAllBytes(tempPath, stream.ToArray());

            using var chef = CreateNewChef();
            using var recipe = chef.LoadTemplate(tempPath);
            recipe.AddVariable("Name", "Bob");
            using var dish = recipe.Generate(outputPath);

            File.Exists(outputPath).Should().BeTrue();
        }
        finally
        {
            Cleanup(tempPath);
            Cleanup(outputPath);
        }
    }

    [Fact]
    public void ExcelDocument_SaveAs_FileExistsAfterDispose()
    {
        // Verifies the P0 bug fix: SaveAs output is NOT deleted by Dispose
        string tempPath = TempXlsx();
        string outputPath = TempXlsx();
        try
        {
            using var stream = ExcelTestHelper.CreateXlsx((1, 1, "{{Name}}"));
            File.WriteAllBytes(tempPath, stream.ToArray());

            using var chef = CreateNewChef();
            using var recipe = chef.LoadTemplate(tempPath);
            recipe.AddVariable("Name", "DisposalTest");

            using (var dish = recipe.Generate())
            {
                dish.SaveAs(outputPath);
            } // Dispose triggers here — file must survive

            File.Exists(outputPath).Should().BeTrue("SaveAs output must not be deleted by Dispose");
        }
        finally
        {
            Cleanup(tempPath);
            Cleanup(outputPath);
        }
    }

    [Fact]
    public void ExcelRecipe_GlobalVariables_Available()
    {
        string tempPath = TempXlsx();
        try
        {
            using var stream = ExcelTestHelper.CreateXlsx((1, 1, "{{UserName}}"));
            File.WriteAllBytes(tempPath, stream.ToArray());

            using var chef = CreateNewChef();
            using var recipe = chef.LoadTemplate(tempPath);
            using var dish = recipe.Generate();

            var resultStream = new MemoryStream();
            dish.SaveAs(resultStream);

            var value = ExcelTestHelper.ReadCellValue(resultStream, 1, 1);
            value.Should().Be(Environment.UserName);
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
