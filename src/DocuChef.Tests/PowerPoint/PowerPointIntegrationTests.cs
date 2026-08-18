using DocuChef.Presentation;
using FluentAssertions;
using Xunit.Abstractions;

namespace DocuChef.Tests.PowerPoint;

public class PowerPointIntegrationTests : TestBase
{
    public PowerPointIntegrationTests(ITestOutputHelper output) : base(output) { }

    [Fact]
    public void Chef_LoadTemplate_Pptx_ReturnsPowerPointRecipe()
    {
        string tempPath = TempPptx();
        try
        {
            using var stream = PowerPointTestHelper.CreatePptx("Hello");
            File.WriteAllBytes(tempPath, stream.ToArray());

            using var chef = CreateNewChef();
            using var recipe = chef.LoadTemplate(tempPath);

            recipe.Should().BeOfType<PowerPointRecipe>();
        }
        finally { Cleanup(tempPath); }
    }

    [Fact]
    public void PowerPointRecipe_BasicVariable_Replaces()
    {
        string tempPath = TempPptx();
        try
        {
            using var stream = PowerPointTestHelper.CreatePptx("${Name}");
            File.WriteAllBytes(tempPath, stream.ToArray());

            using var chef = CreateNewChef();
            using var recipe = chef.LoadTemplate(tempPath);
            recipe.AddVariable("Name", "World");
            using var dish = recipe.Generate();

            var resultStream = new MemoryStream();
            dish.SaveAs(resultStream);

            var texts = PowerPointTestHelper.ReadAllText(resultStream);
            texts.Should().Contain(t => t.Contains("World"),
                "variable ${Name} should be replaced with 'World'");
        }
        finally { Cleanup(tempPath); }
    }

    [Fact]
    public void PowerPointRecipe_GlobalVariables_Available()
    {
        string tempPath = TempPptx();
        try
        {
            using var stream = PowerPointTestHelper.CreatePptx("${Today:yyyy-MM-dd}");
            File.WriteAllBytes(tempPath, stream.ToArray());

            using var chef = CreateNewChef();
            using var recipe = chef.LoadTemplate(tempPath);
            using var dish = recipe.Generate();

            var resultStream = new MemoryStream();
            dish.SaveAs(resultStream);

            var texts = PowerPointTestHelper.ReadAllText(resultStream);
            string expectedDate = DateTime.Today.ToString("yyyy-MM-dd");
            texts.Should().Contain(t => t.Contains(expectedDate),
                $"global variable ${{Today:yyyy-MM-dd}} should produce '{expectedDate}'");
        }
        finally { Cleanup(tempPath); }
    }

    [Fact]
    public void PowerPointRecipe_Cook_CreatesOutputFile()
    {
        string tempPath = TempPptx();
        string outputPath = TempPptx();
        try
        {
            using var stream = PowerPointTestHelper.CreatePptx("${Title}");
            File.WriteAllBytes(tempPath, stream.ToArray());

            using var chef = CreateNewChef();
            using var recipe = chef.LoadTemplate(tempPath);
            recipe.AddVariable("Title", "CookTest");
            recipe.Cook(outputPath);

            File.Exists(outputPath).Should().BeTrue();
        }
        finally
        {
            Cleanup(tempPath);
            Cleanup(outputPath);
        }
    }

    [Fact]
    public void Chef_PrepareDish_OneStep_CreatesOutputFile()
    {
        string tempPath = TempPptx();
        string outputPath = TempPptx();
        try
        {
            using var stream = PowerPointTestHelper.CreatePptx("${Title}");
            File.WriteAllBytes(tempPath, stream.ToArray());

            using var chef = CreateNewChef();
            chef.PrepareDish(tempPath, new { Title = "PrepareDishTest" }, outputPath);

            File.Exists(outputPath).Should().BeTrue();

            using var resultStream = new MemoryStream(File.ReadAllBytes(outputPath));
            var texts = PowerPointTestHelper.ReadAllText(resultStream);
            texts.Should().Contain(t => t.Contains("PrepareDishTest"),
                "PrepareDish should bind the supplied data just like the Culinary API composition it wraps");
        }
        finally
        {
            Cleanup(tempPath);
            Cleanup(outputPath);
        }
    }

    private static string TempPptx() =>
        Path.Combine(Path.GetTempPath(), $"docuchef_test_{Guid.NewGuid():N}.pptx");

    private static void Cleanup(string path)
    {
        if (File.Exists(path)) File.Delete(path);
    }
}
