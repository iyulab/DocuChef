using DocuChef.Word;
using DocumentFormat.OpenXml.Packaging;
using FluentAssertions;
using Xunit;
using Xunit.Abstractions;
using WTable = DocumentFormat.OpenXml.Wordprocessing.Table;
using WTableRow = DocumentFormat.OpenXml.Wordprocessing.TableRow;

namespace DocuChef.Tests.Word;

public class WordIntegrationTests : TestBase
{
    public WordIntegrationTests(ITestOutputHelper output) : base(output)
    {
    }

    [Fact]
    public void Chef_LoadTemplate_Docx_ReturnsWordRecipe()
    {
        string tempPath = Path.Combine(Path.GetTempPath(), $"docuchef_test_{Guid.NewGuid()}.docx");
        try
        {
            using var stream = WordTestHelper.CreateDocx("Hello");
            File.WriteAllBytes(tempPath, stream.ToArray());

            using var chef = CreateNewChef();
            using var recipe = chef.LoadTemplate(tempPath);

            recipe.Should().BeOfType<WordRecipe>();
        }
        finally
        {
            if (File.Exists(tempPath)) File.Delete(tempPath);
        }
    }

    [Fact]
    public void WordRecipe_SimpleVariable_Generates()
    {
        string tempPath = Path.Combine(Path.GetTempPath(), $"docuchef_test_{Guid.NewGuid()}.docx");
        try
        {
            using var templateStream = WordTestHelper.CreateDocx("Hello ${Name}!");
            File.WriteAllBytes(tempPath, templateStream.ToArray());

            using var chef = CreateNewChef();
            using var recipe = chef.LoadTemplate(tempPath);
            recipe.AddVariable("Name", "World");
            using var dish = recipe.Generate();

            var resultStream = new MemoryStream();
            dish.SaveAs(resultStream);

            var texts = WordTestHelper.ReadParagraphTexts(resultStream);
            texts.Should().Contain(t => t.Contains("Hello World!"));
        }
        finally
        {
            if (File.Exists(tempPath)) File.Delete(tempPath);
        }
    }

    [Fact]
    public void WordRecipe_TableRepetition_Generates()
    {
        string tempPath = Path.Combine(Path.GetTempPath(), $"docuchef_test_{Guid.NewGuid()}.docx");
        try
        {
            using var templateStream = WordTestHelper.CreateDocxWithTable(
                headerTexts: new[] { "Name", "Price" },
                templateRowTexts: new[] { "${Items[].Name}", "${Items[].Price}" });
            File.WriteAllBytes(tempPath, templateStream.ToArray());

            var items = new List<Dictionary<string, object>>
            {
                new() { { "Name", "Apple" }, { "Price", "1.00" } },
                new() { { "Name", "Banana" }, { "Price", "0.50" } },
                new() { { "Name", "Cherry" }, { "Price", "2.00" } },
            };

            using var chef = CreateNewChef();
            using var recipe = chef.LoadTemplate(tempPath);
            recipe.AddVariable("Items", items);
            using var dish = recipe.Generate();

            var resultStream = new MemoryStream();
            dish.SaveAs(resultStream);

            var rows = WordTestHelper.ReadTableRows(resultStream);
            // header + 3 data rows
            rows.Should().HaveCount(4);
            rows[1][0].Should().Be("Apple");
            rows[2][0].Should().Be("Banana");
            rows[3][0].Should().Be("Cherry");
        }
        finally
        {
            if (File.Exists(tempPath)) File.Delete(tempPath);
        }
    }

    [Fact]
    public void WordRecipe_CookExtension_Works()
    {
        string tempPath = Path.Combine(Path.GetTempPath(), $"docuchef_test_{Guid.NewGuid()}.docx");
        string outputPath = Path.Combine(Path.GetTempPath(), $"docuchef_output_{Guid.NewGuid()}.docx");
        try
        {
            using var templateStream = WordTestHelper.CreateDocx("Hello ${Name}!");
            File.WriteAllBytes(tempPath, templateStream.ToArray());

            using var chef = CreateNewChef();
            using var recipe = chef.LoadTemplate(tempPath);
            recipe.AddVariable("Name", "CookTest");
            recipe.Cook(outputPath);

            File.Exists(outputPath).Should().BeTrue();

            using var fs = File.OpenRead(outputPath);
            var texts = WordTestHelper.ReadParagraphTexts(fs);
            texts.Should().Contain(t => t.Contains("Hello CookTest!"));
        }
        finally
        {
            if (File.Exists(tempPath)) File.Delete(tempPath);
            if (File.Exists(outputPath)) File.Delete(outputPath);
        }
    }

    [Fact]
    public void WordRecipe_GlobalVariables_Available()
    {
        string tempPath = Path.Combine(Path.GetTempPath(), $"docuchef_test_{Guid.NewGuid()}.docx");
        try
        {
            using var templateStream = WordTestHelper.CreateDocx("User: ${UserName}");
            File.WriteAllBytes(tempPath, templateStream.ToArray());

            using var chef = CreateNewChef();
            using var recipe = chef.LoadTemplate(tempPath);
            // Do NOT add UserName variable — it should come from RegisterStandardGlobalVariables
            using var dish = recipe.Generate();

            var resultStream = new MemoryStream();
            dish.SaveAs(resultStream);

            var texts = WordTestHelper.ReadParagraphTexts(resultStream);
            var expectedUserName = Environment.UserName;
            texts.Should().Contain(t => t.Contains($"User: {expectedUserName}"));
        }
        finally
        {
            if (File.Exists(tempPath)) File.Delete(tempPath);
        }
    }
}
