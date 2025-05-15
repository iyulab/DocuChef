using DocuChef.PowerPoint;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using FluentAssertions;
using Xunit.Abstractions;

namespace DocuChef.Tests;

/// <summary>
/// Tests for PowerPoint #foreach directive slide generation functionality
/// </summary>
public class PPTForeachSlideCountTests : TestBase
{
    private readonly string _tempDirectory;

    public PPTForeachSlideCountTests(ITestOutputHelper output) : base(output)
    {
        _tempDirectory = Path.Combine(Path.GetTempPath(), "DocuChefTests", Guid.NewGuid().ToString());
        Directory.CreateDirectory(_tempDirectory);
    }

    public override void Dispose()
    {
        try { if (Directory.Exists(_tempDirectory)) Directory.Delete(_tempDirectory, true); }
        catch { }
        base.Dispose();
    }

    [Fact]
    public void Basic_Foreach_Creates_Correct_Number_Of_Slides()
    {
        // Arrange
        var chef = CreateNewChef();
        var templatePath = Path.Combine(_tempDirectory, "foreach_basic_template.pptx");

        // Create a simple template
        using (var presentationDoc = PPTHelper.CreateBasicPresentation(templatePath))
        {
            var slidePart = PPTHelper.AddSlide(presentationDoc);

            // Add a single array reference shape
            PPTHelper.AddTextShape(slidePart, "${Items[0].Name}", "ItemShape", 1, 1524000, 1524000, 6096000, 800000);

            // Add #foreach directive
            PPTHelper.AddNotesSlide(slidePart, "#foreach: Items");

            presentationDoc.Save();
        }

        // Create test data with specific number of items
        var items = new List<TestItem>();
        for (int i = 1; i <= 7; i++)
        {
            items.Add(new TestItem { Name = $"Item {i}" });
        }

        var recipe = chef.LoadPowerPointTemplate(templatePath);
        recipe.AddVariable("Items", items);

        // Act
        var document = recipe.Generate();
        var outputPath = Path.Combine(_tempDirectory, "foreach_basic_output.pptx");
        document.SaveAs(outputPath);

        // Assert
        using var resultDocument = PresentationDocument.Open(outputPath, false);
        var slideIds = resultDocument.PresentationPart.Presentation.SlideIdList.Elements<SlideId>().ToList();

        // Verify we have exactly 7 slides (one for each item)
        slideIds.Count.Should().Be(7, "Should generate one slide per item in the collection");
    }

    [Fact]
    public void Foreach_With_Max_Parameter_Creates_Correct_Number_Of_Slides()
    {
        // Arrange
        var chef = CreateNewChef();
        var templatePath = Path.Combine(_tempDirectory, "foreach_max_template.pptx");

        // Create template with max parameter
        using (var presentationDoc = PPTHelper.CreateBasicPresentation(templatePath))
        {
            var slidePart = PPTHelper.AddSlide(presentationDoc);

            // Add shapes for 3 items per slide
            PPTHelper.AddTextShape(slidePart, "${Items[0].Name}", "Item1", 1, 1524000, 1524000, 6096000, 800000);
            PPTHelper.AddTextShape(slidePart, "${Items[1].Name}", "Item2", 2, 1524000, 2524000, 6096000, 800000);
            PPTHelper.AddTextShape(slidePart, "${Items[2].Name}", "Item3", 3, 1524000, 3524000, 6096000, 800000);

            // Add #foreach with max: 3
            PPTHelper.AddNotesSlide(slidePart, "#foreach: Items, max: 3");

            presentationDoc.Save();
        }

        // Create test data with 10 items
        var items = new List<TestItem>();
        for (int i = 1; i <= 10; i++)
        {
            items.Add(new TestItem { Name = $"Item {i}" });
        }

        var recipe = chef.LoadPowerPointTemplate(templatePath);
        recipe.AddVariable("Items", items);

        // Act
        var document = recipe.Generate();
        var outputPath = Path.Combine(_tempDirectory, "foreach_max_output.pptx");
        document.SaveAs(outputPath);

        // Assert
        using var resultDocument = PresentationDocument.Open(outputPath, false);
        var slideIds = resultDocument.PresentationPart.Presentation.SlideIdList.Elements<SlideId>().ToList();

        // Verify we have 4 slides (ceiling(10/3) = 4)
        slideIds.Count.Should().Be(4, "Should generate 4 slides for 10 items with max=3");
    }

    [Fact]
    public void Foreach_With_Offset_Parameter_Creates_Correct_Number_Of_Slides()
    {
        // Arrange
        var chef = CreateNewChef();
        var templatePath = Path.Combine(_tempDirectory, "foreach_offset_template.pptx");

        // Create template with offset parameter
        using (var presentationDoc = PPTHelper.CreateBasicPresentation(templatePath))
        {
            var slidePart = PPTHelper.AddSlide(presentationDoc);

            // Add a single array reference
            PPTHelper.AddTextShape(slidePart, "${Items[0].Name}", "ItemShape", 1, 1524000, 1524000, 6096000, 800000);

            // Add #foreach with offset: 5
            PPTHelper.AddNotesSlide(slidePart, "#foreach: Items, offset: 5");

            presentationDoc.Save();
        }

        // Create test data with 10 items
        var items = new List<TestItem>();
        for (int i = 1; i <= 10; i++)
        {
            items.Add(new TestItem { Name = $"Item {i}" });
        }

        var recipe = chef.LoadPowerPointTemplate(templatePath);
        recipe.AddVariable("Items", items);

        // Act
        var document = recipe.Generate();
        var outputPath = Path.Combine(_tempDirectory, "foreach_offset_output.pptx");
        document.SaveAs(outputPath);

        // Assert
        using var resultDocument = PresentationDocument.Open(outputPath, false);
        var slideIds = resultDocument.PresentationPart.Presentation.SlideIdList.Elements<SlideId>().ToList();

        // Verify we have 5 slides (items 6-10, starting from offset 5)
        slideIds.Count.Should().Be(5, "Should generate 5 slides for items 6-10 with offset=5");
    }

    [Fact]
    public void Foreach_With_Max_And_Offset_Parameters_Creates_Correct_Number_Of_Slides()
    {
        // Arrange
        var chef = CreateNewChef();
        var templatePath = Path.Combine(_tempDirectory, "foreach_max_offset_template.pptx");

        // Create template with max and offset parameters
        using (var presentationDoc = PPTHelper.CreateBasicPresentation(templatePath))
        {
            var slidePart = PPTHelper.AddSlide(presentationDoc);

            // Add shapes for 2 items per slide
            PPTHelper.AddTextShape(slidePart, "${Items[0].Name}", "Item1", 1, 1524000, 1524000, 6096000, 800000);
            PPTHelper.AddTextShape(slidePart, "${Items[1].Name}", "Item2", 2, 1524000, 2524000, 6096000, 800000);

            // Add #foreach with max: 2 and offset: 3
            PPTHelper.AddNotesSlide(slidePart, "#foreach: Items, max: 2, offset: 3");

            presentationDoc.Save();
        }

        // Create test data with 12 items
        var items = new List<TestItem>();
        for (int i = 1; i <= 12; i++)
        {
            items.Add(new TestItem { Name = $"Item {i}" });
        }

        var recipe = chef.LoadPowerPointTemplate(templatePath);
        recipe.AddVariable("Items", items);

        // Act
        var document = recipe.Generate();
        var outputPath = Path.Combine(_tempDirectory, "foreach_max_offset_output.pptx");
        document.SaveAs(outputPath);

        // Assert
        using var resultDocument = PresentationDocument.Open(outputPath, false);
        var slideIds = resultDocument.PresentationPart.Presentation.SlideIdList.Elements<SlideId>().ToList();

        // Verify we have 5 slides ((12-3)/2 = 4.5 -> ceiling = 5)
        // Items 4-5, 6-7, 8-9, 10-11, 12
        slideIds.Count.Should().Be(5, "Should generate 5 slides for items 4-12 with max=2 and offset=3");
    }

    [Fact]
    public void Foreach_With_Empty_Collection_Creates_No_Slides()
    {
        // Arrange
        var chef = CreateNewChef();
        var templatePath = Path.Combine(_tempDirectory, "foreach_empty_template.pptx");

        // Create template
        using (var presentationDoc = PPTHelper.CreateBasicPresentation(templatePath))
        {
            var slidePart = PPTHelper.AddSlide(presentationDoc);

            // Add a reference
            PPTHelper.AddTextShape(slidePart, "${Items[0].Name}", "ItemShape", 1, 1524000, 1524000, 6096000, 800000);

            // Add #foreach directive
            PPTHelper.AddNotesSlide(slidePart, "#foreach: Items");

            presentationDoc.Save();
        }

        // Create empty test data
        var items = new List<TestItem>();

        var recipe = chef.LoadPowerPointTemplate(templatePath);
        recipe.AddVariable("Items", items);

        // Act
        var document = recipe.Generate();
        var outputPath = Path.Combine(_tempDirectory, "foreach_empty_output.pptx");
        document.SaveAs(outputPath);

        // Assert
        using var resultDocument = PresentationDocument.Open(outputPath, false);
        var slideIds = resultDocument.PresentationPart.Presentation.SlideIdList.Elements<SlideId>().ToList();

        // Verify we have 0 slides (empty collection = no slides)
        slideIds.Count.Should().Be(0, "Should generate no slides for empty collection");
    }

    // Simple test data class
    private class TestItem
    {
        public string Name { get; set; }
    }
}