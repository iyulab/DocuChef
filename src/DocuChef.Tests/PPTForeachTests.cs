using DocuChef.PowerPoint.Processing.Directives;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using FluentAssertions;
using Xunit.Abstractions;

namespace DocuChef.Tests;

/// <summary>
/// Tests for PowerPoint #foreach directive and auto-detection functionality
/// </summary>
public class PPTForeachTests : TestBase
{
    private readonly string _tempDirectory;

    public PPTForeachTests(ITestOutputHelper output) : base(output)
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

        // Create a simple template with a Products[0].Name reference
        using (var presentationDoc = PPTHelper.CreateBasicPresentation(templatePath))
        {
            var slidePart = PPTHelper.AddSlide(presentationDoc);

            // Add shapes with array references
            PPTHelper.AddTextShape(slidePart,
                "Product: ${Products[0].Name}",
                "ProductNameShape", 1, 1524000, 1524000, 6096000, 800000);

            PPTHelper.AddTextShape(slidePart,
                "Price: ${Products[0].Price}",
                "ProductPriceShape", 2, 1524000, 2524000, 6096000, 800000);

            // Add #foreach directive to the slide notes
            PPTHelper.AddNotesSlide(slidePart, "#foreach: Products");

            presentationDoc.Save();
        }

        // Create test data with 5 products
        var products = new List<Product>
        {
            new Product { Name = "Product A", Price = 10.99m },
            new Product { Name = "Product B", Price = 20.50m },
            new Product { Name = "Product C", Price = 15.75m },
            new Product { Name = "Product D", Price = 8.25m },
            new Product { Name = "Product E", Price = 30.00m }
        };

        var recipe = chef.LoadPowerPointTemplate(templatePath);
        recipe.AddVariable("Products", products);

        // Act
        var document = recipe.Generate();
        var outputPath = Path.Combine(_tempDirectory, "foreach_basic_output.pptx");
        document.SaveAs(outputPath);

        // Assert
        using var resultDocument = PresentationDocument.Open(outputPath, false);

        // Verify we have 5 slides (one for each product)
        var slideIds = resultDocument.PresentationPart.Presentation.SlideIdList.Elements<SlideId>().ToList();
        slideIds.Count.Should().Be(5, "Each product should generate its own slide");

        // Verify content of each slide
        for (int i = 0; i < slideIds.Count; i++)
        {
            var slidePart = (SlidePart)resultDocument.PresentationPart.GetPartById(slideIds[i].RelationshipId);
            var slideText = string.Join(" ", PPTHelper.GetTextElements(slidePart));
            _output.WriteLine($"Slide {i + 1} text: '{slideText}'");

            // Each slide should contain the corresponding product name and price
            slideText.Should().Contain(products[i].Name, $"Slide {i + 1} should show {products[i].Name}");
            slideText.Should().Contain(products[i].Price.ToString(), $"Slide {i + 1} should show {products[i].Price}");
        }
    }

    [Fact]
    public void Auto_Detection_Without_Foreach_Creates_Correct_Number_Of_Slides()
    {
        // Arrange
        var chef = CreateNewChef();
        var templatePath = Path.Combine(_tempDirectory, "auto_detect_template.pptx");

        // Create a template WITHOUT #foreach directive
        using (var presentationDoc = PPTHelper.CreateBasicPresentation(templatePath))
        {
            var slidePart = PPTHelper.AddSlide(presentationDoc);

            // Add shapes with array references
            PPTHelper.AddTextShape(slidePart,
                "Product: ${Products[0].Name}",
                "ProductNameShape", 1, 1524000, 1524000, 6096000, 800000);

            PPTHelper.AddTextShape(slidePart,
                "Price: ${Products[0].Price}",
                "ProductPriceShape", 2, 1524000, 2524000, 6096000, 800000);

            // NO directive - auto-detection should work

            presentationDoc.Save();
        }

        // Create test data with 5 products
        var products = new List<Product>
        {
            new Product { Name = "Auto A", Price = 10.99m },
            new Product { Name = "Auto B", Price = 20.50m },
            new Product { Name = "Auto C", Price = 15.75m },
            new Product { Name = "Auto D", Price = 8.25m },
            new Product { Name = "Auto E", Price = 30.00m }
        };

        var recipe = chef.LoadPowerPointTemplate(templatePath);
        recipe.AddVariable("Products", products);

        // Act
        var document = recipe.Generate();
        var outputPath = Path.Combine(_tempDirectory, "auto_detect_output.pptx");
        document.SaveAs(outputPath);

        // Assert
        using var resultDocument = PresentationDocument.Open(outputPath, false);

        // Verify we have 5 slides (one for each product)
        var slideIds = resultDocument.PresentationPart.Presentation.SlideIdList.Elements<SlideId>().ToList();
        slideIds.Count.Should().Be(5, "Each product should generate its own slide even without #foreach directive");

        // Verify content of each slide
        for (int i = 0; i < slideIds.Count; i++)
        {
            var slidePart = (SlidePart)resultDocument.PresentationPart.GetPartById(slideIds[i].RelationshipId);
            var slideText = string.Join(" ", PPTHelper.GetTextElements(slidePart));
            _output.WriteLine($"Slide {i + 1} text: '{slideText}'");

            // Each slide should contain the corresponding product name and price
            slideText.Should().Contain(products[i].Name, $"Slide {i + 1} should show {products[i].Name}");
            slideText.Should().Contain(products[i].Price.ToString(), $"Slide {i + 1} should show {products[i].Price}");
        }
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
        var items = new List<Product>();

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

    [Fact]
    public void Auto_Detection_With_Empty_Collection_Creates_No_Slides()
    {
        // Arrange
        var chef = CreateNewChef();
        var templatePath = Path.Combine(_tempDirectory, "auto_detect_empty_template.pptx");

        // Create template without #foreach directive
        using (var presentationDoc = PPTHelper.CreateBasicPresentation(templatePath))
        {
            var slidePart = PPTHelper.AddSlide(presentationDoc);

            // Add a reference
            PPTHelper.AddTextShape(slidePart, "${Items[0].Name}", "ItemShape", 1, 1524000, 1524000, 6096000, 800000);

            // NO #foreach directive - should auto-detect

            presentationDoc.Save();
        }

        // Create empty test data
        var items = new List<Product>();

        var recipe = chef.LoadPowerPointTemplate(templatePath);
        recipe.AddVariable("Items", items);

        // Act
        var document = recipe.Generate();
        var outputPath = Path.Combine(_tempDirectory, "auto_detect_empty_output.pptx");
        document.SaveAs(outputPath);

        // Assert
        using var resultDocument = PresentationDocument.Open(outputPath, false);
        var slideIds = resultDocument.PresentationPart.Presentation.SlideIdList.Elements<SlideId>().ToList();

        // Verify we have 0 slides (empty collection = no slides)
        slideIds.Count.Should().Be(0, "Should generate no slides for empty collection even with auto-detection");
    }

    [Fact]
    public void Foreach_With_Max_Parameter_Creates_Correct_Number_Of_Slides()
    {
        // Arrange
        var chef = CreateNewChef();
        var templatePath = Path.Combine(_tempDirectory, "foreach_max_template.pptx");

        // Create template with max parameter in foreach directive
        using (var presentationDoc = PPTHelper.CreateBasicPresentation(templatePath))
        {
            var slidePart = PPTHelper.AddSlide(presentationDoc);

            // Add multiple items per slide
            PPTHelper.AddTextShape(slidePart,
                "Products List",
                "HeaderShape", 1, 1524000, 1024000, 6096000, 800000);

            PPTHelper.AddTextShape(slidePart,
                "1. ${Products[0].Name} - ${Products[0].Price:C2}",
                "Product1Shape", 2, 1524000, 2000000, 6096000, 800000);

            PPTHelper.AddTextShape(slidePart,
                "2. ${Products[1].Name} - ${Products[1].Price:C2}",
                "Product2Shape", 3, 1524000, 3000000, 6096000, 800000);

            PPTHelper.AddTextShape(slidePart,
                "3. ${Products[2].Name} - ${Products[2].Price:C2}",
                "Product3Shape", 4, 1524000, 4000000, 6096000, 800000);

            // Add #foreach directive with max parameter to the slide notes
            PPTHelper.AddNotesSlide(slidePart, "#foreach: Products, max: 3");

            presentationDoc.Save();
        }

        // Create test data with 8 products
        var products = new List<Product>();
        for (int i = 1; i <= 8; i++)
        {
            products.Add(new Product { Name = $"Product {i}", Price = i * 10.0m });
        }

        var recipe = chef.LoadPowerPointTemplate(templatePath);
        recipe.AddVariable("Products", products);

        // Act
        var document = recipe.Generate();
        var outputPath = Path.Combine(_tempDirectory, "foreach_max_output.pptx");
        document.SaveAs(outputPath);

        // Assert
        using var resultDocument = PresentationDocument.Open(outputPath, false);

        // Verify we have 3 slides (ceiling(8/3) = 3)
        var slideIds = resultDocument.PresentationPart.Presentation.SlideIdList.Elements<SlideId>().ToList();
        slideIds.Count.Should().Be(3, "We need 3 slides to display 8 products with 3 products per slide");

        // Log all slide texts for debugging
        for (int i = 0; i < slideIds.Count; i++)
        {
            var slidePart = (SlidePart)resultDocument.PresentationPart.GetPartById(slideIds[i].RelationshipId);
            var slideText = string.Join(" ", PPTHelper.GetTextElements(slidePart));
            _output.WriteLine($"Slide {i + 1} text: '{slideText}'");
        }

        // Verify first slide has products 1-3
        var slide1Part = (SlidePart)resultDocument.PresentationPart.GetPartById(slideIds[0].RelationshipId);
        var slide1Text = string.Join(" ", PPTHelper.GetTextElements(slide1Part));
        slide1Text.Should().Contain("Product 1");
        slide1Text.Should().Contain("Product 2");
        slide1Text.Should().Contain("Product 3");

        // Verify second slide has products 4-6
        var slide2Part = (SlidePart)resultDocument.PresentationPart.GetPartById(slideIds[1].RelationshipId);
        var slide2Text = string.Join(" ", PPTHelper.GetTextElements(slide2Part));
        slide2Text.Should().Contain("Product 4");
        slide2Text.Should().Contain("Product 5");
        slide2Text.Should().Contain("Product 6");

        // Verify third slide has products 7-8 (and no Product 9)
        var slide3Part = (SlidePart)resultDocument.PresentationPart.GetPartById(slideIds[2].RelationshipId);
        var slide3Text = string.Join(" ", PPTHelper.GetTextElements(slide3Part));
        slide3Text.Should().Contain("Product 7");
        slide3Text.Should().Contain("Product 8");
        slide3Text.Should().NotContain("Product 9"); // There is no Product 9
    }

    [Fact]
    public void Auto_Detection_With_Multiple_Indices_Creates_Correct_Number_Of_Slides()
    {
        // Arrange
        var chef = CreateNewChef();
        var templatePath = Path.Combine(_tempDirectory, "auto_detect_multiple_template.pptx");

        // Create template WITHOUT #foreach directive but with multiple indices
        using (var presentationDoc = PPTHelper.CreateBasicPresentation(templatePath))
        {
            var slidePart = PPTHelper.AddSlide(presentationDoc);

            // Add multiple items per slide (WITHOUT #foreach directive)
            PPTHelper.AddTextShape(slidePart,
                "Products List",
                "HeaderShape", 1, 1524000, 1024000, 6096000, 800000);

            PPTHelper.AddTextShape(slidePart,
                "1. ${Products[0].Name} - ${Products[0].Price:C2}",
                "Product1Shape", 2, 1524000, 2000000, 6096000, 800000);

            PPTHelper.AddTextShape(slidePart,
                "2. ${Products[1].Name} - ${Products[1].Price:C2}",
                "Product2Shape", 3, 1524000, 3000000, 6096000, 800000);

            PPTHelper.AddTextShape(slidePart,
                "3. ${Products[2].Name} - ${Products[2].Price:C2}",
                "Product3Shape", 4, 1524000, 4000000, 6096000, 800000);

            // NO directive - auto-detection should work with 3 items per slide

            presentationDoc.Save();
        }

        // Create test data with 8 products
        var products = new List<Product>();
        for (int i = 1; i <= 8; i++)
        {
            products.Add(new Product { Name = $"Auto Product {i}", Price = i * 10.0m });
        }

        var recipe = chef.LoadPowerPointTemplate(templatePath);
        recipe.AddVariable("Products", products);

        // Act
        var document = recipe.Generate();
        var outputPath = Path.Combine(_tempDirectory, "auto_detect_multiple_output.pptx");
        document.SaveAs(outputPath);

        // Assert
        using var resultDocument = PresentationDocument.Open(outputPath, false);

        // Verify we have 3 slides (ceiling(8/3) = 3)
        var slideIds = resultDocument.PresentationPart.Presentation.SlideIdList.Elements<SlideId>().ToList();
        slideIds.Count.Should().Be(3, "Should auto-detect 3 products per slide and create 3 slides for 8 products");

        // Log all slide texts for debugging
        for (int i = 0; i < slideIds.Count; i++)
        {
            var slidePart = (SlidePart)resultDocument.PresentationPart.GetPartById(slideIds[i].RelationshipId);
            var slideText = string.Join(" ", PPTHelper.GetTextElements(slidePart));
            _output.WriteLine($"Slide {i + 1} text: '{slideText}'");
        }

        // Verify appropriate product grouping on slides
        var slide1Part = (SlidePart)resultDocument.PresentationPart.GetPartById(slideIds[0].RelationshipId);
        var slide1Text = string.Join(" ", PPTHelper.GetTextElements(slide1Part));
        slide1Text.Should().Contain("Auto Product 1");
        slide1Text.Should().Contain("Auto Product 2");
        slide1Text.Should().Contain("Auto Product 3");

        var slide3Part = (SlidePart)resultDocument.PresentationPart.GetPartById(slideIds[2].RelationshipId);
        var slide3Text = string.Join(" ", PPTHelper.GetTextElements(slide3Part));
        slide3Text.Should().Contain("Auto Product 7");
        slide3Text.Should().Contain("Auto Product 8");
    }

    // Simple test data class
    private class Product
    {
        public string Name { get; set; }
        public decimal Price { get; set; }
    }
}