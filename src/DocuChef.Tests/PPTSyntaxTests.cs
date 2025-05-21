//using Xunit.Abstractions;
//using FluentAssertions;
//using DocumentFormat.OpenXml.Packaging;

//namespace DocuChef.Tests;

///// <summary>
///// Tests for PowerPoint template syntax
///// </summary>
//public class PPTSyntaxTests : TestBase
//{
//    private readonly string _tempDirectory;
//    private readonly string _templatePath;

//    public PPTSyntaxTests(ITestOutputHelper output) : base(output)
//    {
//        _tempDirectory = Path.Combine(Path.GetTempPath(), "DocuChefTests", Guid.NewGuid().ToString());
//        Directory.CreateDirectory(_tempDirectory);
//        _templatePath = Path.Combine(_tempDirectory, "syntax_template.pptx");
//        PPTHelper.CreateBasicSyntaxTemplate(_templatePath);
//    }

//    public override void Dispose()
//    {
//        try { if (Directory.Exists(_tempDirectory)) Directory.Delete(_tempDirectory, true); }
//        catch { }
//        base.Dispose();
//    }

//    [Fact]
//    public void Basic_Expression_Interpolates_Values()
//    {
//        // Arrange
//        var chef = CreateNewChef();
//        var recipe = chef.LoadPowerPointTemplate(_templatePath);
//        recipe.AddVariable("Title", "Test Title");
//        recipe.AddVariable("Price", 1234.56);

//        // Act
//        var document = recipe.Generate();
//        var outputPath = Path.Combine(_tempDirectory, "basic_test.pptx");
//        document.SaveAs(outputPath);

//        // Assert
//        using var presentationDocument = PresentationDocument.Open(outputPath, false);
//        var slidePart = PPTHelper.GetFirstSlidePart(presentationDocument);
//        var textElements = PPTHelper.GetTextElements(slidePart);

//        textElements.Should().Contain(t => t == "Test Title");
//        textElements.Should().Contain(t => t.Contains("1,234.56"));
//    }

//    [Fact]
//    public void Object_Properties_Are_Accessed_Correctly()
//    {
//        // Arrange
//        var chef = CreateNewChef();
//        var recipe = chef.LoadPowerPointTemplate(_templatePath);

//        var product = new TestProduct
//        {
//            Name = "Premium Product",
//            Details = new ProductDetails { SKU = "ABC123" }
//        };

//        recipe.AddVariable("Product", product);

//        // Act
//        var document = recipe.Generate();
//        var outputPath = Path.Combine(_tempDirectory, "object_props_test.pptx");
//        document.SaveAs(outputPath);

//        // Assert
//        using var presentationDocument = PresentationDocument.Open(outputPath, false);
//        var slidePart = PPTHelper.GetFirstSlidePart(presentationDocument);
//        var textElements = PPTHelper.GetTextElements(slidePart);

//        textElements.Should().Contain(t => t.Contains("Premium Product"));
//        textElements.Should().Contain(t => t.Contains("ABC123"));
//    }

//    [Fact]
//    public void Conditional_Expressions_Evaluate_Correctly()
//    {
//        // Arrange
//        var chef = CreateNewChef();
//        var recipe = chef.LoadPowerPointTemplate(_templatePath);
//        recipe.AddVariable("InStock", true);
//        recipe.AddVariable("Quantity", 5);

//        // Act
//        var document = recipe.Generate();
//        var outputPath = Path.Combine(_tempDirectory, "conditional_test.pptx");
//        document.SaveAs(outputPath);

//        // Assert
//        using var presentationDocument = PresentationDocument.Open(outputPath, false);
//        var slidePart = PPTHelper.GetFirstSlidePart(presentationDocument);
//        var textElements = PPTHelper.GetTextElements(slidePart);

//        textElements.Should().Contain(t => t.Contains("In Stock"));
//        textElements.Should().Contain(t => t.Contains("Low Stock"));
//    }

//    [Fact]
//    public void If_Directive_Controls_Shape_Visibility()
//    {
//        // Arrange
//        var chef = CreateNewChef();
//        var recipe = chef.LoadPowerPointTemplate(_templatePath);
//        recipe.AddVariable("ShowElement", false);

//        // Act
//        var document = recipe.Generate();
//        var outputPath = Path.Combine(_tempDirectory, "directive_test.pptx");
//        document.SaveAs(outputPath);

//        // Assert
//        using var presentationDocument = PresentationDocument.Open(outputPath, false);
//        var slidePart = PPTHelper.GetFirstSlidePart(presentationDocument);
//        var testShape = PPTHelper.FindShapeByName(slidePart, "TestShape");

//        var isHidden = testShape?.NonVisualShapeProperties?.NonVisualDrawingProperties?.Hidden?.Value;
//        isHidden.Should().BeTrue();
//    }

//    private class TestProduct
//    {
//        public string Name { get; set; }
//        public ProductDetails Details { get; set; }
//    }

//    private class ProductDetails
//    {
//        public string SKU { get; set; }
//    }
//}