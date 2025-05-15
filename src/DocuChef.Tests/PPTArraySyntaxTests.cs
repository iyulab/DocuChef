using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using FluentAssertions;
using System.Text;
using Xunit.Abstractions;

namespace DocuChef.Tests;

/// <summary>
/// Tests for PowerPoint array syntax
/// </summary>
public class PPTArraySyntaxTests : TestBase
{
    private readonly string _tempDirectory;
    private readonly string _templatePath;

    public PPTArraySyntaxTests(ITestOutputHelper output) : base(output)
    {
        _tempDirectory = Path.Combine(Path.GetTempPath(), "DocuChefTests", Guid.NewGuid().ToString());
        Directory.CreateDirectory(_tempDirectory);
        _templatePath = Path.Combine(_tempDirectory, "array_template.pptx");
        PPTHelper.CreateArraySyntaxTemplate(_templatePath);
    }

    public override void Dispose()
    {
        try { if (Directory.Exists(_tempDirectory)) Directory.Delete(_tempDirectory, true); }
        catch { }
        base.Dispose();
    }

    [Fact]
    public void Basic_Array_References_Resolve_Correctly()
    {
        // Arrange
        var chef = CreateNewChef();
        var recipe = chef.LoadPowerPointTemplate(_templatePath);

        var products = new List<Product>
        {
            new Product { Name = "Laptop", Price = 1299.99 },
            new Product { Name = "Smartphone", Price = 899.99 }
        };

        recipe.AddVariable("Products", products);

        // Act
        var document = recipe.Generate();
        var outputPath = Path.Combine(_tempDirectory, "basic_array_test.pptx");
        document.SaveAs(outputPath);

        // Assert
        using var presentationDocument = PresentationDocument.Open(outputPath, false);
        var slidePart = PPTHelper.GetFirstSlidePart(presentationDocument);
        var textElements = PPTHelper.GetTextElements(slidePart);

        textElements.Should().Contain(t => t.Contains("Product 1: Laptop"));
        textElements.Should().Contain(t => t.Contains("Product 2: Smartphone"));
        textElements.Should().Contain(t => t.Contains("$1299.99"));
        textElements.Should().Contain(t => t.Contains("$899.99"));
    }

    [Fact]
    public void Array_With_More_Items_Creates_Multiple_Slides()
    {
        // Arrange
        var chef = CreateNewChef();
        var recipe = chef.LoadPowerPointTemplate(_templatePath);

        var products = new List<Product>
        {
            new Product { Name = "Laptop", Price = 1299.99 },
            new Product { Name = "Smartphone", Price = 899.99 },
            new Product { Name = "Tablet", Price = 699.99 },
            new Product { Name = "Headphones", Price = 249.99 },
            new Product { Name = "Monitor", Price = 499.99 }
        };

        recipe.AddVariable("Products", products);
        recipe.AddVariable("ProductCount", products.Count); // 길이 정보를 명시적으로 제공

        // Act
        var document = recipe.Generate();
        var outputPath = Path.Combine(_tempDirectory, "array_multiple_slides_test.pptx");
        document.SaveAs(outputPath);

        // Assert
        using var presentationDocument = PresentationDocument.Open(outputPath, false);
        var slideCount = presentationDocument.PresentationPart.Presentation.SlideIdList.Count();

        // 5개 제품에 2개 항목씩 표시하면 최소 3개 슬라이드 필요
        slideCount.Should().BeGreaterThanOrEqualTo(3);

        // 모든 슬라이드의 내용을 검사하여 각 제품이 포함되어 있는지 확인
        var allSlideText = new StringBuilder();

        foreach (var slideId in presentationDocument.PresentationPart.Presentation.SlideIdList.Elements<SlideId>())
        {
            var slidePart = (SlidePart)presentationDocument.PresentationPart.GetPartById(slideId.RelationshipId);
            var textElements = PPTHelper.GetTextElements(slidePart);
            foreach (var text in textElements)
            {
                allSlideText.Append(text).Append(" ");
            }
        }

        var combinedText = allSlideText.ToString();

        // 모든 제품명이 어느 슬라이드에든 포함되어야 함
        combinedText.Should().Contain("Laptop");
        combinedText.Should().Contain("Smartphone");
        combinedText.Should().Contain("Tablet");
        combinedText.Should().Contain("Headphones");
        combinedText.Should().Contain("Monitor");
    }

    [Fact]
    public void Out_Of_Bounds_Array_References_Show_Empty()
    {
        // Arrange
        var chef = CreateNewChef();
        var recipe = chef.LoadPowerPointTemplate(_templatePath);

        // Only add one product
        var products = new List<Product>
        {
            new Product { Name = "Laptop", Price = 1299.99 }
        };

        recipe.AddVariable("Products", products);

        // Act
        var document = recipe.Generate();
        var outputPath = Path.Combine(_tempDirectory, "array_out_of_bounds_test.pptx");
        document.SaveAs(outputPath);

        // Assert
        using var presentationDocument = PresentationDocument.Open(outputPath, false);
        var slidePart = PPTHelper.GetFirstSlidePart(presentationDocument);
        var textElements = PPTHelper.GetTextElements(slidePart);

        // First product should be present
        textElements.Should().Contain(t => t.Contains("Product 1: Laptop"));

        // Second product reference should show empty or be hidden
        var product2Shape = PPTHelper.FindShapeByName(slidePart, "Product2Shape");
        if (product2Shape != null)
        {
            // Either the shape should be hidden
            bool? isHidden = product2Shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Hidden?.Value;

            // Or the text should be empty/default
            var product2Text = product2Shape.Descendants<DocumentFormat.OpenXml.Drawing.Text>()
                .Select(t => t.Text)
                .FirstOrDefault();

            (isHidden == true || product2Text == "Product 2:  - $").Should().BeTrue();
        }
    }

    private class Product
    {
        public string Name { get; set; }
        public double Price { get; set; }
    }
}