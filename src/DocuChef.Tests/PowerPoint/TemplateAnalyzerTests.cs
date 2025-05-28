using System.Collections.Generic;
using DocuChef.Presentation.Models;
using DocuChef.Presentation.Processors;
using FluentAssertions;
using Xunit;
using Xunit.Abstractions;

namespace DocuChef.Tests.PowerPoint;

/// <summary>
/// Tests for TemplateAnalyzer - Step 1: Template Analysis and SlideInfo Generation
/// Validates template scanning, directive parsing, and binding expression extraction
/// </summary>
public class TemplateAnalyzerTests : TestBase
{
    public TemplateAnalyzerTests(ITestOutputHelper output) : base(output) { }

    [Fact]
    public void Analyze_StaticSlideWithoutBindings_ReturnsStaticSlideInfo()
    {
        // Arrange
        var analyzer = new TemplateAnalyzer();
        var slideNotes = "";
        var mockSlidePart = CreateMockSlidePart("Title: Welcome to Our Presentation");

        // Act
        var result = analyzer.Analyze(mockSlidePart, slideNotes, 0);

        // Assert
        result.Should().NotBeNull();
        result.SlideId.Should().Be(0);
        result.Type.Should().Be(SlideType.Static);
        result.HasArrayReferences.Should().BeFalse();
        result.BindingExpressions.Should().BeEmpty();
        result.Directives.Should().BeEmpty();
    }

    [Fact]
    public void Analyze_SlideWithSimpleArrayBinding_DetectsArrayPattern()
    {
        // Arrange
        var analyzer = new TemplateAnalyzer();
        var slideNotes = "";
        var slideContent = "Product 1: ${Products[0].Name}\nProduct 2: ${Products[1].Name}";
        var mockSlidePart = CreateMockSlidePart(slideContent);

        // Act
        var result = analyzer.Analyze(mockSlidePart, slideNotes, 0);

        // Assert
        result.Should().NotBeNull();
        result.Type.Should().Be(SlideType.Source);
        result.HasArrayReferences.Should().BeTrue();
        result.MaxArrayIndex.Should().Be(1);
        result.ItemsPerSlide.Should().Be(2);
        result.CollectionName.Should().Be("Products");
        result.BindingExpressions.Should().HaveCount(2);
        result.BindingExpressions[0].OriginalExpression.Should().Be("${Products[0].Name}");
        result.BindingExpressions[1].OriginalExpression.Should().Be("${Products[1].Name}");
    }    [Fact]
    public void Analyze_SlideWithExplicitForeachDirective_ParsesDirectiveCorrectly()
    {
        // Arrange
        var analyzer = new TemplateAnalyzer();
        var slideNotes = "#foreach: Products, max: 1";  // PPT는 디자인 중심이므로 한 슬라이드당 하나의 항목
        var slideContent = "${Products[0].Name}";
        var mockSlidePart = CreateMockSlidePart(slideContent);

        // Act
        var result = analyzer.Analyze(mockSlidePart, slideNotes, 0);

        // Assert
        result.Should().NotBeNull();
        result.Directives.Should().HaveCount(1);
        result.Directives[0].Type.Should().Be(DirectiveType.Foreach);
        result.Directives[0].CollectionPath.Should().Be("Products");
        result.Directives[0].MaxItems.Should().Be(1);
    }

    [Fact]
    public void Analyze_SlideWithContextOperator_DetectsNestedContext()
    {
        // Arrange
        var analyzer = new TemplateAnalyzer();
        var slideNotes = "";
        var slideContent = "Category: ${Categories[0].Name}\nItem: ${Categories>Items[0].Name}";
        var mockSlidePart = CreateMockSlidePart(slideContent);

        // Act
        var result = analyzer.Analyze(mockSlidePart, slideNotes, 0);

        // Assert
        result.Should().NotBeNull();
        result.BindingExpressions.Should().HaveCount(2);
        
        var categoryExpression = result.BindingExpressions.First(e => e.OriginalExpression.Contains("Categories[0]"));
        categoryExpression.DataPath.Should().Be("Categories[0].Name");
        categoryExpression.ArrayIndices.Should().ContainKey("Categories");
        
        var itemExpression = result.BindingExpressions.First(e => e.OriginalExpression.Contains("Categories>Items"));
        itemExpression.DataPath.Should().Be("Categories>Items[0].Name");
        itemExpression.UsesContextOperator.Should().BeTrue();
    }

    [Fact]
    public void Analyze_SlideWithRangeDirectives_ParsesRangeBeginAndEnd()
    {
        // Arrange
        var analyzer = new TemplateAnalyzer();
        var slideNotes = "#range: begin, Categories\n#foreach: Categories";
        var slideContent = "${Categories[0].Name}";
        var mockSlidePart = CreateMockSlidePart(slideContent);

        // Act
        var result = analyzer.Analyze(mockSlidePart, slideNotes, 0);

        // Assert
        result.Should().NotBeNull();
        result.Directives.Should().HaveCount(2);
        
        var rangeDirective = result.Directives.First(d => d.Type == DirectiveType.Range);
        rangeDirective.CollectionPath.Should().Be("Categories");
        rangeDirective.RangeBoundary.Should().Be(RangeBoundary.Begin);
        
        var foreachDirective = result.Directives.First(d => d.Type == DirectiveType.Foreach);
        foreachDirective.CollectionPath.Should().Be("Categories");
    }

    [Fact]
    public void Analyze_SlideWithAliasDirective_ParsesAliasCorrectly()
    {
        // Arrange
        var analyzer = new TemplateAnalyzer();
        var slideNotes = "#alias: Categories>Products as Items";
        var slideContent = "${Items[0].Name}";
        var mockSlidePart = CreateMockSlidePart(slideContent);

        // Act
        var result = analyzer.Analyze(mockSlidePart, slideNotes, 0);

        // Assert
        result.Should().NotBeNull();
        result.Directives.Should().HaveCount(1);
        result.Directives[0].Type.Should().Be(DirectiveType.Alias);
        result.Directives[0].CollectionPath.Should().Be("Categories>Products");
        result.Directives[0].AliasName.Should().Be("Items");
    }

    [Fact]
    public void Analyze_SlideWithAutomaticDirectiveGeneration_GeneratesImplicitDirectives()
    {
        // Arrange
        var analyzer = new TemplateAnalyzer();
        var slideNotes = ""; // No explicit directives
        var slideContent = "${Items[0].Name}\n${Items[1].Name}\n${Items[2].Name}";
        var mockSlidePart = CreateMockSlidePart(slideContent);

        // Act
        var result = analyzer.Analyze(mockSlidePart, slideNotes, 0);

        // Assert
        result.Should().NotBeNull();
        result.Type.Should().Be(SlideType.Source);
        result.CollectionName.Should().Be("Items");
        result.MaxArrayIndex.Should().Be(2);
        result.ItemsPerSlide.Should().Be(3);
        
        // Should have automatically generated foreach directive
        result.Directives.Should().ContainSingle(d => 
            d.Type == DirectiveType.Foreach && 
            d.CollectionPath == "Items");
    }

    [Fact]
    public void Analyze_SlideWithMixedExpressions_ExtractsAllBindingExpressions()
    {
        // Arrange
        var analyzer = new TemplateAnalyzer();
        var slideNotes = "";
        var slideContent = @"
            Title: ${Report.Title}
            Date: ${Report.Date:yyyy-MM-dd}
            Items:
            - ${Items[0].Name}: ${Items[0].Price:C}
            - ${Items[1].Name}: ${Items[1].Price:C}
            Total: ${Items.Length} items
        ";
        var mockSlidePart = CreateMockSlidePart(slideContent);

        // Act
        var result = analyzer.Analyze(mockSlidePart, slideNotes, 0);

        // Assert
        result.Should().NotBeNull();
        result.BindingExpressions.Should().HaveCount(7);
        
        // Check format specifiers
        var dateExpression = result.BindingExpressions.First(e => e.OriginalExpression.Contains("Date:"));
        dateExpression.FormatSpecifier.Should().Be("yyyy-MM-dd");
        
        var priceExpressions = result.BindingExpressions.Where(e => e.OriginalExpression.Contains("Price:"));
        priceExpressions.Should().AllSatisfy(e => e.FormatSpecifier.Should().Be("C"));
    }

    private MockSlidePart CreateMockSlidePart(string content)
    {
        // Mock implementation - in real tests, this would create a proper SlidePart
        // For now, return a mock that contains the text content
        return new MockSlidePart { Content = content };
    }
}

// Mock class for testing - replace with proper DocumentFormat.OpenXml.Packaging.SlidePart in integration tests
public class MockSlidePart
{
    public string Content { get; set; } = string.Empty;
}
