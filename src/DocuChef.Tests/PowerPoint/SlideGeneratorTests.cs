using DocuChef.Presentation.Processors;
using DocuChef.Presentation.Models;
using DocuChef.Presentation.Exceptions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using Xunit;
using Xunit.Abstractions;

namespace DocuChef.Tests.PowerPoint;

public class SlideGeneratorTests : TestBase
{
    private readonly SlideGenerator _slideGenerator;

    public SlideGeneratorTests(ITestOutputHelper output) : base(output)
    {
        _slideGenerator = new SlideGenerator();
    }

    [Fact]
    public void GenerateSlides_WithValidSlidePlan_ShouldGenerateSlides()
    {
        // Arrange
        using var presentation = CreateMockPresentation();
        var slidePlan = new SlidePlan();
        slidePlan.SlideInstances.Add(new SlideInstance
        {
            SourceSlideId = 0,
            Position = 0,
            IndexOffset = 0
        });
        slidePlan.SlideInstances.Add(new SlideInstance
        {
            SourceSlideId = 0,
            Position = 1,
            IndexOffset = 1
        });

        // Act & Assert - Should not throw
        _slideGenerator.GenerateSlides(presentation, slidePlan);
        
        // Verify slides were created
        var slideIdList = presentation.PresentationPart?.Presentation.SlideIdList;
        Assert.NotNull(slideIdList);
        Assert.Equal(2, slideIdList.Count());
    }

    [Fact]
    public void GenerateSlides_WithNullPresentation_ShouldThrowArgumentNullException()
    {
        // Arrange
        var slidePlan = new SlidePlan();

        // Act & Assert
        Assert.Throws<ArgumentNullException>(() => 
            _slideGenerator.GenerateSlides(null!, slidePlan));
    }

    [Fact]
    public void GenerateSlides_WithNullSlidePlan_ShouldReturnWithoutAction()
    {
        // Arrange
        using var presentation = CreateMockPresentation();

        // Act - Should not throw
        _slideGenerator.GenerateSlides(presentation, null!);

        // Assert - Original slide count should remain unchanged
        var slideIdList = presentation.PresentationPart?.Presentation.SlideIdList;
        Assert.NotNull(slideIdList);
        Assert.Single(slideIdList); // Original template slide
    }

    [Fact]
    public void GenerateSlides_WithEmptySlidePlan_ShouldReturnWithoutAction()
    {
        // Arrange
        using var presentation = CreateMockPresentation();
        var slidePlan = new SlidePlan();

        // Act - Should not throw
        _slideGenerator.GenerateSlides(presentation, slidePlan);

        // Assert - Original slide count should remain unchanged
        var slideIdList = presentation.PresentationPart?.Presentation.SlideIdList;
        Assert.NotNull(slideIdList);
        Assert.Single(slideIdList); // Original template slide
    }

    [Fact]
    public void CloneSlide_WithValidSlideId_ShouldCloneSlide()
    {
        // Arrange
        using var presentation = CreateMockPresentation();

        // Act
        var clonedSlide = _slideGenerator.CloneSlide(presentation, 0);

        // Assert
        Assert.NotNull(clonedSlide);
        Assert.NotNull(clonedSlide.Slide);
    }

    [Fact]
    public void CloneSlide_WithNullPresentation_ShouldThrowArgumentNullException()
    {
        // Act & Assert
        Assert.Throws<ArgumentNullException>(() => 
            _slideGenerator.CloneSlide(null!, 0));
    }

    [Fact]
    public void CloneSlide_WithNegativeSlideId_ShouldThrowArgumentException()
    {
        // Arrange
        using var presentation = CreateMockPresentation();

        // Act & Assert
        Assert.Throws<ArgumentException>(() => 
            _slideGenerator.CloneSlide(presentation, -1));
    }

    [Fact]
    public void CloneSlide_WithInvalidSlideId_ShouldCreateRequiredSlides()
    {
        // Arrange
        using var presentation = CreateMockPresentation();
        
        // Act - Request slide ID beyond current count
        var clonedSlide = _slideGenerator.CloneSlide(presentation, 2);

        // Assert
        Assert.NotNull(clonedSlide);
        var slideIdList = presentation.PresentationPart?.Presentation.SlideIdList;
        Assert.NotNull(slideIdList);
        Assert.True(slideIdList.Count() >= 3); // Should have created slides up to ID 2
    }

    [Fact]
    public void AdjustBindingExpressions_WithZeroOffset_ShouldReturnOriginalExpressions()
    {
        // Arrange
        var expressions = new List<BindingExpression>
        {
            new BindingExpression
            {
                OriginalExpression = "${Categories[0].Name}",
                DataPath = "Categories[0].Name",
                ArrayIndices = new Dictionary<string, int> { { "Categories", 0 } }
            }
        };

        // Act
        var result = _slideGenerator.AdjustBindingExpressions(expressions, 0);

        // Assert
        Assert.Equal(expressions, result);
        Assert.Equal("${Categories[0].Name}", result[0].OriginalExpression);
        Assert.Equal(0, result[0].ArrayIndices["Categories"]);
    }

    [Fact]
    public void AdjustBindingExpressions_WithPositiveOffset_ShouldAdjustArrayIndices()
    {
        // Arrange
        var expressions = new List<BindingExpression>
        {
            new BindingExpression
            {
                OriginalExpression = "${Categories[0].Name}",
                DataPath = "Categories[0].Name",
                ArrayIndices = new Dictionary<string, int> { { "Categories", 0 } },
                UsesContextOperator = true
            }
        };

        // Act
        var result = _slideGenerator.AdjustBindingExpressions(expressions, 2);

        // Assert
        Assert.NotEqual(expressions, result); // Should be different instances
        Assert.Single(result);
        Assert.Equal(2, result[0].ArrayIndices["Categories"]);
        Assert.Contains("[2]", result[0].DataPath);
    }

    [Fact]
    public void AdjustBindingExpressions_WithComplexExpression_ShouldPreserveStructure()
    {
        // Arrange
        var expressions = new List<BindingExpression>
        {
            new BindingExpression
            {
                OriginalExpression = "${Categories[0].Items[1].Name}",
                DataPath = "Categories[0].Items[1].Name",
                FormatSpecifier = "C",
                IsConditional = false,
                IsMethodCall = false,
                UsesContextOperator = true,
                ArrayIndices = new Dictionary<string, int> 
                { 
                    { "Categories", 0 },
                    { "Items", 1 }
                }
            }
        };

        // Act
        var result = _slideGenerator.AdjustBindingExpressions(expressions, 1);

        // Assert
        Assert.Single(result);
        var adjusted = result[0];
        Assert.Equal("C", adjusted.FormatSpecifier);
        Assert.False(adjusted.IsConditional);
        Assert.False(adjusted.IsMethodCall);
        Assert.True(adjusted.UsesContextOperator);
        Assert.Equal(1, adjusted.ArrayIndices["Categories"]);
        Assert.Equal(2, adjusted.ArrayIndices["Items"]);
    }

    [Fact]
    public void AdjustSingleExpression_WithValidExpression_ShouldAdjustCorrectly()
    {
        // Arrange
        var expression = new BindingExpression
        {
            OriginalExpression = "${Items[0].Price}",
            DataPath = "Items[0].Price",
            ArrayIndices = new Dictionary<string, int> { { "Items", 0 } }
        };

        // Act
        var result = _slideGenerator.AdjustSingleExpression(expression, 3);

        // Assert
        Assert.NotNull(result);
        Assert.NotEqual(expression, result); // Should be different instance
        Assert.Equal(3, result.ArrayIndices["Items"]);
        Assert.Contains("[3]", result.DataPath);
    }

    [Fact]
    public void AdjustSingleExpression_WithNullExpression_ShouldReturnEmptyExpression()
    {
        // Act
        var result = _slideGenerator.AdjustSingleExpression(null!, 1);

        // Assert
        Assert.NotNull(result);
        Assert.Empty(result.ArrayIndices);
    }

    [Fact]
    public void AdjustSingleExpression_WithZeroOffset_ShouldReturnOriginal()
    {
        // Arrange
        var expression = new BindingExpression
        {
            OriginalExpression = "${Name}",
            DataPath = "Name"
        };

        // Act
        var result = _slideGenerator.AdjustSingleExpression(expression, 0);

        // Assert
        Assert.Equal(expression, result);
    }

    [Fact]
    public void InsertSlideAtPosition_WithValidSlide_ShouldInsertCorrectly()
    {
        // Arrange
        using var presentation = CreateMockPresentation();
        var slidePart = presentation.PresentationPart!.AddNewPart<SlidePart>();
        slidePart.Slide = new Slide();

        // Act
        _slideGenerator.InsertSlideAtPosition(presentation, slidePart, 0);

        // Assert
        var slideIdList = presentation.PresentationPart.Presentation.SlideIdList;
        Assert.NotNull(slideIdList);
        Assert.Equal(2, slideIdList.Count()); // Original + inserted
    }

    [Fact]
    public void InsertSlideAtPosition_BeyondCurrentCount_ShouldAppendSlide()
    {
        // Arrange
        using var presentation = CreateMockPresentation();
        var slidePart = presentation.PresentationPart!.AddNewPart<SlidePart>();
        slidePart.Slide = new Slide();

        // Act
        _slideGenerator.InsertSlideAtPosition(presentation, slidePart, 100);

        // Assert
        var slideIdList = presentation.PresentationPart.Presentation.SlideIdList;
        Assert.NotNull(slideIdList);
        Assert.Equal(2, slideIdList.Count()); // Should append, not create 100 slides
    }

    [Fact]
    public void ValidateSlideGeneration_WithNullPresentation_ShouldThrowArgumentNullException()
    {
        // Act & Assert
        Assert.Throws<ArgumentNullException>(() => 
            _slideGenerator.ValidateSlideGeneration(null!, 0));
    }

    [Fact]
    public void ValidateSlideGeneration_WithNegativeSlideId_ShouldThrowArgumentException()
    {
        // Arrange
        using var presentation = CreateMockPresentation();

        // Act & Assert
        var exception = Assert.Throws<ArgumentException>(() => 
            _slideGenerator.ValidateSlideGeneration(presentation, -1));
        Assert.Contains("must be non-negative", exception.Message);
    }

    [Fact]
    public void ValidateSlideGeneration_WithValidInputs_ShouldNotThrow()
    {
        // Arrange
        using var presentation = CreateMockPresentation();

        // Act & Assert - Should not throw
        _slideGenerator.ValidateSlideGeneration(presentation, 0);
    }

    [Fact]
    public void UpdateBindingExpressions_WithValidExpressions_ShouldUpdateSlideText()
    {
        // Arrange
        using var presentation = CreateMockPresentation();
        var slidePart = presentation.PresentationPart!.SlideParts.First();
        var adjustedExpressions = new List<BindingExpression>
        {
            new BindingExpression
            {
                OriginalExpression = "${Name}",
                DataPath = "Name"
            }
        };

        // Act - Should not throw
        _slideGenerator.UpdateBindingExpressions(slidePart, adjustedExpressions);

        // Assert - Method should complete without exception
        Assert.NotNull(slidePart.Slide);
    }

    [Fact]
    public void UpdateBindingExpressions_WithNullSlidePart_ShouldReturnWithoutError()
    {
        // Arrange
        var adjustedExpressions = new List<BindingExpression>
        {
            new BindingExpression { OriginalExpression = "${Name}", DataPath = "Name" }
        };

        // Act & Assert - Should not throw
        _slideGenerator.UpdateBindingExpressions(null!, adjustedExpressions);
    }

    [Fact]
    public void UpdateBindingExpressions_WithEmptyExpressions_ShouldReturnWithoutError()
    {
        // Arrange
        using var presentation = CreateMockPresentation();
        var slidePart = presentation.PresentationPart!.SlideParts.First();

        // Act & Assert - Should not throw
        _slideGenerator.UpdateBindingExpressions(slidePart, new List<BindingExpression>());
    }

    private PresentationDocument CreateMockPresentation()
    {
        var stream = new MemoryStream();
        var presentation = PresentationDocument.Create(stream, PresentationDocumentType.Presentation);
        
        var presentationPart = presentation.AddPresentationPart();
        presentationPart.Presentation = new DocumentFormat.OpenXml.Presentation.Presentation();
        presentationPart.Presentation.SlideIdList = new SlideIdList();

        // Add a template slide
        var slidePart = presentationPart.AddNewPart<SlidePart>();
        slidePart.Slide = new Slide();
        
        var slideId = new SlideId();
        slideId.Id = 256;
        slideId.RelationshipId = presentationPart.GetIdOfPart(slidePart);
        presentationPart.Presentation.SlideIdList.AppendChild(slideId);

        return presentation;
    }
}
