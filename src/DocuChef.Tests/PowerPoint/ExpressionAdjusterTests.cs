using Xunit;
using Xunit.Abstractions;
using DocuChef.Presentation.Processors;
using DocuChef.Presentation.Models;
using System.Collections.Generic;

namespace DocuChef.Tests.PowerPoint;

public class ExpressionAdjusterTests : TestBase
{
    private readonly DataBinder _dataBinder;

    public ExpressionAdjusterTests(ITestOutputHelper output) : base(output)
    {
        _dataBinder = new DataBinder();
    }

    [Fact(DisplayName = "ApplyIndexOffset applies correct offset to array indices")]
    public void ApplyIndexOffset_AppliesCorrectOffset_ToArrayIndices()
    {
        // Arrange
        var expression = new BindingExpression
        {
            OriginalExpression = "${Products[0].Name}",
            DataPath = "Products[0].Name",
            ArrayIndices = new Dictionary<string, int> { { "Products", 0 } }
        };
        int indexOffset = 3;

        // Act
        var adjusted = _dataBinder.ApplyIndexOffset(expression, indexOffset);

        // Assert
        Assert.NotNull(adjusted);
        Assert.Equal(3, adjusted.ArrayIndices["Products"]);
        Assert.Contains("[3]", adjusted.DataPath);
        Assert.Contains("[3]", adjusted.OriginalExpression);
    }

    [Fact(DisplayName = "ApplyIndexOffset handles zero offset correctly")]
    public void ApplyIndexOffset_HandlesZeroOffset_Correctly()
    {
        // Arrange
        var expression = new BindingExpression
        {
            OriginalExpression = "${Items[5].Value}",
            DataPath = "Items[5].Value",
            ArrayIndices = new Dictionary<string, int> { { "Items", 5 } }
        };
        int indexOffset = 0;

        // Act
        var adjusted = _dataBinder.ApplyIndexOffset(expression, indexOffset);

        // Assert
        Assert.Equal(expression.OriginalExpression, adjusted.OriginalExpression);
        Assert.Equal(expression.DataPath, adjusted.DataPath);
        Assert.Equal(5, adjusted.ArrayIndices["Items"]);
    }

    [Fact(DisplayName = "ApplyIndexOffset handles negative offset correctly")]
    public void ApplyIndexOffset_HandlesNegativeOffset_Correctly()
    {
        // Arrange
        var expression = new BindingExpression
        {
            OriginalExpression = "${Categories[2].Products[1].Name}",
            DataPath = "Categories[2].Products[1].Name",
            ArrayIndices = new Dictionary<string, int> { { "Categories", 2 }, { "Products", 1 } }
        };
        int indexOffset = -1;

        // Act
        var adjusted = _dataBinder.ApplyIndexOffset(expression, indexOffset);

        // Assert - Negative offsets should return original expression
        Assert.Equal(expression.OriginalExpression, adjusted.OriginalExpression);
        Assert.Equal(expression.DataPath, adjusted.DataPath);
        Assert.Equal(2, adjusted.ArrayIndices["Categories"]);
        Assert.Equal(1, adjusted.ArrayIndices["Products"]);
    }

    [Fact(DisplayName = "ApplyIndexOffset handles multiple array indices correctly")]
    public void ApplyIndexOffset_HandlesMultipleArrayIndices_Correctly()
    {
        // Arrange
        var expression = new BindingExpression
        {
            OriginalExpression = "${Categories[0].Products[1].Name}",
            DataPath = "Categories[0].Products[1].Name",
            ArrayIndices = new Dictionary<string, int> { { "Categories", 0 }, { "Products", 1 } }
        };
        int indexOffset = 2;

        // Act
        var adjusted = _dataBinder.ApplyIndexOffset(expression, indexOffset);

        // Assert
        Assert.NotNull(adjusted);
        Assert.Equal(2, adjusted.ArrayIndices["Categories"]);
        Assert.Equal(3, adjusted.ArrayIndices["Products"]);
        Assert.Contains("[2]", adjusted.DataPath);
        Assert.Contains("[3]", adjusted.DataPath);
        Assert.Contains("[2]", adjusted.OriginalExpression);
        Assert.Contains("[3]", adjusted.OriginalExpression);
    }

    [Fact(DisplayName = "ApplyIndexOffset preserves non-array properties")]
    public void ApplyIndexOffset_PreservesNonArrayProperties_Correctly()
    {
        // Arrange
        var expression = new BindingExpression
        {
            OriginalExpression = "${Products[0].Name:C}",
            DataPath = "Products[0].Name",
            FormatSpecifier = "C",
            IsConditional = true,
            IsMethodCall = false,
            UsesContextOperator = true,
            ArrayIndices = new Dictionary<string, int> { { "Products", 0 } }
        };
        int indexOffset = 1;

        // Act
        var adjusted = _dataBinder.ApplyIndexOffset(expression, indexOffset);

        // Assert
        Assert.NotNull(adjusted);
        Assert.Equal("C", adjusted.FormatSpecifier);
        Assert.True(adjusted.IsConditional);
        Assert.False(adjusted.IsMethodCall);
        Assert.True(adjusted.UsesContextOperator);
        Assert.Equal(1, adjusted.ArrayIndices["Products"]);
    }

    [Fact(DisplayName = "ApplyIndexOffset creates deep clone without modifying original")]
    public void ApplyIndexOffset_CreatesDeepClone_WithoutModifyingOriginal()
    {
        // Arrange
        var originalExpression = new BindingExpression
        {
            OriginalExpression = "${Items[0].Value}",
            DataPath = "Items[0].Value",
            ArrayIndices = new Dictionary<string, int> { { "Items", 0 } }
        };
        int indexOffset = 5;

        // Act
        var adjusted = _dataBinder.ApplyIndexOffset(originalExpression, indexOffset);

        // Assert
        Assert.NotNull(adjusted);
        Assert.NotSame(originalExpression, adjusted);
        Assert.NotSame(originalExpression.ArrayIndices, adjusted.ArrayIndices);
        
        // Original should remain unchanged
        Assert.Equal("${Items[0].Value}", originalExpression.OriginalExpression);
        Assert.Equal("Items[0].Value", originalExpression.DataPath);
        Assert.Equal(0, originalExpression.ArrayIndices["Items"]);
        
        // Adjusted should have offset applied
        Assert.Contains("[5]", adjusted.DataPath);
        Assert.Equal(5, adjusted.ArrayIndices["Items"]);
    }    [Fact(DisplayName = "ApplyIndexOffset handles null expression gracefully")]
    public void ApplyIndexOffset_HandlesNullExpression_Gracefully()
    {
        // Arrange
        BindingExpression? expression = null;
        int indexOffset = 3;

        // Act
        var adjusted = _dataBinder.ApplyIndexOffset(expression!, indexOffset);

        // Assert
        Assert.NotNull(adjusted);
        Assert.Equal(string.Empty, adjusted.OriginalExpression ?? string.Empty);
        Assert.Equal(string.Empty, adjusted.DataPath ?? string.Empty);
    }

    [Fact(DisplayName = "ApplyIndexOffset handles expression without array indices")]
    public void ApplyIndexOffset_HandlesExpressionWithoutArrayIndices_Correctly()
    {
        // Arrange
        var expression = new BindingExpression
        {
            OriginalExpression = "${Company.Name}",
            DataPath = "Company.Name",
            ArrayIndices = new Dictionary<string, int>() // Empty array indices
        };
        int indexOffset = 2;

        // Act
        var adjusted = _dataBinder.ApplyIndexOffset(expression, indexOffset);

        // Assert
        Assert.NotNull(adjusted);
        Assert.Equal("${Company.Name}", adjusted.OriginalExpression);
        Assert.Equal("Company.Name", adjusted.DataPath);
        Assert.Empty(adjusted.ArrayIndices);
    }

    [Fact(DisplayName = "ApplyIndexOffset handles complex nested array paths")]
    public void ApplyIndexOffset_HandlesComplexNestedArrayPaths_Correctly()
    {
        // Arrange
        var expression = new BindingExpression
        {
            OriginalExpression = "${Categories[0].SubCategories[1].Products[2].Tags[0].Name}",
            DataPath = "Categories[0].SubCategories[1].Products[2].Tags[0].Name",
            ArrayIndices = new Dictionary<string, int>
            {
                { "Categories", 0 },
                { "SubCategories", 1 },
                { "Products", 2 },
                { "Tags", 0 }
            }
        };
        int indexOffset = 3;

        // Act
        var adjusted = _dataBinder.ApplyIndexOffset(expression, indexOffset);

        // Assert
        Assert.NotNull(adjusted);
        Assert.Equal(3, adjusted.ArrayIndices["Categories"]);
        Assert.Equal(4, adjusted.ArrayIndices["SubCategories"]);
        Assert.Equal(5, adjusted.ArrayIndices["Products"]);
        Assert.Equal(3, adjusted.ArrayIndices["Tags"]);
          Assert.Contains("[3]", adjusted.DataPath);
        Assert.Contains("[4]", adjusted.DataPath);
        Assert.Contains("[5]", adjusted.DataPath);
        Assert.Equal(4, adjusted.DataPath.Split('[').Length - 1); // Should have 4 array references
    }

    [Fact(DisplayName = "ApplyIndexOffset handles format specifiers in complex expressions")]
    public void ApplyIndexOffset_HandlesFormatSpecifiers_InComplexExpressions()
    {
        // Arrange
        var expression = new BindingExpression
        {
            OriginalExpression = "${Products[0].Price:C2}",
            DataPath = "Products[0].Price",
            FormatSpecifier = "C2",
            ArrayIndices = new Dictionary<string, int> { { "Products", 0 } }
        };
        int indexOffset = 7;

        // Act
        var adjusted = _dataBinder.ApplyIndexOffset(expression, indexOffset);

        // Assert
        Assert.NotNull(adjusted);
        Assert.Equal("C2", adjusted.FormatSpecifier);
        Assert.Equal(7, adjusted.ArrayIndices["Products"]);
        Assert.Contains("[7]", adjusted.DataPath);
        Assert.Contains("[7]", adjusted.OriginalExpression);
    }

    [Theory(DisplayName = "ApplyIndexOffset works with various index offset values")]
    [InlineData(1)]
    [InlineData(5)]
    [InlineData(10)]
    [InlineData(100)]
    public void ApplyIndexOffset_WorksWithVariousOffsetValues(int indexOffset)
    {
        // Arrange
        var expression = new BindingExpression
        {
            OriginalExpression = "${Items[2].Name}",
            DataPath = "Items[2].Name",
            ArrayIndices = new Dictionary<string, int> { { "Items", 2 } }
        };

        // Act
        var adjusted = _dataBinder.ApplyIndexOffset(expression, indexOffset);

        // Assert
        Assert.NotNull(adjusted);
        Assert.Equal(2 + indexOffset, adjusted.ArrayIndices["Items"]);
        Assert.Contains($"[{2 + indexOffset}]", adjusted.DataPath);
        Assert.Contains($"[{2 + indexOffset}]", adjusted.OriginalExpression);
    }
}
