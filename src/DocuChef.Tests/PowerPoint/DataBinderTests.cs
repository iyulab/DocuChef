using DocuChef.Presentation.Models;
using DocuChef.Presentation.Processors;
using FluentAssertions;
using System.Globalization;
using Xunit;
using Xunit.Abstractions;

namespace DocuChef.Tests.PowerPoint;

public class DataBinderTests : TestBase
{
    private readonly DataBinder _dataBinder;    public DataBinderTests(ITestOutputHelper output) : base(output)
    {
        _dataBinder = new DataBinder();
    }

    [Fact]
    public void ResolveExpression_SimpleExpression_ReturnsValue()
    {        // Arrange
        var expression = new BindingExpression { OriginalExpression = "Name" };
        var data = new Dictionary<string, object> { ["Name"] = "Test Product" };

        // Act
        var result = _dataBinder.ResolveExpression(expression.OriginalExpression, data);

        // Assert
        result.Should().Be("Test Product");
    }

    [Fact]
    public void ResolveExpression_NestedProperty_ReturnsValue()
    {
        // Arrange
        var expression = new BindingExpression { OriginalExpression = "Product.Name" };
        var product = new { Name = "Laptop", Price = 999 };
        var data = new Dictionary<string, object> { ["Product"] = product };

        // Act
        var result = _dataBinder.ResolveExpression(expression.OriginalExpression, data);

        // Assert
        result.Should().Be("Laptop");
    }

    [Fact]
    public void ResolveExpression_ArrayIndexAccess_ReturnsValue()
    {
        // Arrange
        var expression = new BindingExpression { OriginalExpression = "Products[0].Name" };
        var products = new[]
        {
            new { Name = "Laptop", Price = 999 },
            new { Name = "Mouse", Price = 25 }
        };
        var data = new Dictionary<string, object> { ["Products"] = products };

        // Act
        var result = _dataBinder.ResolveExpression(expression.OriginalExpression, data);

        // Assert
        result.Should().Be("Laptop");
    }

    [Fact]
    public void ResolveExpression_ContextOperator_ReturnsCorrectValue()
    {
        // Arrange
        var expression = new BindingExpression 
        { 
            OriginalExpression = "Categories>Items[0].Name",
            UsesContextOperator = true
        };
        
        var categories = new[]
        {
            new 
            { 
                Name = "Electronics",
                Items = new[]
                {
                    new { Name = "Smartphone", Price = 999 },
                    new { Name = "Laptop", Price = 1299 }
                }
            },
            new 
            { 
                Name = "Furniture",
                Items = new[]
                {
                    new { Name = "Sofa", Price = 799 },
                    new { Name = "Table", Price = 499 }
                }
            }
        };
        var data = new Dictionary<string, object> { ["Categories"] = categories };

        // Act
        var result = _dataBinder.ResolveExpression(expression.OriginalExpression, data);

        // Assert
        result.Should().Be("Smartphone");
    }

    [Fact]
    public void ResolveExpression_ContextOperator_WithComplexPath_ReturnsCorrectValue()
    {
        // Arrange
        var expression = new BindingExpression 
        { 
            OriginalExpression = "Departments>Members[1].Position",
            UsesContextOperator = true
        };
        
        var departments = new[]
        {
            new 
            { 
                Name = "Engineering",
                Manager = "John Smith",
                Members = new[]
                {
                    new { Name = "Alice", Position = "Senior Developer" },
                    new { Name = "Bob", Position = "Junior Developer" },
                    new { Name = "Carol", Position = "DevOps Engineer" }
                }
            }
        };
        var data = new Dictionary<string, object> { ["Departments"] = departments };

        // Act
        var result = _dataBinder.ResolveExpression(expression.OriginalExpression, data);

        // Assert
        result.Should().Be("Junior Developer");
    }

    [Fact]
    public void ResolveExpression_ContextOperator_MultipleNestingLevels_ReturnsCorrectValue()
    {
        // Arrange
        var expression = new BindingExpression 
        { 
            OriginalExpression = "Companies>Departments>Teams[0].Name",
            UsesContextOperator = true
        };
        
        var companies = new[]
        {
            new 
            { 
                Name = "TechCorp",
                Departments = new[]
                {
                    new 
                    {
                        Name = "Engineering",
                        Teams = new[]
                        {
                            new { Name = "Frontend Team", Size = 5 },
                            new { Name = "Backend Team", Size = 8 }
                        }
                    }
                }
            }
        };
        var data = new Dictionary<string, object> { ["Companies"] = companies };

        // Act
        var result = _dataBinder.ResolveExpression(expression.OriginalExpression, data);

        // Assert
        result.Should().Be("Frontend Team");
    }

    [Fact]
    public void ResolveExpression_EmptyExpression_ReturnsEmptyString()
    {
        // Arrange
        var expression = new BindingExpression { OriginalExpression = "" };
        var data = new Dictionary<string, object>();

        // Act
        var result = _dataBinder.ResolveExpression(expression.OriginalExpression, data);

        // Assert
        result.Should().Be(string.Empty);
    }    [Fact]
    public void ResolveExpression_NullExpression_ReturnsEmptyString()
    {
        // Arrange
        string? expression = null;
        var data = new Dictionary<string, object>();

        // Act
        var result = _dataBinder.ResolveExpression(expression!, data);

        // Assert
        result.Should().Be(string.Empty);
    }

    [Fact]
    public void ResolveExpression_NonExistentProperty_ReturnsEmptyOrNull()
    {
        // Arrange
        var expression = new BindingExpression { OriginalExpression = "NonExistentProperty" };
        var data = new Dictionary<string, object> { ["Name"] = "Test" };

        // Act
        var result = _dataBinder.ResolveExpression(expression.OriginalExpression, data);

        // Assert
        // DollarSignEngine should return null or empty for non-existent properties
        (result == null || result.ToString() == string.Empty).Should().BeTrue();
    }

    [Fact]
    public void ResolveExpression_StringExpression_ReturnsValue()
    {
        // Arrange
        var data = new Dictionary<string, object> { ["Name"] = "Test Product" };

        // Act
        var result = _dataBinder.ResolveExpression("Name", data);

        // Assert
        result.Should().Be("Test Product");
    }

    [Fact]
    public void ApplyIndexOffset_WithArrayIndices_CorrectlyOffsetsIndices()
    {
        // Arrange
        var expression = new BindingExpression
        {
            OriginalExpression = "Products[0].Name",
            DataPath = "Products[0].Name",
            ArrayIndices = new Dictionary<string, int> { ["Products"] = 0 }
        };

        // Act
        var result = _dataBinder.ApplyIndexOffset(expression, 2);

        // Assert
        result.OriginalExpression.Should().Be("Products[2].Name");
        result.DataPath.Should().Be("Products[2].Name");
        result.ArrayIndices["Products"].Should().Be(2);
    }

    [Fact]
    public void ApplyIndexOffset_WithZeroOffset_ReturnsUnchangedExpression()
    {
        // Arrange
        var expression = new BindingExpression
        {
            OriginalExpression = "Products[0].Name",
            DataPath = "Products[0].Name",
            ArrayIndices = new Dictionary<string, int> { ["Products"] = 0 }
        };

        // Act
        var result = _dataBinder.ApplyIndexOffset(expression, 0);

        // Assert
        result.OriginalExpression.Should().Be("Products[0].Name");
        result.DataPath.Should().Be("Products[0].Name");
        result.ArrayIndices["Products"].Should().Be(0);
    }

    [Fact]
    public void ResolveExpression_WithDollarSignSyntax_ReturnsCorrectValue()
    {
        // Arrange
        var expression = new BindingExpression { OriginalExpression = "${Name}" };
        var data = new Dictionary<string, object> { ["Name"] = "Test Product" };

        // Act
        var result = _dataBinder.ResolveExpression(expression.OriginalExpression, data);

        // Assert
        result.Should().Be("Test Product");
    }

    [Fact]
    public void ResolveExpression_ContextOperator_ConvertedToUnderscore_ReturnsCorrectValue()
    {
        // Arrange
        var expression = new BindingExpression 
        { 
            OriginalExpression = "Categories>Products[0].Name",
            UsesContextOperator = true
        };
        
        // Create test data that matches the expected context variable structure
        var categories = new[]
        {
            new 
            { 
                Name = "Electronics",
                Products = new[]
                {
                    new { Name = "Smartphone", Price = 999 },
                    new { Name = "Laptop", Price = 1299 }
                }
            }
        };
        var data = new Dictionary<string, object> { ["Categories"] = categories };

        // Act
        var result = _dataBinder.ResolveExpression(expression.OriginalExpression, data);

        // Assert
        result.Should().Be("Smartphone");
    }

    [Theory]
    [InlineData("Category>Items[0].Name", "Smartphone")]
    [InlineData("Category>Items[1].Price", 1299)]
    [InlineData("Category>Items[0].Price", 999)]
    public void ResolveExpression_ContextOperatorConversion_HandlesVariousExpressions(string expression, object expected)
    {
        // Arrange
        var bindingExpression = new BindingExpression 
        { 
            OriginalExpression = expression,
            UsesContextOperator = true
        };
        
        var category = new 
        { 
            Name = "Electronics",
            Items = new[]
            {
                new { Name = "Smartphone", Price = 999 },
                new { Name = "Laptop", Price = 1299 }
            }
        };
        var data = new Dictionary<string, object> { ["Category"] = category };

        // Act
        var result = _dataBinder.ResolveExpression(bindingExpression.OriginalExpression, data);        // Assert
        result.Should().Be(expected.ToString());
    }

    [Fact]
    public void ConvertToTemplate_ContextOperator_ConvertsToUnderscoreFormat()
    {
        // This test verifies that the context operator conversion works internally
        // We test this indirectly through the ResolveExpression method
        
        // Arrange
        var expression = new BindingExpression 
        { 
            OriginalExpression = "Root>Level1>Level2[0].Value",
            UsesContextOperator = true
        };
        
        var data = new Dictionary<string, object> 
        { 
            ["Root"] = new 
            {
                Level1 = new 
                {
                    Level2 = new[] { new { Value = "DeepValue" } }
                }
            }
        };

        // Act
        var result = _dataBinder.ResolveExpression(expression.OriginalExpression, data);

        // Assert
        result.Should().Be("DeepValue");
    }
}
