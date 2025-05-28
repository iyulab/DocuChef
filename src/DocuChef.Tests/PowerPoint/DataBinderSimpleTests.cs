using DocuChef.Presentation.Models;
using DocuChef.Presentation.Processors;
using FluentAssertions;
using System.Globalization;
using Xunit;
using Xunit.Abstractions;

namespace DocuChef.Tests.PowerPoint;

public class DataBinderSimpleTests : TestBase
{
    private readonly DataBinder _dataBinder;

    public DataBinderSimpleTests(ITestOutputHelper output) : base(output)
    {
        _dataBinder = new DataBinder();
    }

    [Fact]
    public void ResolveExpression_SimpleExpression_Works()
    {
        // Arrange
        var expression = new BindingExpression { OriginalExpression = "Name" };
        var data = new Dictionary<string, object> { ["Name"] = "Test Product" };

        // Act
        var result = _dataBinder.ResolveExpression(expression.OriginalExpression, data);

        // Assert
        result.Should().Be("Test Product");
    }

    [Fact]
    public void ResolveExpression_ContextOperator_Simple_Works()
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
            }
        };
        var data = new Dictionary<string, object> { ["Categories"] = categories };

        // Act
        var result = _dataBinder.ResolveExpression(expression.OriginalExpression, data);

        // Assert
        result.Should().Be("Smartphone");
    }

    [Theory]
    [InlineData("Categories>Items[0].Name", "Smartphone")]
    [InlineData("Categories>Items[1].Name", "Laptop")]
    [InlineData("Categories>Items[0].Price", 999)]
    [InlineData("Categories>Items[1].Price", 1299)]
    public void ResolveExpression_ContextOperator_VariousExpressions_Work(string expression, object expected)
    {
        // Arrange
        var bindingExpression = new BindingExpression 
        { 
            OriginalExpression = expression,
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
            }
        };
        var data = new Dictionary<string, object> { ["Categories"] = categories };        // Act
        var result = _dataBinder.ResolveExpression(bindingExpression.OriginalExpression, data);

        // Assert
        result.Should().Be(expected.ToString());
    }
}
