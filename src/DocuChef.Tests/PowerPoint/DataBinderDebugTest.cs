using DocuChef.Presentation.Models;
using DocuChef.Presentation.Processors;
using FluentAssertions;
using System.Globalization;
using Xunit;
using Xunit.Abstractions;

namespace DocuChef.Tests.PowerPoint;

public class DataBinderDebugTest : TestBase
{
    private readonly DataBinder _dataBinder;

    public DataBinderDebugTest(ITestOutputHelper output) : base(output)
    {
        _dataBinder = new DataBinder();
    }

    [Fact]
    public void Debug_SimpleContextOperator_Test()
    {
        // Simple test case to debug the context operator
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
        var result = _dataBinder.ResolveExpression(expression.OriginalExpression, data);        // Debug output
        _output.WriteLine($"Expression: {expression.OriginalExpression}");
        _output.WriteLine($"Result: {result}");
        
        // For now, just check it doesn't crash
        result.Should().NotBeNull();
    }
}
