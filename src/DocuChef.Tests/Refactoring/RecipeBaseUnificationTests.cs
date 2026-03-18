using DocuChef;
using DocuChef.Excel;
using DocuChef.Presentation;
using FluentAssertions;
using Xunit;

namespace DocuChef.Tests.Refactoring;

public class RecipeBaseUnificationTests
{
    [Fact]
    public void ExcelRecipe_Should_Inherit_RecipeBase()
    {
        typeof(ExcelRecipe).BaseType.Should().Be(typeof(RecipeBase));
    }

    [Fact]
    public void PowerPointRecipe_Should_Inherit_RecipeBase()
    {
        typeof(PowerPointRecipe).BaseType.Should().Be(typeof(RecipeBase));
    }

    [Fact]
    public void IRecipe_Should_Have_Generate_Method()
    {
        var method = typeof(IRecipe).GetMethod("Generate", Type.EmptyTypes);
        method.Should().NotBeNull();
        method!.ReturnType.Should().Be(typeof(IDish));
    }

    [Fact]
    public void IRecipe_Should_Have_Generate_With_OutputPath_Method()
    {
        var method = typeof(IRecipe).GetMethod("Generate", new[] { typeof(string) });
        method.Should().NotBeNull();
        method!.ReturnType.Should().Be(typeof(IDish));
    }

    [Fact]
    public void CookDish_Extension_Should_Work_Polymorphically()
    {
        var cookDishMethod = typeof(RecipeExtensions).GetMethod("CookDish");
        cookDishMethod.Should().NotBeNull();
        var parameters = cookDishMethod!.GetParameters();
        parameters.Should().HaveCount(1);
        parameters[0].ParameterType.Should().Be(typeof(IRecipe));
    }
}
