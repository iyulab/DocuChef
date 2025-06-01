using DocuChef.Excel;
using DocuChef.Presentation;

namespace DocuChef;

/// <summary>
/// Provides cooking-themed extension methods for DocuChef
/// </summary>
public static class ChefExtensions
{    /// <summary>
     /// Loads a template as a recipe
     /// </summary>
    public static IRecipe LoadRecipe(this Chef chef, string templatePath)
    {
        return chef.LoadTemplate(templatePath);
    }

    /// <summary>
    /// Loads an Excel template as a recipe
    /// </summary>
    public static ExcelRecipe LoadExcelRecipe(this Chef chef, string templatePath, ExcelOptions? options = null)
    {
        return chef.LoadExcelTemplate(templatePath, options);
    }

    /// <summary>
    /// Loads an Excel template from a stream as a recipe
    /// </summary>
    public static ExcelRecipe LoadExcelRecipe(this Chef chef, Stream templateStream, ExcelOptions? options = null)
    {
        return chef.LoadExcelTemplate(templateStream, options);
    }

    /// <summary>
    /// Loads a PowerPoint template as a recipe
    /// </summary>
    public static PowerPointRecipe LoadPowerPointRecipe(this Chef chef, string templatePath, PowerPointOptions? options = null)
    {
        return chef.LoadPowerPointTemplate(templatePath, options);
    }

    /// <summary>
    /// Loads a PowerPoint template from a stream as a recipe
    /// </summary>
    public static PowerPointRecipe LoadPowerPointRecipe(this Chef chef, Stream templateStream, PowerPointOptions? options = null)
    {
        return chef.LoadPowerPointTemplate(templateStream, options);
    }
}
