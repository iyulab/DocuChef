using DocuChef.Excel;
using DocuChef.Presentation;
using System.Reflection;

namespace DocuChef;

/// <summary>
/// Interface for all document template recipes
/// </summary>
public interface IRecipe : IDisposable
{
    /// <summary>
    /// Adds a variable to the template
    /// </summary>
    void AddVariable(string name, object value);

    /// <summary>
    /// Adds variables from a source object
    /// </summary>
    void AddVariable(object data);

    /// <summary>
    /// Clears all variables from the template
    /// </summary>
    void ClearVariables();

    /// <summary>
    /// Registers a global variable
    /// </summary>
    void RegisterGlobalVariable(string name, object value);
}

/// <summary>
/// Base implementation for document templates
/// </summary>
public abstract class RecipeBase : IRecipe
{
    protected readonly Dictionary<string, object> Variables = new();
    protected readonly Dictionary<string, Func<object>> GlobalVariables = new();
    protected bool IsDisposed;

    /// <summary>
    /// Adds variables from a source object
    /// </summary>
    public virtual void AddVariable(object data)
    {
        if (data == null)
            throw new ArgumentNullException(nameof(data)); if (data is IDictionary dictionary)
        {
            foreach (DictionaryEntry entry in dictionary)
            {
                var key = entry.Key?.ToString();
                if (key != null)
                {
                    AddVariable(key, entry.Value ?? string.Empty);
                }
            }
        }
        else
        {
            // Get all properties and fields using extension method
            var properties = data.GetProperties();
            foreach (var kvp in properties)
            {
                AddVariable(kvp.Key, kvp.Value);
            }
        }
    }

    /// <summary>
    /// Adds a variable to the template
    /// </summary>
    public abstract void AddVariable(string name, object value);

    /// <summary>
    /// Clears all variables from the template
    /// </summary>
    public virtual void ClearVariables()
    {
        Variables.Clear();
    }

    /// <summary>
    /// Registers a global variable
    /// </summary>
    public virtual void RegisterGlobalVariable(string name, object value)
    {
        if (string.IsNullOrEmpty(name))
            throw new ArgumentNullException(nameof(name));

        if (value is Func<object> valueFactory)
        {
            GlobalVariables[name] = valueFactory;
        }
        else
        {
            GlobalVariables[name] = () => value;
        }
    }

    /// <summary>
    /// Registers standard built-in global variables
    /// </summary>
    protected virtual void RegisterStandardGlobalVariables()
    {
        // Register date/time related variables
        RegisterGlobalVariable("Today", () => DateTime.Today);
        RegisterGlobalVariable("Now", () => DateTime.Now);
        RegisterGlobalVariable("Year", () => DateTime.Now.Year);
        RegisterGlobalVariable("Month", () => DateTime.Now.Month);
        RegisterGlobalVariable("Day", () => DateTime.Now.Day);

        // Register system variables
        RegisterGlobalVariable("MachineName", Environment.MachineName);
        RegisterGlobalVariable("UserName", Environment.UserName);
        RegisterGlobalVariable("OSVersion", Environment.OSVersion.ToString());
        RegisterGlobalVariable("ProcessorCount", Environment.ProcessorCount);
    }

    /// <summary>
    /// Disposes resources
    /// </summary>
    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    /// <summary>
    /// Disposes resources
    /// </summary>
    protected virtual void Dispose(bool disposing)
    {
        IsDisposed = true;
    }

    /// <summary>
    /// Throws an ObjectDisposedException if the object is disposed
    /// </summary>
    protected void ThrowIfDisposed([System.Runtime.CompilerServices.CallerMemberName] string memberName = "")
    {
        if (IsDisposed)
            throw new ObjectDisposedException(GetType().Name, $"Cannot access {memberName} after the object is disposed.");
    }
}

/// <summary>
/// Provides cooking-themed extension methods for recipes
/// </summary>
public static class RecipeExtensions
{
    /// <summary>
    /// Adds an ingredient (variable) to the recipe
    /// </summary>
    public static T AddIngredient<T>(this T recipe, string name, object value) where T : IRecipe
    {
        recipe.AddVariable(name, value);
        return recipe;
    }

    /// <summary>
    /// Adds ingredients (variables) from an object to the recipe
    /// </summary>
    public static T AddIngredients<T>(this T recipe, object data) where T : IRecipe
    {
        recipe.AddVariable(data);
        return recipe;
    }

    /// <summary>
    /// Clears all ingredients (variables) from the recipe
    /// </summary>
    public static T ClearIngredients<T>(this T recipe) where T : IRecipe
    {
        recipe.ClearVariables();
        return recipe;
    }

    /// <summary>
    /// Registers a cooking technique (function) for Excel recipes
    /// </summary>
    public static ExcelRecipe RegisterTechnique(this ExcelRecipe recipe, string name, Action<ClosedXML.Excel.IXLCell, object, string[]> function)
    {
        recipe.RegisterFunction(name, function);
        return recipe;
    }

    /// <summary>
    /// Cooks (generates) a recipe to the specified output path
    /// </summary>
    public static void Cook(this IRecipe recipe, string outputPath)
    {
        if (string.IsNullOrEmpty(outputPath))
            throw new ArgumentNullException(nameof(outputPath));

        if (recipe is ExcelRecipe excelRecipe)
        {
            CookExcelRecipe(excelRecipe, outputPath);
        }
        else if (recipe is PowerPointRecipe powerPointRecipe)
        {
            CookPowerPointRecipe(powerPointRecipe, outputPath);
        }
        else
        {
            throw new InvalidOperationException($"Recipe type {recipe.GetType().Name} is not supported");
        }
    }

    /// <summary>
    /// Cooks (generates) a recipe and returns the resulting dish
    /// </summary>
    public static IDish CookDish(this IRecipe recipe)
    {
        if (recipe is ExcelRecipe excelRecipe)
        {
            return excelRecipe.Generate();
        }
        else if (recipe is PowerPointRecipe powerPointRecipe)
        {
            return powerPointRecipe.Generate();
        }
        else
        {
            throw new InvalidOperationException($"Recipe type {recipe.GetType().Name} is not supported");
        }
    }

    /// <summary>
    /// Cooks (generates) an Excel recipe
    /// </summary>
    private static void CookExcelRecipe(ExcelRecipe recipe, string outputPath)
    {
        var document = recipe.Generate();
        document.SaveAs(outputPath);
    }

    /// <summary>
    /// Cooks (generates) a PowerPoint recipe
    /// </summary>
    private static void CookPowerPointRecipe(PowerPointRecipe recipe, string outputPath)
    {
        // Generate the document and save it - consistent with Excel approach
        var document = recipe.Generate();
        document.SaveAs(outputPath);
    }
}