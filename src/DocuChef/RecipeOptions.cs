using System.Globalization;
using ClosedXML.Report.XLCustom;
using DocuChef.Excel;
using DocuChef.Presentation;

namespace DocuChef;

/// <summary>
/// Options for document generation
/// </summary>
public class RecipeOptions
{
    /// <summary>
    /// Culture info for formatting numbers, dates, etc.
    /// </summary>
    public CultureInfo CultureInfo { get; set; } = CultureInfo.CurrentCulture;

    /// <summary>
    /// Excel-specific options
    /// </summary>
    public ExcelOptions? Excel { get; set; }

    /// <summary>
    /// PowerPoint-specific options
    /// </summary>
    public PowerPointOptions? PowerPoint { get; set; }

    /// <summary>
    /// Word-specific options (TBD)
    /// </summary>
    // public WordOptions Word { get; set; } = new WordOptions();

    /// <summary>
    /// Whether to enable verbose logging
    /// </summary>
    public bool EnableVerboseLogging { get; set; } = false;

    /// <summary>
    /// Whether to throw exceptions for missing variables instead of showing placeholders
    /// </summary>
    public bool ThrowOnMissingVariable { get; set; } = false;

    /// <summary>
    /// Maximum number of items to process in iterations (like foreach)
    /// </summary>
    public int MaxIterationItems { get; set; } = 1000;

    internal ExcelOptions GetExcelOptions()
    {
        Excel ??= new ExcelOptions()
        {
            EnableVerboseLogging = EnableVerboseLogging,
            ThrowOnMissingVariable = ThrowOnMissingVariable
        };
        return Excel;
    }

    internal PowerPointOptions GetPowerPointOptions()
    {
        PowerPoint ??= new PowerPointOptions()
        {
            EnableVerboseLogging = EnableVerboseLogging,
            ThrowOnMissingVariable = ThrowOnMissingVariable
        };
        return PowerPoint;
    }
}