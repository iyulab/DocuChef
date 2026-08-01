using System.Globalization;
using ClosedXML.Report.XLCustom;
using DocuChef.Excel;
using DocuChef.Presentation;
using DocuChef.Word;

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
    /// Word-specific options
    /// </summary>
    public WordOptions? Word { get; set; }

    /// <summary>
    /// Whether to enable verbose logging
    /// </summary>
    public bool EnableVerboseLogging { get; set; } = false;

    /// <summary>
    /// Whether to throw exceptions for missing variables instead of showing placeholders.
    /// Honored by the Word and PowerPoint pipelines; see <see cref="ExcelOptions"/> for why
    /// the Excel path cannot currently act on it.
    /// </summary>
    public bool ThrowOnMissingVariable { get; set; } = false;

    internal ExcelOptions GetExcelOptions()
    {
        Excel ??= new ExcelOptions()
        {
            EnableVerboseLogging = EnableVerboseLogging,
            // Still propagated so the value round-trips for callers reading it back; the
            // Excel pipeline has no way to act on it until the template engine exposes one.
#pragma warning disable CS0618
            ThrowOnMissingVariable = ThrowOnMissingVariable
#pragma warning restore CS0618
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

    internal WordOptions GetWordOptions()
    {
        Word ??= new WordOptions()
        {
            EnableVerboseLogging = EnableVerboseLogging,
            ThrowOnMissingVariable = ThrowOnMissingVariable
        };
        return Word;
    }
}