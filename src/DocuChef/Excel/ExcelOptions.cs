using ClosedXML.Report.XLCustom;

namespace DocuChef.Excel;

/// <summary>
/// Options for customizing Excel template processing
/// </summary>
public class ExcelOptions
{
    /// <summary>
    /// Options for the underlying XLCustomTemplate
    /// </summary>
    public XLCustomTemplateOptions TemplateOptions { get; set; } = new XLCustomTemplateOptions
    {
        UseGlobalRegistry = true,
        RegisterBuiltInFunctions = true
    };

    /// <summary>
    /// Whether to automatically register built-in functions
    /// </summary>
    public bool RegisterBuiltInFunctions { get; set; } = true;

    /// <summary>
    /// Whether to populate global variables
    /// </summary>
    public bool RegisterGlobalVariables { get; set; } = true;

    /// <summary>
    /// Whether to enable verbose logging for debugging
    /// </summary>
    public bool EnableVerboseLogging { get; set; } = false;

    /// <summary>
    /// Whether to throw exceptions for missing variables instead of showing placeholders
    /// </summary>
    public bool ThrowOnMissingVariable { get; set; } = false;
}