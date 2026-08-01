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
    /// Not honored on the Excel path. The Word and PowerPoint pipelines bind expressions
    /// themselves and respect this flag, but Excel delegates binding to XLCustomTemplate,
    /// whose options expose no equivalent setting — an unresolved identifier is written into
    /// the cell as a diagnostic string regardless of what is set here.
    /// </summary>
    [Obsolete("Not honored for Excel templates; the setting has no effect. Tracked upstream.")]
    public bool ThrowOnMissingVariable { get; set; } = false;
}