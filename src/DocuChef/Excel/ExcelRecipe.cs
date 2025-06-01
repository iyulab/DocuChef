using ClosedXML.Excel;
using ClosedXML.Report.XLCustom;
using DocuChef.Extensions;
using DocuChef.Logging;

namespace DocuChef.Excel;

/// <summary>
/// Represents an Excel template for document generation
/// </summary>
public class ExcelRecipe : RecipeBase
{
    private readonly ExcelOptions _options;
    private readonly XLCustomTemplate _template;
    private readonly string _templatePath;

    /// <summary>
    /// Creates a new Excel template from a file
    /// </summary>
    public ExcelRecipe(string templatePath, ExcelOptions? options = null)
    {
        if (string.IsNullOrEmpty(templatePath))
            throw new ArgumentNullException(nameof(templatePath));

        if (!File.Exists(templatePath))
            throw new FileNotFoundException("Template file not found", templatePath);

        _options = options ?? new ExcelOptions();
        _templatePath = templatePath;

        try
        {
            Logger.Debug($"Initializing Excel template from {templatePath}");
            _template = new XLCustomTemplate(templatePath, _options.TemplateOptions);

            InitializeTemplate();
        }
        catch (Exception ex)
        {
            Logger.Error($"Failed to initialize Excel template from {templatePath}", ex);
            throw new DocuChefException($"Failed to initialize Excel template: {ex.Message}", ex);
        }
    }

    /// <summary>
    /// Creates a new Excel template from a stream
    /// </summary>
    public ExcelRecipe(Stream templateStream, ExcelOptions? options = null)
    {
        if (templateStream == null)
            throw new ArgumentNullException(nameof(templateStream));

        _options = options ?? new ExcelOptions();

        try
        {
            // Save to a temporary file to preserve the template content
            _templatePath = Path.Combine(Path.GetTempPath(), $"DocuChef_{Guid.NewGuid():N}.xlsx");
            Logger.Debug($"Saving template stream to temporary file: {_templatePath}");

            using (var fileStream = new FileStream(_templatePath, FileMode.Create, FileAccess.Write))
            {
                templateStream.Position = 0;
                templateStream.CopyTo(fileStream);
            }

            Logger.Debug("Initializing Excel template from stream");
            _template = new XLCustomTemplate(_templatePath, _options.TemplateOptions);

            InitializeTemplate();
        }
        catch (Exception ex)
        {
            Logger.Error("Failed to initialize Excel template from stream", ex);
            throw new DocuChefException($"Failed to initialize Excel template from stream: {ex.Message}", ex);
        }
    }

    /// <summary>
    /// Initialize the template with built-in functions and global variables
    /// </summary>
    private void InitializeTemplate()
    {
        if (_options.RegisterBuiltInFunctions)
        {
            _template.RegisterBuiltIns();
            Logger.Debug("Registered built-in functions for Excel template");
        }

        if (_options.RegisterGlobalVariables)
        {
            RegisterStandardGlobalVariables();
            Logger.Debug("Registered global variables for Excel template");
        }
    }

    /// <summary>
    /// Adds a variable to the template
    /// </summary>
    public override void AddVariable(string name, object value)
    {
        ThrowIfDisposed();

        if (string.IsNullOrEmpty(name))
            throw new ArgumentNullException(nameof(name));

        try
        {
            Logger.Debug($"Adding variable '{name}' to Excel template");
            _template.AddVariable(name, value);
            Variables[name] = value;
        }
        catch (Exception ex)
        {
            Logger.Error($"Failed to add variable '{name}'", ex);
            throw new DocuChefException($"Failed to add variable '{name}': {ex.Message}", ex);
        }
    }

    /// <summary>
    /// Registers a custom function for cell processing
    /// </summary>
    public void RegisterFunction(string name, Action<IXLCell, object, string[]> function)
    {
        ThrowIfDisposed();

        if (string.IsNullOrEmpty(name))
            throw new ArgumentNullException(nameof(name));

        if (function == null)
            throw new ArgumentNullException(nameof(function));

        try
        {
            // Convert to XLFunctionHandler - which is what XLCustomTemplate.RegisterFunction expects
            XLFunctionHandler handler = (cell, value, parameters) => function(cell, value, parameters);
            _template.RegisterFunction(name, handler);
            Logger.Debug($"Registered function '{name}' for Excel template");
        }
        catch (Exception ex)
        {
            Logger.Error($"Failed to register function '{name}'", ex);
            throw new DocuChefException($"Failed to register function '{name}': {ex.Message}", ex);
        }
    }

    /// <summary>
    /// Generates the document from the template
    /// </summary>
    public ExcelDocument Generate()
    {
        ThrowIfDisposed();

        try
        {
            Logger.Debug("Generating Excel document from template");
            _template.Generate();

            // Get the workbook from the template
            var workbook = _template.Workbook;
            if (workbook == null)
            {
                throw new DocuChefException("Failed to retrieve workbook from template after generation.");
            }

            // Create a temporary file path for the generated workbook if needed
            string outputPath = Path.Combine(Path.GetTempPath(), $"DocuChef_{Guid.NewGuid():N}.xlsx");

            // Create document with file path reference
            var document = new ExcelDocument(workbook, outputPath);

            // Save to the temporary file so it can be opened later
            workbook.SaveAs(outputPath);

            Logger.Info("Excel document generated successfully");
            return document;
        }
        catch (Exception ex)
        {
            Logger.Error("Failed to generate Excel document", ex);
            throw new DocuChefException($"Failed to generate Excel document: {ex.Message}", ex);
        }
    }

    /// <summary>
    /// Generates the document directly to a file path
    /// </summary>
    public void GenerateToFile(string outputPath)
    {
        ThrowIfDisposed();

        try
        {
            // Create directory if it doesn't exist
            FileExtensions.EnsureDirectoryExists(outputPath); Logger.Debug($"Generating Excel document directly to: {outputPath}");
            _template.Generate();

            // Get the workbook from the template
            var workbook = _template.Workbook;
            if (workbook == null)
            {
                throw new DocuChefException("Failed to retrieve workbook from template after generation.");
            }

            // Save directly to the output path
            workbook.SaveAs(outputPath);

            Logger.Info($"Excel document generated successfully to: {outputPath}");
        }
        catch (Exception ex)
        {
            Logger.Error($"Failed to generate Excel document to {outputPath}", ex);
            throw new DocuChefException($"Failed to generate Excel document: {ex.Message}", ex);
        }
    }

    /// <summary>
    /// Disposes resources
    /// </summary>
    protected override void Dispose(bool disposing)
    {
        if (IsDisposed) return;

        if (disposing)
        {
            _template?.Dispose();

            // Delete temporary template file if it was created from a stream
            if (!string.IsNullOrEmpty(_templatePath) &&
                _templatePath.StartsWith(Path.GetTempPath()) &&
                File.Exists(_templatePath))
            {
                try
                {
                    File.Delete(_templatePath);
                    Logger.Debug($"Deleted temporary template file: {_templatePath}");
                }
                catch (Exception ex)
                {
                    Logger.Warning($"Failed to delete temporary template file: {ex.Message}");
                }
            }

            Logger.Debug("Excel template disposed");
        }

        base.Dispose(disposing);
    }
}