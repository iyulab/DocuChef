using DocuChef.PowerPoint.Processing;
using DocumentFormat.OpenXml.Packaging;

namespace DocuChef.PowerPoint;

/// <summary>
/// Represents a PowerPoint template for document generation using DollarSignEngine for expression evaluation
/// </summary>
public class PowerPointRecipe : RecipeBase
{
    private readonly PowerPointOptions _options;
    private readonly string _templatePath;
    private PresentationDocument _presentationDocument;

    /// <summary>
    /// Creates a new PowerPoint template from a file
    /// </summary>
    public PowerPointRecipe(string templatePath, PowerPointOptions options = null)
    {
        if (string.IsNullOrEmpty(templatePath))
            throw new ArgumentNullException(nameof(templatePath));

        if (!File.Exists(templatePath))
            throw new FileNotFoundException("Template file not found", templatePath);

        _templatePath = templatePath;
        _options = options ?? new PowerPointOptions();

        // Ensure text formatting is preserved by default
        _options.PreserveTextFormatting = true;

        Logger.Debug($"PowerPoint recipe initialized from file: {templatePath}");

        if (_options.RegisterBuiltInFunctions)
            RegisterBuiltInFunctions();

        if (_options.RegisterGlobalVariables)
            RegisterStandardGlobalVariables();
    }

    /// <summary>
    /// Creates a new PowerPoint template from a stream
    /// </summary>
    public PowerPointRecipe(Stream templateStream, PowerPointOptions options = null)
    {
        if (templateStream == null)
            throw new ArgumentNullException(nameof(templateStream));

        _options = options ?? new PowerPointOptions();

        // Create a temporary file to work with
        _templatePath = ".pptx".GetTempFilePath();

        try
        {
            Logger.Debug($"Creating temporary template file: {_templatePath}");
            templateStream.CopyToFile(_templatePath);

            if (_options.RegisterBuiltInFunctions)
                RegisterBuiltInFunctions();

            if (_options.RegisterGlobalVariables)
                RegisterStandardGlobalVariables();

            Logger.Debug("PowerPoint recipe initialized from stream");
        }
        catch (Exception ex)
        {
            Logger.Error("Failed to create temporary template file", ex);
            throw new DocuChefException($"Failed to create temporary template file: {ex.Message}", ex);
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

        Variables[name] = value;
        Logger.Debug($"Added variable: {name}");
    }

    /// <summary>
    /// Registers a custom function for PowerPoint processing
    /// </summary>
    public void RegisterFunction(PowerPointFunction function)
    {
        ThrowIfDisposed();

        if (function == null)
            throw new ArgumentNullException(nameof(function));

        if (string.IsNullOrEmpty(function.Name))
            throw new ArgumentException("Function name cannot be null or empty", nameof(function));

        if (function.Handler == null)
            throw new ArgumentException("Function handler cannot be null", nameof(function));

        var functionName = function.Name;
        Variables[$"ppt.{functionName}"] = function;
        Logger.Debug($"Registered function: {functionName}");
    }

    /// <summary>
    /// Registers built-in functions
    /// </summary>
    private void RegisterBuiltInFunctions()
    {
        // Register PowerPoint specific functions through PowerPointFunctions class
        PowerPointFunctions.RegisterBuiltInFunctions(this);
        Logger.Debug("Registered built-in PowerPoint functions");
    }

    /// <summary>
    /// Generates the document from the template
    /// </summary>
    public PowerPointDocument Generate()
    {
        ThrowIfDisposed();

        try
        {
            // Create a copy of the template to work with
            string outputPath = ".pptx".GetTempFilePath();
            Logger.Debug($"Creating output file: {outputPath}");

            File.Copy(_templatePath, outputPath, true);

            // Open the presentation
            _presentationDocument = PresentationDocument.Open(outputPath, true);
            Logger.Debug("Opened presentation document for editing");

            // Process the template with DollarSignEngine-based processor
            var processor = new PowerPointProcessor(_presentationDocument, _options);
            Logger.Info("Processing PowerPoint template with DollarSignEngine...");

            // Extract PowerPoint functions from variables
            var powerPointFunctions = Variables
                .Where(v => v.Key.StartsWith("ppt.") && v.Value is PowerPointFunction)
                .ToDictionary(
                    v => v.Key.Substring(4), // Remove "ppt." prefix
                    v => v.Value as PowerPointFunction
                );

            // Process the template
            processor.Process(Variables, GlobalVariables, powerPointFunctions);

            // Return the generated document
            Logger.Info("PowerPoint document generated successfully");
            return new PowerPointDocument(_presentationDocument, outputPath);
        }
        catch (Exception ex)
        {
            _presentationDocument?.Dispose();
            Logger.Error("Failed to generate PowerPoint document", ex);
            throw new DocuChefException($"Failed to generate PowerPoint document: {ex.Message}", ex);
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
            try
            {
                _presentationDocument?.Dispose();
                Logger.Debug("Presentation document disposed");

                // Delete temp file if created from stream and option is set
                if (_options.CleanupTemporaryFiles &&
                    _templatePath != null &&
                    _templatePath.Contains("DocuChef_") &&
                    File.Exists(_templatePath))
                {
                    File.Delete(_templatePath);
                    Logger.Debug($"Temporary file deleted: {_templatePath}");
                }
            }
            catch (Exception ex)
            {
                Logger.Error("Error disposing PowerPoint recipe resources", ex);
                // Ignore disposal errors
            }
        }

        base.Dispose(disposing);
    }
}