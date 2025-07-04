using System.Collections.Concurrent;
using System.Text.RegularExpressions;
using DocuChef.Presentation.Models;
using DocuChef.Presentation.Processors;
using DocuChef.Presentation.Context;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Drawing;
using DocuChef.Logging;
using DocuChef.Progress;

namespace DocuChef.Presentation;

/// <summary>
/// Processes PowerPoint templates and generates output documents
/// </summary>
public class PowerPointRecipe : IRecipe
{
    private readonly string? templatePath;
    private readonly PowerPointOptions options;
    private readonly MemoryStream? templateMemoryStream;
    private readonly Dictionary<string, object> variables = new Dictionary<string, object>();
    private readonly Dictionary<string, object> globalVariables = new Dictionary<string, object>();
    private readonly Dictionary<string, Func<object, object>> customFunctions = new Dictionary<string, Func<object, object>>();
    private object? dataObject;
    private DocuChef.Presentation.Functions.PPTFunctions? _pptFunctions;
    private ProgressCallback? _progressCallback;

    /// <summary>
    /// Creates a new PowerPoint recipe using a template file
    /// </summary>
    /// <param name="templatePath">Path to the template file</param>
    /// <param name="options">Options for processing</param>
    public PowerPointRecipe(string templatePath, PowerPointOptions options)
    {
        this.templatePath = templatePath;
        this.options = options ?? new PowerPointOptions();

        // Validate template path
        if (string.IsNullOrEmpty(templatePath) || !File.Exists(templatePath))
            throw new ArgumentException("Template file not found", nameof(templatePath));
    }

    /// <summary>
    /// Creates a new PowerPoint recipe using a template file with progress callback
    /// </summary>
    /// <param name="templatePath">Path to the template file</param>
    /// <param name="options">Options for processing</param>
    /// <param name="progressCallback">Progress callback</param>
    public PowerPointRecipe(string templatePath, PowerPointOptions options, ProgressCallback progressCallback)
        : this(templatePath, options)
    {
        _progressCallback = progressCallback;
    }

    /// <summary>
    /// Creates a new PowerPoint recipe using a template stream
    /// </summary>
    /// <param name="templateStream">Stream containing the template</param>
    /// <param name="powerPointOptions">Options for processing</param>
    public PowerPointRecipe(Stream templateStream, PowerPointOptions powerPointOptions)
    {
        if (templateStream == null)
            throw new ArgumentNullException(nameof(templateStream));

        this.options = powerPointOptions ?? new PowerPointOptions();

        // Copy the stream to memory for reuse
        templateMemoryStream = new MemoryStream();
        templateStream.CopyTo(templateMemoryStream);
        templateMemoryStream.Position = 0;
    }

    /// <summary>
    /// Creates a new PowerPoint recipe using a template stream with progress callback
    /// </summary>
    /// <param name="templateStream">Stream containing the template</param>
    /// <param name="powerPointOptions">Options for processing</param>
    /// <param name="progressCallback">Progress callback</param>
    public PowerPointRecipe(Stream templateStream, PowerPointOptions powerPointOptions, ProgressCallback progressCallback)
        : this(templateStream, powerPointOptions)
    {
        _progressCallback = progressCallback;
    }

    /// <summary>
    /// Adds a named variable to the recipe
    /// </summary>
    /// <param name="name">Variable name</param>
    /// <param name="value">Variable value</param>
    public void AddVariable(string name, object value)
    {
        if (string.IsNullOrEmpty(name))
            throw new ArgumentException("Variable name cannot be null or empty", nameof(name));

        variables[name] = value;
    }

    /// <summary>
    /// Adds an object as the main data source
    /// </summary>
    /// <param name="data">Data object</param>
    public void AddVariable(object data)
    {
        if (data == null)
            throw new ArgumentNullException(nameof(data));

        dataObject = data;
    }

    /// <summary>
    /// Registers a global variable
    /// </summary>
    /// <param name="name">Variable name</param>
    /// <param name="value">Variable value</param>
    public void RegisterGlobalVariable(string name, object value)
    {
        if (string.IsNullOrEmpty(name))
            throw new ArgumentException("Variable name cannot be null or empty", nameof(name));

        globalVariables[name] = value;
    }

    /// <summary>
    /// Sets the progress callback for tracking processing progress
    /// </summary>
    /// <param name="progressCallback">Progress callback</param>
    public void SetProgressCallback(ProgressCallback progressCallback)
    {
        _progressCallback = progressCallback;
    }

    /// <summary>
    /// Clears all variables
    /// </summary>
    public void ClearVariables()
    {
        variables.Clear();
        dataObject = null;
    }

    /// <summary>
    /// Disposes resources
    /// </summary>
    public void Dispose()
    {
        templateMemoryStream?.Dispose();
    }

    /// <summary>
    /// Generates the PowerPoint document using the new context-based processor
    /// </summary>
    /// <param name="outputPath">Path for the output file</param>
    /// <returns>Document result</returns>
    public IDish Cook(string outputPath)
    {
        if (string.IsNullOrEmpty(outputPath))
            throw new ArgumentException("Output path cannot be null or empty", nameof(outputPath));

        try
        {
            using var templateDocument = OpenTemplateDocument();
            var combinedData = CombineData();

            // Use the new context-based processor with progress reporting
            var processor = new ContextBasedPowerPointProcessor();
            var result = processor.ProcessPresentation(templateDocument, options, combinedData, _progressCallback);            // Save the working document to the output path
            if (result is PowerPointDocument pptDoc)
            {
                // Save the PowerPoint document to the specified output path
                pptDoc.SaveAs(outputPath);

                if (options.EnableVerboseLogging)
                {
                    Logger.Debug($"PowerPoint document saved to: {outputPath}");
                }
            }

            return result;
        }
        catch (Exception ex)
        {
            Logger.Error($"Error generating PowerPoint document: {ex.Message}", ex);
            throw;
        }
    }

    /// <summary>
    /// Generates the PowerPoint document and returns it
    /// </summary>
    /// <returns>Generated PowerPoint document</returns>
    public PowerPointDocument Generate()
    {
        try
        {
            using var templateDocument = OpenTemplateDocument();
            var combinedData = CombineData();

            // Use the new context-based processor with progress reporting
            var processor = new ContextBasedPowerPointProcessor();
            var result = processor.ProcessPresentation(templateDocument, options, combinedData, _progressCallback);

            if (result is PowerPointDocument pptDoc)
            {
                if (options.EnableVerboseLogging)
                {
                    Logger.Debug("PowerPoint document generated success fully");
                }
                return pptDoc;
            }
            else
            {
                throw new InvalidOperationException("Failed to generate PowerPoint document");
            }
        }
        catch (Exception ex)
        {
            Logger.Error($"Error generating PowerPoint document: {ex.Message}", ex);
            throw;
        }
    }

    /// <summary>
    /// Opens the template document for processing
    /// </summary>
    private PresentationDocument OpenTemplateDocument()
    {
        if (!string.IsNullOrEmpty(templatePath))
        {
            return PresentationDocument.Open(templatePath, false);
        }
        else if (templateMemoryStream != null)
        {
            templateMemoryStream.Position = 0;
            return PresentationDocument.Open(templateMemoryStream, false);
        }
        else
        {
            throw new InvalidOperationException("No template source available");
        }
    }

    /// <summary>
    /// Combines all data sources into a single dictionary
    /// </summary>
    private Dictionary<string, object> CombineData()
    {
        var combinedData = new Dictionary<string, object>(variables);

        // Add global variables
        foreach (var kvp in globalVariables)
        {
            combinedData[kvp.Key] = kvp.Value;
        }

        // Add properties from data object
        if (dataObject != null)
        {
            if (dataObject is IDictionary<string, object> dictData)
            {
                foreach (var kvp in dictData)
                {
                    combinedData[kvp.Key] = kvp.Value;
                }
            }
            else
            {
                var properties = dataObject.GetType().GetProperties();
                foreach (var prop in properties)
                {
                    if (prop.CanRead)
                    {
                        try
                        {
                            var propertyValue = prop.GetValue(dataObject);
                            if (propertyValue != null)
                            {
                                combinedData[prop.Name] = propertyValue;
                            }
                        }
                        catch (Exception ex)
                        {
                            if (options.EnableVerboseLogging)
                            {
                                Logger.Warning($"Could not read property {prop.Name}: {ex.Message}");
                            }
                        }
                    }
                }
            }
        }

        // Add PowerPoint functions object to make ppt.Image() etc. available
        if (_pptFunctions == null)
        {
            _pptFunctions = new DocuChef.Presentation.Functions.PPTFunctions(combinedData);
            Logger.Debug("CombineData: Created new PPTFunctions instance");
        }
        else
        {
            Logger.Debug("CombineData: Reusing existing PPTFunctions instance");
        }
        combinedData["ppt"] = _pptFunctions;

        return combinedData;
    }

    // NOTE: All data binding related methods have been removed
    // Data binding is now handled exclusively in DataBinder.cs via DollarSignEngine
    // The following methods are no longer needed:
    // - BindData()
    // - ScanAndReplaceExpressionsInSlide()
    // - ReplaceTextInSlide()
    // - ExtractExpressionsFromSlide()

    /// <summary>
    /// Processes image placeholders in the slide
    /// </summary>
    private void ProcessImagePlaceholders(SlidePart slidePart, DataBinder dataBinder)
    {
        if (slidePart?.Slide == null)
            return;

        try
        {
            // Find all text elements that might contain image placeholders
            var textElements = slidePart.Slide.Descendants<DocumentFormat.OpenXml.Drawing.Text>().ToList();

            foreach (var textElement in textElements)
            {
                if (string.IsNullOrEmpty(textElement.Text))
                    continue;

                // Look for ppt.Image() function calls
                if (textElement.Text.Contains("ppt.Image("))
                {
                    if (options.EnableVerboseLogging)
                    {
                        Logger.Debug($"Found image placeholder: {textElement.Text}");
                    }

                    // Process the image function
                    // Note: Image processing is handled by PowerPointFunctionHandler
                    // This is just for logging/debugging purposes
                }
            }
        }
        catch (Exception ex)
        {
            if (options.EnableVerboseLogging)
            {
                Logger.Warning($"Error processing image placeholders: {ex.Message}");
            }
        }
    }
}
