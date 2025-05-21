using DocuChef.Presentation.Core;
using DocuChef.Presentation.Directives;
using DocuChef.Presentation.Models;
using DocuChef.Presentation.Processing;
using DocuChef.Presentation.Utils;

namespace DocuChef.Presentation;

/// <summary>
/// Represents a PowerPoint template for document generation
/// </summary>
public class PowerPointRecipe : RecipeBase
{
    private readonly PowerPointOptions _options;
    private readonly string _templatePath;
    private readonly Stream _templateStream;
    private string _workingTemplatePath; // Path to the working copy in temp folder
    private TemplateProcessor _templateProcessor;
    private ContextProcessor _contextProcessor;
    private PlanProcessor _planProcessor;
    private PresentationProcessor _presentationProcessor;
    private List<SlideInfo> _templateSlides;
    private PresentationPlan _plan;

    // Direct data source reference for processing
    private object _effectiveDataSource;

    /// <summary>
    /// Template analysis results
    /// </summary>
    public TemplateAnalysisResult AnalysisResult { get; private set; }

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

        // Set current options for static access
        PowerPointOptions.SetCurrentOptions(_options);

        AnalysisResult = new TemplateAnalysisResult();

        try
        {
            Logger.Debug($"Initializing PowerPoint template from {templatePath}");

            // Create working copy in temp folder to avoid file access conflicts
            _workingTemplatePath = CreateWorkingCopy(_templatePath);

            InitializeRecipe();
        }
        catch (Exception ex)
        {
            CleanupWorkingCopy(); // Cleanup on error
            Logger.Error($"Failed to initialize PowerPoint template from {templatePath}", ex);
            throw new DocuChefException($"Failed to initialize PowerPoint template: {ex.Message}", ex);
        }
    }

    /// <summary>
    /// Creates a new PowerPoint template from a stream
    /// </summary>
    public PowerPointRecipe(Stream templateStream, PowerPointOptions options = null)
    {
        if (templateStream == null)
            throw new ArgumentNullException(nameof(templateStream));

        _templateStream = templateStream;
        _options = options ?? new PowerPointOptions();

        // Set current options for static access
        PowerPointOptions.SetCurrentOptions(_options);

        AnalysisResult = new TemplateAnalysisResult();

        try
        {
            Logger.Debug("Initializing PowerPoint template from stream");

            // Save stream to temporary file
            _workingTemplatePath = Path.Combine(Path.GetTempPath(), $"DocuChef_{Guid.NewGuid().ToString("N")}.pptx");
            using (var fileStream = new FileStream(_workingTemplatePath, FileMode.Create, FileAccess.Write))
            {
                _templateStream.Position = 0;
                _templateStream.CopyTo(fileStream);
            }

            InitializeRecipe();
        }
        catch (Exception ex)
        {
            CleanupWorkingCopy(); // Cleanup on error
            Logger.Error("Failed to initialize PowerPoint template from stream", ex);
            throw new DocuChefException($"Failed to initialize PowerPoint template: {ex.Message}", ex);
        }
    }

    /// <summary>
    /// Creates a working copy of the template file in a temporary location
    /// </summary>
    private string CreateWorkingCopy(string sourcePath)
    {
        string tempPath = Path.Combine(Path.GetTempPath(), $"DocuChef_Working_{Guid.NewGuid().ToString("N")}.pptx");

        try
        {
            // Copy with retry in case of file access conflicts
            CopyFileWithRetry(sourcePath, tempPath, 3);
            Logger.Debug($"Created working copy at: {tempPath}");
            return tempPath;
        }
        catch (Exception ex)
        {
            Logger.Error($"Failed to create working copy: {ex.Message}", ex);
            throw new DocuChefException($"Failed to create working copy of template: {ex.Message}", ex);
        }
    }

    /// <summary>
    /// Copies a file with retry logic for handling file access conflicts
    /// </summary>
    private void CopyFileWithRetry(string sourcePath, string destinationPath, int maxRetries)
    {
        int retryCount = 0;
        bool success = false;

        while (!success && retryCount < maxRetries)
        {
            try
            {
                File.Copy(sourcePath, destinationPath, true);
                success = true;
            }
            catch (IOException ex)
            {
                retryCount++;

                if (retryCount >= maxRetries)
                {
                    Logger.Error($"Failed to copy file after {maxRetries} attempts: {ex.Message}");
                    throw;
                }

                // Wait a bit before retrying
                Logger.Warning($"Retry {retryCount}/{maxRetries} - File access conflict, waiting before retry...");
                System.Threading.Thread.Sleep(500 * retryCount); // Incremental backoff
            }
        }
    }

    /// <summary>
    /// Cleans up the working copy if it exists
    /// </summary>
    private void CleanupWorkingCopy()
    {
        if (!string.IsNullOrEmpty(_workingTemplatePath) && File.Exists(_workingTemplatePath))
        {
            try
            {
                File.Delete(_workingTemplatePath);
                Logger.Debug($"Cleaned up working copy: {_workingTemplatePath}");
            }
            catch (Exception ex)
            {
                Logger.Warning($"Failed to clean up working copy: {ex.Message}");
            }
        }
    }

    /// <summary>
    /// Initialize the recipe processors
    /// </summary>
    private void InitializeRecipe()
    {
        // Initialize the processors
        _templateProcessor = new TemplateProcessor();
        _planProcessor = new PlanProcessor();
        _presentationProcessor = new PresentationProcessor();
        _templateSlides = new List<SlideInfo>();

        // Normalize template quotes for consistent processing
        PPTUtils.NormalizeTemplateQuotes(_workingTemplatePath);

        // Initialize with empty object as data source
        _effectiveDataSource = new object();
        _contextProcessor = new ContextProcessor(_effectiveDataSource);

        if (_options.AnalyzeOnInit)
        {
            AnalyzeTemplate();
        }

        if (_options.RegisterGlobalVariables)
        {
            RegisterStandardGlobalVariables();
            Logger.Debug("Registered global variables for PowerPoint template");
        }
    }

    /// <summary>
    /// Analyzes the template to understand its structure
    /// </summary>
    public TemplateAnalysisResult AnalyzeTemplate()
    {
        Logger.Info($"Analyzing template: {_workingTemplatePath}");
        _templateSlides.Clear();

        try
        {
            using (PresentationDocument presentationDoc = PresentationDocument.Open(_workingTemplatePath, false))
            {
                ValidatePresentationDocument(presentationDoc);
                _templateSlides = _templateProcessor.AnalyzeTemplateSlides(presentationDoc.PresentationPart);
            }
        }
        catch (Exception ex)
        {
            Logger.Error($"Error analyzing template: {ex.Message}", ex);
            throw;
        }

        // Create analysis result
        AnalysisResult = CreateAnalysisResult(_templateSlides);
        Logger.Info($"Template analysis complete. Found {AnalysisResult.TotalSlides} slides.");

        return AnalysisResult;
    }

    /// <summary>
    /// Validates that the presentation document is valid
    /// </summary>
    private void ValidatePresentationDocument(PresentationDocument presentationDoc)
    {
        if (presentationDoc.PresentationPart == null)
        {
            throw new InvalidOperationException("Invalid template: Missing presentation part");
        }
    }

    /// <summary>
    /// Creates an analysis result from template slides
    /// </summary>
    private TemplateAnalysisResult CreateAnalysisResult(List<SlideInfo> slides)
    {
        return new TemplateAnalysisResult
        {
            TotalSlides = slides.Count,
            ForeachSourceSlides = slides.Count(s => s.DirectiveType == DirectiveType.Foreach),
            IfSourceSlides = slides.Count(s => s.DirectiveType == DirectiveType.If),
            OriginalSlides = slides.Count(s => !s.HasDirective),
            ImplicitDirectives = slides.Count(s => s.HasImplicitDirective)
        };
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
            Logger.Debug($"Adding variable '{name}' to PowerPoint template");
            Variables[name] = value;

            // Update the effective data source after modifying variables
            UpdateEffectiveDataSource();
        }
        catch (Exception ex)
        {
            Logger.Error($"Failed to add variable '{name}'", ex);
            throw new DocuChefException($"Failed to add variable '{name}': {ex.Message}", ex);
        }
    }

    /// <summary>
    /// Adds variables from a source object
    /// </summary>
    public override void AddVariable(object data)
    {
        base.AddVariable(data);

        // Update the effective data source after adding variables
        UpdateEffectiveDataSource();
    }

    /// <summary>
    /// Sets the hierarchy delimiter used for nested collection paths
    /// </summary>
    /// <param name="delimiter">The delimiter to use (e.g. ">", "::", etc.)</param>
    /// <returns>The current PowerPointRecipe instance for method chaining</returns>
    public PowerPointRecipe SetHierarchyDelimiter(string delimiter)
    {
        ThrowIfDisposed();

        if (string.IsNullOrEmpty(delimiter))
            throw new ArgumentNullException(nameof(delimiter));

        try
        {
            Logger.Debug($"Setting hierarchy delimiter to '{delimiter}'");

            // Update the option
            _options.HierarchyDelimiter = delimiter;

            // Update current options for static access
            PowerPointOptions.SetCurrentOptions(_options);

            // If the template has already been analyzed, re-analyze it with the new delimiter
            if (_templateSlides.Count > 0 && _options.AnalyzeOnInit)
            {
                Logger.Debug("Re-analyzing template with new hierarchy delimiter");
                AnalyzeTemplate();
            }

            return this;
        }
        catch (Exception ex)
        {
            Logger.Error($"Failed to set hierarchy delimiter: {ex.Message}", ex);
            throw new DocuChefException($"Failed to set hierarchy delimiter: {ex.Message}", ex);
        }
    }

    /// <summary>
    /// Updates the effective data source based on current variables
    /// </summary>
    private void UpdateEffectiveDataSource()
    {
        // For PowerPoint, we'll always use a dictionary approach to ensure
        // consistent variable access across different collection types
        var dynamicObj = new System.Dynamic.ExpandoObject() as IDictionary<string, object>;

        // Add all variables to the dynamic object
        foreach (var kvp in Variables)
        {
            dynamicObj[kvp.Key] = kvp.Value;
        }

        _effectiveDataSource = dynamicObj;

        // Re-initialize context processor with updated data source
        _contextProcessor = new ContextProcessor(_effectiveDataSource);

        // Debug log to verify variables
        LogVariables();
    }

    /// <summary>
    /// Logs all variables for debugging
    /// </summary>
    private void LogVariables()
    {
        if (!Logger.IsEnabled || Logger.MinimumLevel > Logger.LogLevel.Debug)
            return;

        Logger.Debug("Current variables:");
        foreach (var kvp in Variables)
        {
            string valueType = kvp.Value?.GetType().Name ?? "null";

            if (kvp.Value is IEnumerable && !(kvp.Value is string))
            {
                int count = 0;
                foreach (var _ in (IEnumerable)kvp.Value) count++;
                Logger.Debug($"  - {kvp.Key}: {valueType} with {count} items");
            }
            else
            {
                Logger.Debug($"  - {kvp.Key}: {valueType}");
            }
        }
    }

    /// <summary>
    /// Creates a presentation generation plan with multi-level nesting support
    /// </summary>
    public PresentationPlan CreatePlan()
    {
        if (_templateSlides.Count == 0)
        {
            AnalyzeTemplate();
        }

        Logger.Info("Creating presentation generation plan with multi-level nesting support...");

        // Add any dynamic global variables to the Variables collection
        foreach (var pair in GlobalVariables)
        {
            try
            {
                Variables[pair.Key] = pair.Value();
            }
            catch (Exception ex)
            {
                Logger.Warning($"Error evaluating dynamic variable '{pair.Key}': {ex.Message}");
                if (_options.ThrowOnMissingVariable)
                    throw;
            }
        }

        // Update data source with latest values
        UpdateEffectiveDataSource();

        // Build the plan using the plan processor
        // The plan processor now respects the original slide order from the template
        _plan = _planProcessor.BuildPlan(_templateSlides, _effectiveDataSource, _contextProcessor);

        // Log slide details
        LogSlideDetails(_plan);

        var summary = _plan.GetSummary();
        Logger.Info($"Created presentation plan: {summary}");
        return _plan;
    }

    /// <summary>
    /// Logs detailed information about each slide in the plan
    /// </summary>
    private void LogSlideDetails(PresentationPlan plan)
    {
        // Log slide details
        Logger.Info("Slide plan details:");
        for (int i = 0; i < plan.IncludedSlides.Count(); i++)
        {
            var slide = plan.IncludedSlides.ElementAt(i);
            string operation = slide.Operation == SlideOperation.Keep ? "Original" : "Clone";
            string context = slide.HasContext ? slide.Context.GetContextDescription() : "No context";

            Logger.Info($"   - Slide {i + 1}: {operation}, {context}");
        }
    }

    /// <summary>
    /// Generates the document from the template
    /// </summary>
    public PowerPointDocument Generate()
    {
        ThrowIfDisposed();

        try
        {
            Logger.Debug("Generating PowerPoint document from template");

            // Ensure we have a plan
            if (_plan == null)
            {
                CreatePlan();
            }

            // Create a temporary file for the output
            string outputPath = Path.Combine(Path.GetTempPath(), $"DocuChef_Output_{Guid.NewGuid().ToString("N")}.pptx");

            // Create a copy of the working template
            File.Copy(_workingTemplatePath, outputPath, true);

            // Open the new presentation for editing with retry logic
            using (PresentationDocument presentationDoc = OpenDocumentWithRetry(outputPath, true))
            {
                // Process the data for PowerPoint 
                var processingData = PrepareDataForProcessing();

                // Execute the presentation generation plan
                _presentationProcessor.ProcessPresentation(presentationDoc, _plan!, processingData);
            }

            Logger.Info("PowerPoint document generated successfully");
            return new PowerPointDocument(outputPath);
        }
        catch (Exception ex)
        {
            Logger.Error("Failed to generate PowerPoint document", ex);
            throw new DocuChefException($"Failed to generate PowerPoint document: {ex.Message}", ex);
        }
    }

    /// <summary>
    /// Opens a PowerPoint document with retry logic for handling access conflicts
    /// </summary>
    private PresentationDocument OpenDocumentWithRetry(string path, bool isEditable, int maxRetries = 3)
    {
        int retryCount = 0;

        while (true)
        {
            try
            {
                return PresentationDocument.Open(path, isEditable);
            }
            catch (IOException ex)
            {
                retryCount++;

                if (retryCount >= maxRetries)
                {
                    Logger.Error($"Failed to open PowerPoint document after {maxRetries} attempts: {ex.Message}");
                    throw;
                }

                // Wait a bit before retrying
                Logger.Warning($"Retry {retryCount}/{maxRetries} - File access conflict, waiting before retry...");
                System.Threading.Thread.Sleep(500 * retryCount); // Incremental backoff
            }
        }
    }

    /// <summary>
    /// Prepares data in the format required by the presentation processor
    /// </summary>
    private object PrepareDataForProcessing()
    {
        // We want to ensure we're passing the variables in the most compatible format
        var data = new Dictionary<string, object>();

        // Copy all variables to the result
        foreach (var kvp in Variables)
        {
            data[kvp.Key] = kvp.Value;
        }

        // Add PPT methods for images and other features
        data["ppt"] = new PPTMethods();

        return data;
    }

    /// <summary>
    /// Updates the template with implicit directives detected during analysis
    /// </summary>
    public void UpdateTemplateWithImplicitDirectives()
    {
        if (_templateSlides.Count == 0)
        {
            AnalyzeTemplate();
        }

        // Find slides with implicit directives
        var slidesWithImplicitDirectives = _templateSlides
            .Where(s => s.HasImplicitDirective && s.Directive != null)
            .ToList();

        if (slidesWithImplicitDirectives.Count == 0)
        {
            Logger.Info("No implicit directives to update in template");
            return;
        }

        Logger.Info($"Updating template with {slidesWithImplicitDirectives.Count} implicit directives");

        try
        {
            using (PresentationDocument presentationDoc = OpenDocumentWithRetry(_workingTemplatePath, true))
            {
                foreach (var slideInfo in slidesWithImplicitDirectives)
                {
                    try
                    {
                        // Get slide part
                        var slidePart = (SlidePart)presentationDoc.PresentationPart.GetPartById(slideInfo.RelationshipId);
                        if (slidePart == null)
                            continue;

                        // Get directive text
                        string directiveText = null;
                        if (slideInfo.Directive is Directives.ForeachDirective foreachDir)
                        {
                            directiveText = $"#foreach: {foreachDir.CollectionName}" +
                                          (foreachDir.MaxItems < int.MaxValue ? $", max: {foreachDir.MaxItems}" : "");
                        }
                        else if (slideInfo.Directive is Directives.IfDirective ifDir)
                        {
                            directiveText = $"#if: {ifDir.Condition}";
                        }

                        // Update slide note with implicit directive
                        if (!string.IsNullOrEmpty(directiveText))
                        {
                            SlideManager.UpdateSlideNoteWithImplicitDirective(slidePart, directiveText);
                            Logger.Info($"Updated slide {slideInfo.SlideId} with implicit directive: {directiveText}");
                        }
                    }
                    catch (Exception ex)
                    {
                        Logger.Warning($"Error updating slide {slideInfo.SlideId} with implicit directive: {ex.Message}");
                    }
                }
            }

            // Re-analyze template to update directive information
            AnalyzeTemplate();
        }
        catch (Exception ex)
        {
            Logger.Error($"Error updating template with implicit directives: {ex.Message}", ex);
            throw;
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
            // Clean up working copy
            CleanupWorkingCopy();

            Logger.Debug("PowerPoint template disposed");
        }

        base.Dispose(disposing);
    }

    /// <summary>
    /// Direct generation to file path without creating PowerPointDocument
    /// </summary>
    public void GenerateToFile(string outputPath)
    {
        ThrowIfDisposed();

        try
        {
            Logger.Debug($"Generating PowerPoint document directly to: {outputPath}");

            // Ensure we have a plan
            if (_plan == null)
            {
                CreatePlan();
            }

            // Create directory if it doesn't exist
            FileExtensions.EnsureDirectoryExists(outputPath);

            // Create a temporary path in case the output file is locked
            string tempOutputPath = Path.Combine(
                Path.GetDirectoryName(outputPath),
                $"Temp_{Path.GetFileName(outputPath)}");

            // Create a copy of the working template
            File.Copy(_workingTemplatePath, tempOutputPath, true);

            // Open the presentation for editing
            using (PresentationDocument presentationDoc = OpenDocumentWithRetry(tempOutputPath, true))
            {
                // Process the data for PowerPoint 
                var processingData = PrepareDataForProcessing();

                // Execute the presentation generation plan
                _presentationProcessor.ProcessPresentation(presentationDoc, _plan, processingData);
            }

            // Move the temporary file to the final location with retry logic
            MoveFileWithRetry(tempOutputPath, outputPath, 3);

            Logger.Info($"PowerPoint document generated successfully to: {outputPath}");
        }
        catch (Exception ex)
        {
            Logger.Error($"Failed to generate PowerPoint document to {outputPath}", ex);
            throw new DocuChefException($"Failed to generate PowerPoint document: {ex.Message}", ex);
        }
    }

    /// <summary>
    /// Moves a file with retry logic for handling file access conflicts
    /// </summary>
    private void MoveFileWithRetry(string sourcePath, string destinationPath, int maxRetries)
    {
        int retryCount = 0;
        bool success = false;

        while (!success && retryCount < maxRetries)
        {
            try
            {
                // Delete destination if it exists, then move the file
                if (File.Exists(destinationPath))
                {
                    File.Delete(destinationPath);
                }

                File.Move(sourcePath, destinationPath);
                success = true;
            }
            catch (IOException ex)
            {
                retryCount++;

                if (retryCount >= maxRetries)
                {
                    Logger.Error($"Failed to move file after {maxRetries} attempts: {ex.Message}");

                    // One last attempt to copy instead of move
                    try
                    {
                        File.Copy(sourcePath, destinationPath, true);
                        File.Delete(sourcePath); // Try to clean up source
                        success = true;
                        Logger.Info("Successfully copied file instead of moving it");
                    }
                    catch (Exception copyEx)
                    {
                        Logger.Error($"Final copy attempt also failed: {copyEx.Message}");
                        throw new DocuChefException($"Could not write to output file: {destinationPath}. Try closing any applications that might be using this file.", ex);
                    }
                }
                else
                {
                    // Wait a bit before retrying
                    Logger.Warning($"Retry {retryCount}/{maxRetries} - File access conflict, waiting before retry...");
                    System.Threading.Thread.Sleep(500 * retryCount); // Incremental backoff
                }
            }
        }

        // Clean up source file if it still exists and we succeeded
        if (success && File.Exists(sourcePath))
        {
            try
            {
                File.Delete(sourcePath);
            }
            catch (Exception ex)
            {
                // Log but don't fail if we can't delete the temp file
                Logger.Warning($"Could not delete temporary file {sourcePath}: {ex.Message}");
            }
        }
    }
}