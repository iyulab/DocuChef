using DocuChef.PowerPoint.Processing;
using DocuChef.PowerPoint.Functions;

namespace DocuChef.PowerPoint;

/// <summary>
/// Main coordinator for PowerPoint document generation using OpenXML
/// </summary>
internal class PowerPointGenerator : IDisposable
{
    private readonly PresentationDocument _document;
    private readonly string _documentPath;
    private readonly PowerPointOptions _options;
    private readonly PowerPointContext _context;
    private bool _isDisposed;

    /// <summary>
    /// Initializes a new PowerPoint generator with template file
    /// </summary>
    public PowerPointGenerator(string templatePath, PowerPointOptions options = null)
    {
        if (string.IsNullOrEmpty(templatePath))
            throw new ArgumentNullException(nameof(templatePath));

        if (!File.Exists(templatePath))
            throw new FileNotFoundException("Template file not found", templatePath);

        _options = options ?? new PowerPointOptions();
        _documentPath = Path.GetTempFileName() + ".pptx";

        // Create a copy of the template
        File.Copy(templatePath, _documentPath, true);

        try
        {
            // Open the document
            _document = PresentationDocument.Open(_documentPath, true);
            _context = CreateInitialContext();

            Logger.Debug($"PowerPoint generator initialized from template: {templatePath}");
        }
        catch (Exception ex)
        {
            CleanupTemporaryFile();
            throw new DocuChefException($"Failed to initialize PowerPoint generator: {ex.Message}", ex);
        }
    }

    /// <summary>
    /// Initializes a new PowerPoint generator with template stream
    /// </summary>
    public PowerPointGenerator(Stream templateStream, PowerPointOptions options = null)
    {
        if (templateStream == null)
            throw new ArgumentNullException(nameof(templateStream));

        _options = options ?? new PowerPointOptions();
        _documentPath = Path.GetTempFileName() + ".pptx";

        try
        {
            // Create temporary file from stream
            using (var fileStream = File.Create(_documentPath))
            {
                templateStream.CopyTo(fileStream);
            }

            // Open the document
            _document = PresentationDocument.Open(_documentPath, true);
            _context = CreateInitialContext();

            Logger.Debug("PowerPoint generator initialized from stream");
        }
        catch (Exception ex)
        {
            CleanupTemporaryFile();
            throw new DocuChefException($"Failed to initialize PowerPoint generator: {ex.Message}", ex);
        }
    }

    /// <summary>
    /// Creates the initial context for the generator
    /// </summary>
    private PowerPointContext CreateInitialContext()
    {
        var context = new PowerPointContext { Options = _options };

        // Set default variables and functions
        context.Variables["_document"] = _document;
        context.Variables["_context"] = context;

        // Initialize navigator
        context.InitializeNavigator();

        return context;
    }

    /// <summary>
    /// Generate PowerPoint document with the provided data
    /// </summary>
    public PowerPointDocument Generate(Dictionary<string, object> variables,
                                       Dictionary<string, Func<object>> globalVariables = null,
                                       IEnumerable<PowerPointFunction> functions = null)
    {
        EnsureNotDisposed();

        try
        {
            // Set up context with variables and functions
            SetupContextData(variables, globalVariables, functions);

            // Process the presentation
            ProcessPresentation();

            // Return the result
            return new PowerPointDocument(_document, _documentPath);
        }
        catch (Exception ex)
        {
            // Cleanup and throw
            _document?.Dispose();
            CleanupTemporaryFile();

            Logger.Error("Failed to generate PowerPoint document", ex);
            throw new DocuChefException($"Failed to generate PowerPoint document: {ex.Message}", ex);
        }
    }

    /// <summary>
    /// Set up context with the provided data
    /// </summary>
    private void SetupContextData(Dictionary<string, object> variables,
                                  Dictionary<string, Func<object>> globalVariables,
                                  IEnumerable<PowerPointFunction> functions)
    {
        // Variables
        if (variables != null)
        {
            foreach (var kvp in variables)
            {
                _context.Variables[kvp.Key] = kvp.Value;
            }
        }

        // Global variables
        if (globalVariables != null)
        {
            foreach (var kvp in globalVariables)
            {
                _context.GlobalVariables[kvp.Key] = kvp.Value;
            }
        }

        // PowerPoint functions
        if (functions != null)
        {
            foreach (var function in functions)
            {
                _context.Functions[function.Name] = function;
            }
        }

        // Register built-in functions if enabled
        if (_options.RegisterBuiltInFunctions)
        {
            RegisterBuiltInFunctions();
        }
    }

    /// <summary>
    /// Register built-in PowerPoint functions
    /// </summary>
    private void RegisterBuiltInFunctions()
    {
        // Image function
        _context.Functions["Image"] = ImageFunction.Create();

        // Chart function
        _context.Functions["Chart"] = ChartFunction.Create();

        // Table function
        _context.Functions["Table"] = TableFunction.Create();

        Logger.Debug("Registered built-in PowerPoint functions");
    }

    /// <summary>
    /// Process the presentation by analyzing and preparing slides
    /// </summary>
    private void ProcessPresentation()
    {
        var presentationPart = _document.PresentationPart;
        if (presentationPart?.Presentation?.SlideIdList == null)
        {
            throw new DocuChefException("Invalid PowerPoint document structure");
        }

        // Prepare variables dictionary
        var variables = PrepareVariablesDictionary();

        // Create processors
        var expressionEvaluator = ProcessorFactory.CreateExpressionEvaluator();
        var slideProcessor = new SlideProcessor(expressionEvaluator, _context);

        // Process in two phases

        // Phase 1: Analyze and prepare slides (may create new slides)
        Logger.Info("Phase 1: Analyzing and preparing slides...");
        var slideIds = GetSlideIds(presentationPart).ToList();

        foreach (var slideId in slideIds)
        {
            var slidePart = (SlidePart)presentationPart.GetPartById(slideId.RelationshipId);
            slideProcessor.AnalyzeAndPrepareSlide(presentationPart, slidePart, variables);
        }

        // Phase 2: Apply bindings to all slides (original and newly created)
        Logger.Info("Phase 2: Applying bindings to all slides...");
        var allSlideIds = GetSlideIds(presentationPart);

        foreach (var slideId in allSlideIds)
        {
            var slidePart = (SlidePart)presentationPart.GetPartById(slideId.RelationshipId);
            slideProcessor.ApplyBindings(slidePart, variables);
        }

        // Save changes
        presentationPart.Presentation.Save();
        Logger.Info("PowerPoint document processing completed successfully");
    }

    /// <summary>
    /// Get all slide IDs from presentation
    /// </summary>
    private IEnumerable<SlideId> GetSlideIds(PresentationPart presentationPart)
    {
        return presentationPart.Presentation.SlideIdList
            .ChildElements.OfType<SlideId>();
    }

    /// <summary>
    /// Prepare variables dictionary with all necessary data
    /// </summary>
    private Dictionary<string, object> PrepareVariablesDictionary()
    {
        var variables = new Dictionary<string, object>(_context.Variables);

        // Add global variables
        foreach (var globalVar in _context.GlobalVariables)
        {
            variables[globalVar.Key] = globalVar.Value();
        }

        // Add PowerPoint functions
        foreach (var function in _context.Functions)
        {
            variables[$"ppt.{function.Key}"] = function.Value;
        }

        return variables;
    }

    /// <summary>
    /// Cleanup temporary file if exists
    /// </summary>
    private void CleanupTemporaryFile()
    {
        if (_options.CleanupTemporaryFiles &&
            !string.IsNullOrEmpty(_documentPath) &&
            File.Exists(_documentPath))
        {
            try
            {
                File.Delete(_documentPath);
                Logger.Debug($"Deleted temporary file: {_documentPath}");
            }
            catch (Exception ex)
            {
                Logger.Warning($"Failed to delete temporary file: {ex.Message}");
            }
        }
    }

    /// <summary>
    /// Ensure the generator is not disposed
    /// </summary>
    private void EnsureNotDisposed()
    {
        if (_isDisposed)
            throw new ObjectDisposedException(nameof(PowerPointGenerator));
    }

    /// <summary>
    /// Dispose resources
    /// </summary>
    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    /// <summary>
    /// Dispose resources
    /// </summary>
    protected virtual void Dispose(bool disposing)
    {
        if (_isDisposed) return;

        if (disposing)
        {
            _document?.Dispose();
            CleanupTemporaryFile();
            Logger.Debug("PowerPoint generator disposed");
        }

        _isDisposed = true;
    }
}