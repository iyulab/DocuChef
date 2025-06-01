using DocuChef.Excel;
using DocuChef.Logging;
using DocuChef.Presentation;

namespace DocuChef;

/// <summary>
/// Main entry point for DocuChef document generation
/// </summary>
public class Chef : IDisposable
{
    private readonly RecipeOptions _options;
    private bool _isDisposed;
    private Dictionary<string, object?> _globalData = [];

    /// <summary>
    /// Creates a new DocuChef instance with default options
    /// </summary>
    public Chef() : this(new RecipeOptions())
    {
    }

    /// <summary>
    /// Creates a new DocuChef instance with the specified options
    /// </summary>
    public Chef(RecipeOptions options)
    {
        _options = options ?? new RecipeOptions();

        // Set up logging based on options
        Logger.MinimumLevel = _options.EnableVerboseLogging ?
            Logger.LogLevel.Debug : Logger.LogLevel.Warning;

        Logger.IsEnabled = true;

        // Set up console logging for verbose mode
        if (_options.EnableVerboseLogging)
        {
            Logger.SetLogHandler((message, level) =>
            {
                string prefix = $"[DocuChef:{level}] ";
                Console.WriteLine($"{prefix}{message}");
            });
        }

        Logger.Debug("DocuChef initialized");
    }

    /// <summary>
    /// Loads a template from the specified path
    /// </summary>
    public IRecipe LoadTemplate(string templatePath)
    {
        ThrowIfDisposed();

        if (string.IsNullOrEmpty(templatePath))
            throw new ArgumentNullException(nameof(templatePath));

        if (!File.Exists(templatePath))
            throw new FileNotFoundException("Template file not found", templatePath);

        string extension = Path.GetExtension(templatePath).ToLowerInvariant();

        Logger.Debug($"Loading template from {templatePath} with extension {extension}");

        return extension switch
        {
            ".xlsx" => LoadExcelTemplate(templatePath),
            ".pptx" => LoadPowerPointTemplate(templatePath),
            // Future: .docx support
            _ => throw new DocuChefException($"Unsupported file format: {extension}")
        };
    }    /// <summary>
         /// Loads an Excel template from the specified path
         /// </summary>
    public ExcelRecipe LoadExcelTemplate(string templatePath, ExcelOptions? options = null)
    {
        ThrowIfDisposed();

        if (string.IsNullOrEmpty(templatePath))
            throw new ArgumentNullException(nameof(templatePath));

        if (!File.Exists(templatePath))
            throw new FileNotFoundException("Template file not found", templatePath);

        try
        {
            var recipe = new ExcelRecipe(templatePath, options ?? _options.GetExcelOptions());
            return recipe;
        }
        catch (Exception ex) when (!(ex is DocuChefException))
        {
            Logger.Error($"Failed to load Excel template from {templatePath}", ex);
            throw new DocuChefException($"Failed to load Excel template: {ex.Message}", ex);
        }
    }

    /// <summary>
    /// Loads an Excel template from a stream
    /// </summary>
    public ExcelRecipe LoadExcelTemplate(Stream templateStream, ExcelOptions? options = null)
    {
        ThrowIfDisposed();

        if (templateStream == null)
            throw new ArgumentNullException(nameof(templateStream));

        if (!templateStream.CanRead)
            throw new ArgumentException("Template stream must be readable", nameof(templateStream));

        try
        {
            var recipe = new ExcelRecipe(templateStream, options ?? _options.GetExcelOptions());
            return recipe;
        }
        catch (Exception ex) when (!(ex is DocuChefException))
        {
            Logger.Error("Failed to load Excel template from stream", ex);
            throw new DocuChefException($"Failed to load Excel template from stream: {ex.Message}", ex);
        }
    }

    /// <summary>
    /// Loads a PowerPoint template from the specified path
    /// </summary>
    public PowerPointRecipe LoadPowerPointTemplate(string templatePath, PowerPointOptions? options = null)
    {
        ThrowIfDisposed();

        if (string.IsNullOrEmpty(templatePath))
            throw new ArgumentNullException(nameof(templatePath));

        if (!File.Exists(templatePath))
            throw new FileNotFoundException("Template file not found", templatePath);

        try
        {
            var recipe = new PowerPointRecipe(templatePath, options ?? _options.GetPowerPointOptions());
            return recipe;
        }
        catch (Exception ex) when (!(ex is DocuChefException))
        {
            Logger.Error($"Failed to load PowerPoint template from {templatePath}", ex);
            throw new DocuChefException($"Failed to load PowerPoint template: {ex.Message}", ex);
        }
    }

    /// <summary>
    /// Loads a PowerPoint template from a stream
    /// </summary>
    public PowerPointRecipe LoadPowerPointTemplate(Stream templateStream, PowerPointOptions? options = null)
    {
        ThrowIfDisposed();

        if (templateStream == null)
            throw new ArgumentNullException(nameof(templateStream));

        if (!templateStream.CanRead)
            throw new ArgumentException("Template stream must be readable", nameof(templateStream));

        try
        {
            var recipe = new PowerPointRecipe(templateStream, options ?? _options.GetPowerPointOptions());
            return recipe;
        }
        catch (Exception ex) when (!(ex is DocuChefException))
        {
            Logger.Error("Failed to load PowerPoint template from stream", ex);
            throw new DocuChefException($"Failed to load PowerPoint template from stream: {ex.Message}", ex);
        }
    }

    /// <summary>
    /// Adds a named data item to the global registry
    /// </summary>
    public Chef AddData(string key, object data)
    {
        ThrowIfDisposed();

        if (string.IsNullOrEmpty(key))
            throw new ArgumentNullException(nameof(key));

        _globalData[key] = data;
        Logger.Debug($"Added global data with key: {key}");

        return this; // For method chaining
    }

    /// <summary>
    /// Adds all properties from an object to the global registry
    /// </summary>
    public Chef AddData(object data)
    {
        ThrowIfDisposed();
        ArgumentNullException.ThrowIfNull(data);

        // Add all properties of the object to the global data dictionary
        var properties = data.GetProperties();
        foreach (var kvp in properties)
        {
            _globalData[kvp.Key] = kvp.Value;
        }

        Logger.Debug($"Added {properties.Count} properties to global data from object");

        return this; // For method chaining
    }

    /// <summary>
    /// Clears all data from the global registry
    /// </summary>
    public Chef ClearData()
    {
        ThrowIfDisposed();

        _globalData.Clear();
        Logger.Debug("Cleared global data");

        return this; // For method chaining
    }

    /// <summary>
    /// Throws if the object is disposed
    /// </summary>
    private void ThrowIfDisposed()
    {
        ObjectDisposedException.ThrowIf(_isDisposed, this);
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
        if (_isDisposed) return;

        if (disposing)
        {
            // Dispose any resources here
            _globalData.Clear();
            Logger.Debug("Chef disposed");
        }

        _isDisposed = true;
    }
}