using DocuChef.Excel;
using DocuChef.PowerPoint;

namespace DocuChef;

/// <summary>
/// Main entry point for DocuChef document generation
/// </summary>
public class Chef : IDisposable
{
    private readonly RecipeOptions _options;
    private readonly Dictionary<string, object> _globalData = new();
    private bool _isDisposed;

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
    }

    /// <summary>
    /// Loads an Excel template from the specified path
    /// </summary>
    public ExcelRecipe LoadExcelTemplate(string templatePath, ExcelOptions options = null)
    {
        ThrowIfDisposed();

        if (string.IsNullOrEmpty(templatePath))
            throw new ArgumentNullException(nameof(templatePath));

        if (!File.Exists(templatePath))
            throw new FileNotFoundException("Template file not found", templatePath);

        try
        {
            var recipe = new ExcelRecipe(templatePath, options ?? _options.Excel);

            // Add global data to the recipe
            foreach (var kvp in _globalData)
            {
                recipe.AddVariable(kvp.Key, kvp.Value);
            }

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
    public ExcelRecipe LoadExcelTemplate(Stream templateStream, ExcelOptions options = null)
    {
        ThrowIfDisposed();

        if (templateStream == null)
            throw new ArgumentNullException(nameof(templateStream));

        try
        {
            var recipe = new ExcelRecipe(templateStream, options ?? _options.Excel);

            // Add global data to the recipe
            foreach (var kvp in _globalData)
            {
                recipe.AddVariable(kvp.Key, kvp.Value);
            }

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
    public PowerPointRecipe LoadPowerPointTemplate(string templatePath, PowerPointOptions options = null)
    {
        ThrowIfDisposed();

        if (string.IsNullOrEmpty(templatePath))
            throw new ArgumentNullException(nameof(templatePath));

        if (!File.Exists(templatePath))
            throw new FileNotFoundException("Template file not found", templatePath);

        try
        {
            var recipe = new PowerPointRecipe(templatePath, options ?? _options.PowerPoint);

            // Add global data to the recipe
            foreach (var kvp in _globalData)
            {
                recipe.AddVariable(kvp.Key, kvp.Value);
            }

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
    public PowerPointRecipe LoadPowerPointTemplate(Stream templateStream, PowerPointOptions options = null)
    {
        ThrowIfDisposed();

        if (templateStream == null)
            throw new ArgumentNullException(nameof(templateStream));

        try
        {
            var recipe = new PowerPointRecipe(templateStream, options ?? _options.PowerPoint);

            // Add global data to the recipe
            foreach (var kvp in _globalData)
            {
                recipe.AddVariable(kvp.Key, kvp.Value);
            }

            return recipe;
        }
        catch (Exception ex) when (!(ex is DocuChefException))
        {
            Logger.Error("Failed to load PowerPoint template from stream", ex);
            throw new DocuChefException($"Failed to load PowerPoint template from stream: {ex.Message}", ex);
        }
    }

    /// <summary>
    /// Adds data to the global registry
    /// </summary>
    public void AddData(object data)
    {
        ThrowIfDisposed();

        if (data == null)
            throw new ArgumentNullException(nameof(data));

        // Add all properties of the object to the global data dictionary
        var properties = data.GetProperties();
        foreach (var kvp in properties)
        {
            _globalData[kvp.Key] = kvp.Value;
        }

        Logger.Debug($"Added {properties.Count} properties to global data from object");
    }

    /// <summary>
    /// Adds named data to the global registry
    /// </summary>
    public void AddData(string key, object data)
    {
        ThrowIfDisposed();

        if (string.IsNullOrEmpty(key))
            throw new ArgumentNullException(nameof(key));

        _globalData[key] = data;
        Logger.Debug($"Added global data with key: {key}");
    }

    /// <summary>
    /// Clears all data from the global registry
    /// </summary>
    public void ClearData()
    {
        ThrowIfDisposed();
        _globalData.Clear();
        Logger.Debug("Cleared global data");
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

    private void ThrowIfDisposed()
    {
        if (_isDisposed)
            throw new ObjectDisposedException(nameof(Chef));
    }
}