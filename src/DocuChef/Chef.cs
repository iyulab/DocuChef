using DocuChef.Excel;
using DocuChef.Presentation;

namespace DocuChef;

/// <summary>
/// Main entry point for DocuChef document generation
/// </summary>
public class Chef : IDisposable
{
    private readonly RecipeOptions _options;
    private readonly Dictionary<string, object> _globalData = new Dictionary<string, object>();
    private readonly Dictionary<string, Func<object>> _dynamicData = new Dictionary<string, Func<object>>();
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

        Logger.IsEnabled = true;

        // Register default dynamic data providers
        RegisterDynamicData("Today", () => DateTime.Today);
        RegisterDynamicData("Now", () => DateTime.Now);
        RegisterDynamicData("UtcNow", () => DateTime.UtcNow);
        RegisterDynamicData("Random", () => new Random());

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
            var recipe = new ExcelRecipe(templatePath, options ?? _options.GetExcelOptions());

            // Add global data to the recipe
            ApplyGlobalData(recipe);

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

        if (!templateStream.CanRead)
            throw new ArgumentException("Template stream must be readable", nameof(templateStream));

        try
        {
            var recipe = new ExcelRecipe(templateStream, options ?? _options.GetExcelOptions());

            // Add global data to the recipe
            ApplyGlobalData(recipe);

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
            var recipe = new PowerPointRecipe(templatePath, options ?? _options.GetPowerPointOptions());

            // Add global data to the recipe
            ApplyGlobalData(recipe);

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

        if (!templateStream.CanRead)
            throw new ArgumentException("Template stream must be readable", nameof(templateStream));

        try
        {
            var recipe = new PowerPointRecipe(templateStream, options ?? _options.GetPowerPointOptions());

            // Add global data to the recipe
            ApplyGlobalData(recipe);

            return recipe;
        }
        catch (Exception ex) when (!(ex is DocuChefException))
        {
            Logger.Error("Failed to load PowerPoint template from stream", ex);
            throw new DocuChefException($"Failed to load PowerPoint template from stream: {ex.Message}", ex);
        }
    }

    /// <summary>
    /// Prepare a dish by loading a template and cooking it with provided data
    /// </summary>
    public IDish PrepareDish(string templatePath, object data)
    {
        ThrowIfDisposed();

        if (string.IsNullOrEmpty(templatePath))
            throw new ArgumentNullException(nameof(templatePath));

        if (data == null)
            throw new ArgumentNullException(nameof(data));

        // Load the appropriate recipe based on file extension
        var recipe = LoadTemplate(templatePath);

        // Add data to the recipe
        recipe.AddVariable(data);

        // Generate the document
        return recipe.CookDish();
    }

    /// <summary>
    /// Prepare a dish and save it to the specified output path
    /// </summary>
    public void PrepareDish(string templatePath, object data, string outputPath)
    {
        ThrowIfDisposed();

        if (string.IsNullOrEmpty(outputPath))
            throw new ArgumentNullException(nameof(outputPath));

        // Prepare the dish
        var dish = PrepareDish(templatePath, data);

        // Save it to the specified path
        dish.SaveAs(outputPath);
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

        if (data == null)
            throw new ArgumentNullException(nameof(data));

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
    /// Registers a dynamic data provider that will be evaluated at document generation time
    /// </summary>
    public Chef RegisterDynamicData(string key, Func<object> dataProvider)
    {
        ThrowIfDisposed();

        if (string.IsNullOrEmpty(key))
            throw new ArgumentNullException(nameof(key));

        if (dataProvider == null)
            throw new ArgumentNullException(nameof(dataProvider));

        _dynamicData[key] = dataProvider;
        Logger.Debug($"Registered dynamic data provider with key: {key}");

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
    /// Clears all dynamic data providers
    /// </summary>
    public Chef ClearDynamicData()
    {
        ThrowIfDisposed();

        _dynamicData.Clear();
        Logger.Debug("Cleared dynamic data providers");

        return this; // For method chaining
    }

    /// <summary>
    /// Applies global and dynamic data to a recipe
    /// </summary>
    private void ApplyGlobalData(IRecipe recipe)
    {
        // Apply static global data
        foreach (var kvp in _globalData)
        {
            recipe.AddVariable(kvp.Key, kvp.Value);
        }

        // Apply dynamic data providers
        foreach (var kvp in _dynamicData)
        {
            recipe.RegisterGlobalVariable(kvp.Key, kvp.Value);
        }
    }

    /// <summary>
    /// Gets all registered data keys
    /// </summary>
    public IEnumerable<string> GetDataKeys()
    {
        ThrowIfDisposed();

        var keys = new HashSet<string>(_globalData.Keys);
        foreach (var dynamicKey in _dynamicData.Keys)
        {
            keys.Add(dynamicKey);
        }

        return keys;
    }

    /// <summary>
    /// Throws if the object is disposed
    /// </summary>
    private void ThrowIfDisposed()
    {
        if (_isDisposed)
            throw new ObjectDisposedException(nameof(Chef));
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
            _dynamicData.Clear();
            Logger.Debug("Chef disposed");
        }

        _isDisposed = true;
    }
}