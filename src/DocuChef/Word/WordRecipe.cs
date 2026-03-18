namespace DocuChef.Word;

/// <summary>
/// Processes Word templates and generates output documents
/// </summary>
public class WordRecipe : RecipeBase
{
    private readonly WordOptions _options;
    private readonly string? _templatePath;
    private readonly MemoryStream? _templateMemoryStream;

    /// <summary>
    /// Creates a new Word recipe using a template file
    /// </summary>
    /// <param name="templatePath">Path to the template file</param>
    /// <param name="options">Options for processing</param>
    public WordRecipe(string templatePath, WordOptions? options = null)
    {
        if (string.IsNullOrEmpty(templatePath))
            throw new ArgumentNullException(nameof(templatePath));

        if (!File.Exists(templatePath))
            throw new FileNotFoundException("Template file not found", templatePath);

        _options = options ?? new WordOptions();
        _templatePath = templatePath;

        RegisterStandardGlobalVariables();
        Logger.Debug($"Word recipe initialized from {templatePath}");
    }

    /// <summary>
    /// Creates a new Word recipe using a template stream
    /// </summary>
    /// <param name="templateStream">Stream containing the template</param>
    /// <param name="options">Options for processing</param>
    public WordRecipe(Stream templateStream, WordOptions? options = null)
    {
        if (templateStream == null)
            throw new ArgumentNullException(nameof(templateStream));

        _options = options ?? new WordOptions();

        // Copy the stream to memory for reuse
        _templateMemoryStream = new MemoryStream();
        templateStream.CopyTo(_templateMemoryStream);
        _templateMemoryStream.Position = 0;

        RegisterStandardGlobalVariables();
        Logger.Debug("Word recipe initialized from stream");
    }

    /// <summary>
    /// Adds a named variable to the recipe
    /// </summary>
    /// <param name="name">Variable name</param>
    /// <param name="value">Variable value</param>
    public override void AddVariable(string name, object value)
    {
        ThrowIfDisposed();

        if (string.IsNullOrEmpty(name))
            throw new ArgumentNullException(nameof(name));

        Variables[name] = value;
    }

    /// <summary>
    /// Generates the Word document from the template
    /// </summary>
    /// <returns>Generated Word document</returns>
    public override IDish Generate()
    {
        ThrowIfDisposed();

        // Will be wired in Task 13 (WordTemplateProcessor)
        throw new NotImplementedException(
            "Word document generation is not yet implemented. " +
            "This will be completed when the WordTemplateProcessor pipeline is wired.");
    }

    /// <summary>
    /// Disposes resources
    /// </summary>
    protected override void Dispose(bool disposing)
    {
        if (IsDisposed) return;

        if (disposing)
        {
            _templateMemoryStream?.Dispose();
            Logger.Debug("Word recipe disposed");
        }

        base.Dispose(disposing);
    }
}
