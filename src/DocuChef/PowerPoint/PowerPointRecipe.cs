using System;
using System.Collections.Generic;
using System.IO;
using DocuChef.Extensions;

namespace DocuChef.PowerPoint;

/// <summary>
/// Represents a PowerPoint template for document generation
/// </summary>
public class PowerPointRecipe : RecipeBase
{
    private readonly PowerPointOptions _options;
    private readonly string _templatePath;
    private readonly bool _isTemporaryFile;
    private PowerPointGenerator _generator;

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
        _isTemporaryFile = false;

        InitializeGenerator();

        if (_options.RegisterBuiltInFunctions)
            RegisterBuiltInFunctions();

        if (_options.RegisterGlobalVariables)
            RegisterStandardGlobalVariables();

        Logger.Debug($"PowerPoint recipe initialized from file: {templatePath}");
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
        _isTemporaryFile = true;

        try
        {
            templateStream.CopyToFile(_templatePath);

            InitializeGenerator();

            if (_options.RegisterBuiltInFunctions)
                RegisterBuiltInFunctions();

            if (_options.RegisterGlobalVariables)
                RegisterStandardGlobalVariables();

            Logger.Debug("PowerPoint recipe initialized from stream");
        }
        catch (Exception ex)
        {
            CleanupTemporaryFile();
            Logger.Error("Failed to create PowerPoint recipe from stream", ex);
            throw new DocuChefException($"Failed to create PowerPoint recipe from stream: {ex.Message}", ex);
        }
    }

    /// <summary>
    /// Initializes the generator
    /// </summary>
    private void InitializeGenerator()
    {
        _generator = new PowerPointGenerator(_templatePath, _options);
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

        Variables[$"ppt.{function.Name}"] = function;
    }

    /// <summary>
    /// Registers multiple functions at once
    /// </summary>
    public void RegisterFunctions(IEnumerable<PowerPointFunction> functions)
    {
        ThrowIfDisposed();

        if (functions == null)
            return;

        foreach (var function in functions)
        {
            RegisterFunction(function);
        }
    }

    /// <summary>
    /// Registers built-in functions
    /// </summary>
    private void RegisterBuiltInFunctions()
    {
        try
        {
            // Register PowerPoint specific functions through PowerPointFunctions class
            PowerPointFunctions.RegisterBuiltInFunctions(this);
            Logger.Debug("Registered built-in PowerPoint functions");
        }
        catch (Exception ex)
        {
            Logger.Warning($"Failed to register some built-in functions: {ex.Message}");
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
            // Extract PowerPoint functions from variables
            var powerPointFunctions = new List<PowerPointFunction>();

            foreach (var entry in Variables)
            {
                if (entry.Key.StartsWith("ppt.") && entry.Value is PowerPointFunction function)
                {
                    powerPointFunctions.Add(function);
                }
            }

            // Generate the document
            return _generator.Generate(Variables, GlobalVariables, powerPointFunctions);
        }
        catch (Exception ex)
        {
            Logger.Error("Failed to generate PowerPoint document", ex);
            throw new DocuChefException($"Failed to generate PowerPoint document: {ex.Message}", ex);
        }
    }

    /// <summary>
    /// Cleanup temporary file if needed
    /// </summary>
    private void CleanupTemporaryFile()
    {
        if (_isTemporaryFile &&
            _options.CleanupTemporaryFiles &&
            !string.IsNullOrEmpty(_templatePath) &&
            File.Exists(_templatePath))
        {
            try
            {
                File.Delete(_templatePath);
                Logger.Debug($"Deleted temporary file: {_templatePath}");
            }
            catch (Exception ex)
            {
                Logger.Warning($"Failed to delete temporary file: {ex.Message}");
            }
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
            _generator?.Dispose();
            CleanupTemporaryFile();
            Logger.Debug("PowerPoint recipe disposed");
        }

        base.Dispose(disposing);
    }
}