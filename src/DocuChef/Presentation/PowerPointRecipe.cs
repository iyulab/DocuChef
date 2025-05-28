using System.Collections.Concurrent;
using System.Text.RegularExpressions;
using DocuChef.Presentation.Models;
using DocuChef.Presentation.Processors;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Drawing;
using DocuChef.Logging;

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
    /// Registers a global variable accessible across all templates
    /// </summary>
    /// <param name="name">Global variable name</param>
    /// <param name="value">Global variable value</param>
    public void RegisterGlobalVariable(string name, object value)
    {
        if (string.IsNullOrEmpty(name))
            throw new ArgumentException("Global variable name cannot be null or empty", nameof(name));
        
        globalVariables[name] = value;
    }

    /// <summary>
    /// Registers a custom function that can be used in template expressions
    /// </summary>
    /// <param name="name">Function name</param>
    /// <param name="function">Function implementation</param>
    public void RegisterFunction(string name, Func<object, object> function)
    {
        if (string.IsNullOrEmpty(name))
            throw new ArgumentException("Function name cannot be null or empty", nameof(name));
        
        if (function == null)
            throw new ArgumentNullException(nameof(function));
        
        customFunctions[name] = function;
    }

    /// <summary>
    /// Generates the output document
    /// </summary>
    /// <returns>The generated PowerPoint document</returns>
    internal IDish Generate()
    {
        try
        {
            if (options.EnableVerboseLogging)
            {
                Logger.Debug("Starting PowerPoint document generation");
            }            // Create a copy of the template for processing using a temporary file approach
            var tempFilePath = System.IO.Path.GetTempFileName() + ".pptx";
            
            if (templateMemoryStream != null)
            {
                templateMemoryStream.Position = 0;
                using (var fileStream = File.Create(tempFilePath))
                {
                    templateMemoryStream.CopyTo(fileStream);
                }
            }
            else
            {
                File.Copy(templatePath!, tempFilePath, true);
            }

            if (options.EnableVerboseLogging)
            {
                Logger.Debug($"Created temporary template file: {tempFilePath}");
            }

            // Process the document using the temporary file
            try
            {
                using (var presentationDocument = PresentationDocument.Open(tempFilePath, true))
                {                
                    // Analyze template structure
                    var slideInfos = AnalyzeTemplateSlides(presentationDocument);

                    if (options.EnableVerboseLogging)
                    {
                        Logger.Debug($"Template analysis complete. Found {slideInfos.Count} slides");
                    }                    // Check if template has array binding elements
                    bool hasArrayElements = slideInfos.Any(si => 
                        si.Directives.Any(d => d.Type == DirectiveType.Foreach || d.Type == DirectiveType.Range) ||
                        !string.IsNullOrEmpty(si.CollectionName));

                    if (options.EnableVerboseLogging)
                    {
                        Logger.Debug($"hasArrayElements check: {hasArrayElements}");
                        Logger.Debug($"SlideInfos count: {slideInfos.Count}");
                        for (int i = 0; i < slideInfos.Count; i++)
                        {
                            var si = slideInfos[i];
                            Logger.Debug($"  Slide {i}: Directives={si.Directives.Count}, CollectionName='{si.CollectionName}'");
                        }
                    }

                    if (!hasArrayElements)
                    {
                        // Design-centered approach: prioritize template design over data structure
                        if (options.EnableVerboseLogging)
                        {
                            Logger.Debug("No array elements in template - performing direct data binding to original slides");
                        }

                        // Bind data directly to the original template slides
                        var dataBinder = new DataBinder();
                        var combinedData = CombineData();

                        var slideList = presentationDocument.PresentationPart?.Presentation?.SlideIdList?.Elements<SlideId>()?.ToList();
                        if (slideList != null)
                        {
                            for (int i = 0; i < slideList.Count; i++)
                            {
                                var slideId = slideList[i];
                                var slidePart = (SlidePart?)presentationDocument.PresentationPart?.GetPartById(slideId.RelationshipId!);
                                
                                if (slidePart?.Slide != null)
                                {
                                    if (options.EnableVerboseLogging)
                                    {
                                        Logger.Debug($"Processing original slide {i + 1} (SlideId: {slideId.Id})");
                                    }
                                    ScanAndReplaceExpressionsInSlide(slidePart, dataBinder, combinedData, i);
                                }
                            }
                        }
                    }                    else
                    {
                        // Generate slide plan
                        var slidePlanGenerator = new SlidePlanGenerator();
                        var slidePlan = slidePlanGenerator.GeneratePlan(slideInfos, CombineData());

                        if (options.EnableVerboseLogging)
                        {
                            Logger.Debug($"Slide plan generated with {slidePlan.SlideInstances.Count} slide instances");
                        }

                        // Generate slides based on the plan
                        var slideGenerator = new SlideGenerator();
                        slideGenerator.GenerateSlides(presentationDocument, slidePlan);

                        if (options.EnableVerboseLogging)
                        {
                            Logger.Debug("Slide generation complete");
                        }                        // Bind data to generated slides  
                        var dataBinder = new DataBinder();
                        var combinedData = CombineData();

                        if (options.EnableVerboseLogging)
                        {
                            Logger.Debug("Starting data binding to generated slides");
                        }

                        // Simple slide processing - process all slides in order
                        var slideList = presentationDocument.PresentationPart?.Presentation?.SlideIdList?.Elements<SlideId>()?.ToList();
                        if (slideList != null)
                        {
                            for (int i = 0; i < slideList.Count; i++)
                            {
                                var slideId = slideList[i];
                                var slidePart = (SlidePart?)presentationDocument.PresentationPart?.GetPartById(slideId.RelationshipId!);
                                
                                if (slidePart?.Slide != null)
                                {
                                    ScanAndReplaceExpressionsInSlide(slidePart, dataBinder, combinedData, i);
                                }
                            }
                        }                        if (options.EnableVerboseLogging)
                        {
                            Logger.Debug("Data binding complete");
                        }
                    }                    // Save changes to the document
                    presentationDocument.Save();
                    
                    if (options.EnableVerboseLogging)
                    {
                        Logger.Debug("Document changes saved to temporary file");
                        
                        // Debug: Verify slide count after save
                        var finalSlideList = presentationDocument.PresentationPart?.Presentation?.SlideIdList?.Elements<SlideId>()?.ToList();
                        Logger.Debug($"Final document has {finalSlideList?.Count ?? 0} slides after save");
                        
                        // Debug: Check text content in final slides
                        if (finalSlideList != null)
                        {
                            for (int i = 0; i < finalSlideList.Count; i++)
                            {
                                var slideId = finalSlideList[i];
                                var slidePart = (SlidePart?)presentationDocument.PresentationPart?.GetPartById(slideId.RelationshipId!);
                                if (slidePart?.Slide != null)
                                {
                                    var textElements = slidePart.Slide.Descendants<DocumentFormat.OpenXml.Drawing.Text>().ToList();
                                    Logger.Debug($"Final slide {i}: {textElements.Count} text elements");
                                    foreach (var textElement in textElements.Take(3)) // Log first 3 for brevity
                                    {
                                        Logger.Debug($"  Text: '{textElement.Text}'");
                                    }
                                }
                            }
                        }
                    }                }
                
                // At this point, the temporary file contains the processed document
                var powerPointDocument = new PowerPointDocument(tempFilePath);
                
                if (options.EnableVerboseLogging)
                {
                    Logger.Debug("PowerPoint document generation completed successfully using temporary file");
                }
                
                return powerPointDocument;
            }
            catch (Exception ex)
            {
                // Clean up temporary file on error
                try { File.Delete(tempFilePath); } catch { }
                Logger.Error($"Error generating PowerPoint document: {ex.Message}", ex);
                throw;
            }
        }
        catch (Exception ex)
        {
            Logger.Error($"Error in PowerPoint document generation: {ex.Message}", ex);
            throw;
        }
    }

    /// <summary>
    /// Combines variables and data objects into a single dictionary
    /// </summary>
    /// <returns>Combined data dictionary</returns>
    private Dictionary<string, object> CombineData()
    {
        var combinedData = new Dictionary<string, object>(globalVariables);
        
        // Add explicit variables
        foreach (var kvp in variables)
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
        
        return combinedData;
    }

    /// <summary>
    /// Binds data to all slides in the presentation
    /// </summary>
    /// <param name="presentationDocument">The presentation document</param>
    /// <param name="slideInfos">Information about slides</param>
    private void BindData(PresentationDocument presentationDocument, List<SlideInfo> slideInfos)
    {
        var data = CombineData();
        var dataBinder = new DataBinder();
        var presentationPart = presentationDocument.PresentationPart;
        
        if (presentationPart == null || presentationPart.Presentation?.SlideIdList == null)
            return;
        
        if (options.EnableVerboseLogging)
        {
            Logger.Debug($"Starting data binding with {data.Count} variables");
            foreach (var kvp in data)
            {
                Logger.Debug($"  Variable: {kvp.Key} = {kvp.Value}");
            }
        }
          // Get all slides
        var slideIds = presentationPart.Presentation.SlideIdList.ChildElements
            .OfType<DocumentFormat.OpenXml.Presentation.SlideId>();
            
        if (options.EnableVerboseLogging)
        {
            Logger.Debug($"Found {slideIds.Count()} actual slides in presentation document");
        }
        
          int slideIndex = 0;
        foreach (var slideId in slideIds)
        {
            if (options.EnableVerboseLogging)
            {
                Logger.Debug($"Processing slide {slideIndex} (RelId: {slideId.RelationshipId?.Value}, Id: {slideId.Id})");
            }
            
            string? relationshipId = slideId.RelationshipId?.Value;
            if (string.IsNullOrEmpty(relationshipId))
            {
                if (options.EnableVerboseLogging)
                {
                    Logger.Warning($"Slide {slideIndex} has no relationship ID, skipping");
                }
                slideIndex++;
                continue;
            }
              SlidePart? slidePart = null;
            try
            {
                slidePart = (SlidePart)presentationPart.GetPartById(relationshipId);
                
                if (options.EnableVerboseLogging)
                {
                    Logger.Debug($"Successfully got slide part for slide {slideIndex}: {slidePart != null}");
                    if (slidePart != null)
                    {
                        Logger.Debug($"Slide {slideIndex} has slide object: {slidePart.Slide != null}");
                    }
                }
            }
            catch (Exception ex)
            {
                if (options.EnableVerboseLogging)
                {
                    Logger.Warning($"Error getting slide part for slide {slideIndex}: {ex.Message}");
                }
                slideIndex++;
                continue;
            }
              if (options.EnableVerboseLogging)
            {
                Logger.Debug($"Processing slide {slideIndex}");
                if (slideInfos.Any())
                {
                    Logger.Debug($"Available slideInfos: {string.Join(", ", slideInfos.Select(s => $"SlideId={s.SlideId}"))}");
                }
                else
                {
                    Logger.Debug("No slideInfos available");
                }
            }
            
            // Get slide info if available
            var slideInfo = slideInfos.FirstOrDefault(s => s.SlideId == slideIndex);
            
            if (options.EnableVerboseLogging)
            {
                Logger.Debug($"SlideInfo found for slide {slideIndex}: {slideInfo != null}");
            }
            
            // Process expressions from slideInfo if available
            if (slideInfo?.BindingExpressions != null)
            {
                if (options.EnableVerboseLogging)
                {
                    Logger.Debug($"Found {slideInfo.BindingExpressions.Count} binding expressions in slide {slideIndex}");
                }
                
                foreach (var expression in slideInfo.BindingExpressions)
                {
                    if (expression?.OriginalExpression == null)
                        continue;
                        
                    try
                    {
                        var value = dataBinder.ResolveExpression(expression.OriginalExpression, data);
                          // Replace in slide text
                        if (slidePart != null)
                        {
                            ReplaceTextInSlide(slidePart, expression.OriginalExpression, value?.ToString() ?? "");
                        }
                        
                        if (options.EnableVerboseLogging)
                        {
                            Logger.Debug($"Replaced '{expression.OriginalExpression}' with '{value?.ToString() ?? ""}' in slide {slideIndex}");
                        }
                    }
                    catch (Exception ex)
                    {
                        if (options.EnableVerboseLogging)
                        {
                            Logger.Warning($"Error processing expression '{expression.OriginalExpression}': {ex.Message}");
                        }
                        // Continue processing other expressions
                    }
                }
            }
            else
            {
                // Fallback: scan slide text directly for expressions and replace them                if (options.EnableVerboseLogging)
                {
                    Logger.Debug($"No binding expressions found for slide {slideIndex}, scanning for expressions directly");
                }
                  try 
                {
                    if (slidePart != null)
                    {
                        ScanAndReplaceExpressionsInSlide(slidePart, dataBinder, data, slideIndex);
                    }
                }
                catch (Exception ex)
                {
                    if (options.EnableVerboseLogging)
                    {
                        Logger.Error($"Error in ScanAndReplaceExpressionsInSlide for slide {slideIndex}: {ex.Message}", ex);
                    }
                }
            }
            
            // Process image placeholders created by PPTFunctions
            try
            {
                if (slidePart != null)
                {
                    ProcessImagePlaceholders(slidePart, dataBinder);
                }
            }
            catch (Exception ex)
            {
                if (options.EnableVerboseLogging)
                {
                    Logger.Warning($"Error processing functions in slide {slideIndex}: {ex.Message}");
                }
            }
            
            slideIndex++;
        }
          if (options.EnableVerboseLogging)
        {
            Logger.Debug("Data binding completed for all slides");
        }
    }    /// <summary>
    /// Scans slide text directly for expressions and replaces them when slideInfo.BindingExpressions is not available
    /// </summary>
    private void ScanAndReplaceExpressionsInSlide(SlidePart slidePart, DataBinder dataBinder, Dictionary<string, object> data, int slideIndex)
    {
        if (options.EnableVerboseLogging)
        {
            Logger.Debug($"ScanAndReplaceExpressionsInSlide called for slide {slideIndex}");
        }
          if (slidePart?.Slide == null)
        {
            if (options.EnableVerboseLogging)
            {
                Logger.Debug($"Slide {slideIndex} is null or has no Slide object, skipping");
            }
            return;
        }
            
        try
        {
            // OPTIMIZATION: First extract all expressions from the slide
            var usedExpressions = ExtractExpressionsFromSlide(slidePart);
            
            if (options.EnableVerboseLogging)
            {
                Logger.Debug($"Slide {slideIndex} contains {usedExpressions.Count} unique expressions: {string.Join(", ", usedExpressions)}");
            }

            // Get all text elements in the slide
            var textElements = slidePart.Slide.Descendants<DocumentFormat.OpenXml.Drawing.Text>();
            
            if (options.EnableVerboseLogging)
            {
                Logger.Debug($"Found {textElements.Count()} text elements in slide {slideIndex}");
            }
            
            foreach (var textElement in textElements)
            {
                if (textElement == null)
                    continue;
                    
                var originalText = textElement.Text ?? "";
                
                if (options.EnableVerboseLogging && !string.IsNullOrWhiteSpace(originalText))
                {
                    Logger.Debug($"Processing text element in slide {slideIndex}: '{originalText}'");
                }
                
                if (string.IsNullOrEmpty(originalText))
                    continue;
                    
                var modifiedText = originalText;
                
                // Look for expressions in the format ${variableName} or ${object.property}
                var expressionRegex = new System.Text.RegularExpressions.Regex(@"\$\{([^}]+)\}", System.Text.RegularExpressions.RegexOptions.Compiled);
                var matches = expressionRegex.Matches(originalText);
                
                if (options.EnableVerboseLogging && matches.Count > 0)
                {
                    Logger.Debug($"Found {matches.Count} expressions in text: '{originalText}'");
                }
                
                foreach (System.Text.RegularExpressions.Match match in matches)
                {
                    var fullExpression = match.Value; // e.g., "${Title}"
                    var expressionContent = match.Groups[1].Value; // e.g., "Title"
                    
                    if (options.EnableVerboseLogging)
                    {
                        Logger.Debug($"Processing expression: '{fullExpression}' (content: '{expressionContent}')");
                    }
                      try
                    {
                        // Use OPTIMIZED DataBinder to resolve the expression with filtered variables and custom functions
                        var value = dataBinder.ResolveExpressionWithFilteredVariables(fullExpression, data, usedExpressions, customFunctions);
                        
                        // Replace the expression in the text
                        modifiedText = modifiedText.Replace(fullExpression, value?.ToString() ?? "");
                        
                        if (options.EnableVerboseLogging)
                        {
                            Logger.Debug($"Replaced '{fullExpression}' with '{value?.ToString() ?? ""}' in slide {slideIndex} (optimized)");
                        }
                    }
                    catch (Exception ex)
                    {
                        if (options.EnableVerboseLogging)
                        {
                            Logger.Warning($"Error resolving expression '{fullExpression}' in slide {slideIndex}: {ex.Message}");
                        }
                        // Continue with other expressions
                    }
                }
                  // Update the text element if any replacements were made
                if (modifiedText != originalText)
                {
                    if (options.EnableVerboseLogging)
                    {
                        Logger.Debug($"Updating text element from '{originalText}' to '{modifiedText}' in slide {slideIndex}");
                        Logger.Debug($"Text element parent: {textElement.Parent?.GetType().Name}");
                        Logger.Debug($"Text element XML before: {textElement.OuterXml}");
                    }
                    
                    textElement.Text = modifiedText;
                    
                    if (options.EnableVerboseLogging)
                    {
                        Logger.Debug($"Text element XML after: {textElement.OuterXml}");
                        Logger.Debug($"Text element text after update: '{textElement.Text}'");
                    }
                }
            }
        }
        catch (Exception ex)
        {
            if (options.EnableVerboseLogging)
            {
                Logger.Warning($"Error scanning expressions in slide {slideIndex}: {ex.Message}");
            }
        }
    }

    /// <summary>
    /// Replaces text in a slide
    /// </summary>
    private void ReplaceTextInSlide(SlidePart slidePart, string expression, string value)
    {
        if (slidePart?.Slide == null || string.IsNullOrEmpty(expression))
            return;
            
        try
        {
            // Using OpenXML API to replace text
            var textElements = slidePart.Slide.Descendants<DocumentFormat.OpenXml.Drawing.Text>();
            foreach (var textElement in textElements)
            {
                if (textElement != null && !string.IsNullOrEmpty(textElement.Text) && textElement.Text.Contains(expression))
                {
                    textElement.Text = textElement.Text.Replace(expression, value ?? "");
                }
            }
        }
        catch (Exception ex)
        {
            // Log the error but don't rethrow to avoid breaking the entire process
            if (options.EnableVerboseLogging)
            {
                Logger.Warning($"Error replacing text in slide: {ex.Message}");
            }
        }
    }    /// <summary>
    /// Generates the output document
    /// </summary>
    /// <param name="outputPath">The output file path</param>
    /// <returns>The generated PowerPoint document</returns>
    public IDish Cook(string outputPath)
    {
        options.OutputPath = outputPath;
        var document = Generate();
        
        // Save the generated document to the specified output path
        document.SaveAs(outputPath);
        
        return document;
    }

    /// <summary>
    /// Generates the output document to a stream
    /// </summary>
    /// <param name="outputStream">The output stream</param>
    /// <returns>The generated PowerPoint document</returns>
    public IDish Cook(Stream outputStream)
    {
        var document = Generate();
        document.SaveAs(outputStream);
        return document;
    }

    /// <summary>
    /// Processes image placeholders created by PPTFunctions
    /// </summary>
    /// <param name="slidePart">The slide part to process</param>
    /// <param name="dataBinder">The data binder containing PPT functions</param>
    private void ProcessImagePlaceholders(SlidePart slidePart, DataBinder dataBinder)
    {
        if (slidePart?.Slide == null)
            return;

        var slide = slidePart.Slide;
        var textElements = slide.Descendants<A.Text>();

        foreach (var textElement in textElements.ToList())
        {
            if (string.IsNullOrEmpty(textElement.Text))
                continue;

            // Look for image placeholders created by PPTFunctions
            var placeholderRegex = new Regex(@"__PPT_IMAGE_([0-9a-fA-F]+)__([^_]+)(?:__(\d+)__(\d+)__(true|false))?__", RegexOptions.Compiled);
            var matches = placeholderRegex.Matches(textElement.Text);

            foreach (Match match in matches)
            {
                try
                {
                    var guid = match.Groups[1].Value;
                    var propertyPath = match.Groups[2].Value;
                    var widthStr = match.Groups[3].Success ? match.Groups[3].Value : null;
                    var heightStr = match.Groups[4].Success ? match.Groups[4].Value : null;
                    var preserveAspectRatioStr = match.Groups[5].Success ? match.Groups[5].Value : "true";

                    // Parse dimensions
                    int? width = null, height = null;
                    if (int.TryParse(widthStr, out var w)) width = w;
                    if (int.TryParse(heightStr, out var h)) height = h;
                    var preserveAspectRatio = bool.TryParse(preserveAspectRatioStr, out var preserve) ? preserve : true;

                    // Get image data from PPTFunctions cache
                    // Note: We would need to modify DataBinder to expose the PPTFunctions instance
                    // For now, we'll log that we found the placeholder
                    if (options.EnableVerboseLogging)
                    {
                        Logger.Debug($"Found PPT image placeholder: {propertyPath} ({width}x{height})");
                    }                    // Remove the placeholder text for now
                    textElement.Text = textElement.Text.Replace(match.Value, "");
                }
                catch (Exception ex)
                {
                    if (options.EnableVerboseLogging)
                    {
                        Logger.Warning($"Error processing image placeholder: {ex.Message}");
                    }
                }
            }
        }
    }

    /// <summary>
    /// Analyzes all slides in the presentation document
    /// </summary>
    /// <param name="presentationDocument">The presentation document</param>
    /// <returns>List of slide information</returns>
    private List<SlideInfo> AnalyzeTemplateSlides(PresentationDocument presentationDocument)
    {
        var slideInfos = new List<SlideInfo>();
        var templateAnalyzer = new TemplateAnalyzer();
        
        if (presentationDocument.PresentationPart?.Presentation?.SlideIdList == null)
            return slideInfos;
        
        var slideIds = presentationDocument.PresentationPart.Presentation.SlideIdList.ChildElements
            .OfType<DocumentFormat.OpenXml.Presentation.SlideId>();
        
        int slideIndex = 0;
        foreach (var slideId in slideIds)
        {
            try
            {                string? relationshipId = slideId.RelationshipId?.Value;
                if (string.IsNullOrEmpty(relationshipId))
                {
                    slideIndex++;
                    continue;
                }
                
                var slidePart = (SlidePart)presentationDocument.PresentationPart.GetPartById(relationshipId);
                var slideNotes = GetSlideNotes(slidePart);
                
                // Use the TemplateAnalyzer.Analyze method
                var slideInfo = templateAnalyzer.Analyze(slidePart.Slide, slideNotes, slideIndex);
                slideInfos.Add(slideInfo);
                
                if (options.EnableVerboseLogging)
                {
                    Logger.Debug($"Analyzed slide {slideIndex}: {slideInfo.BindingExpressions?.Count ?? 0} expressions, {slideInfo.Directives?.Count ?? 0} directives");
                }
            }
            catch (Exception ex)
            {
                if (options.EnableVerboseLogging)
                {
                    Logger.Warning($"Error analyzing slide {slideIndex}: {ex.Message}");
                }
                
                // Create a basic slide info on error
                slideInfos.Add(new SlideInfo { SlideId = slideIndex, Type = SlideType.Static });
            }
            
            slideIndex++;
        }
        
        return slideInfos;
    }

    /// <summary>
    /// Gets the notes content for a slide
    /// </summary>
    /// <param name="slidePart">The slide part</param>
    /// <returns>The notes content or empty string if no notes</returns>
    private string GetSlideNotes(SlidePart slidePart)
    {
        try
        {
            if (slidePart.NotesSlidePart?.NotesSlide?.CommonSlideData?.ShapeTree != null)
            {
                var textElements = slidePart.NotesSlidePart.NotesSlide.Descendants<DocumentFormat.OpenXml.Drawing.Text>();
                return string.Join(" ", textElements.Select(t => t.Text ?? "").Where(t => !string.IsNullOrWhiteSpace(t)));
            }
        }
        catch (Exception ex)
        {
            if (options.EnableVerboseLogging)
            {
                Logger.Warning($"Error reading slide notes: {ex.Message}");
            }
        }
        
        return string.Empty;
    }

    /// <summary>
    /// Extracts all expressions from a slide (optimized method from PowerPointProcessor)
    /// </summary>
    private HashSet<string> ExtractExpressionsFromSlide(SlidePart slidePart)
    {
        var expressions = new HashSet<string>();
        
        if (slidePart?.Slide == null)
            return expressions;

        try
        {
            // Get all text elements from the slide
            var textElements = slidePart.Slide.Descendants<DocumentFormat.OpenXml.Drawing.Text>();
            
            foreach (var textElement in textElements)
            {
                if (textElement == null || string.IsNullOrEmpty(textElement.Text))
                    continue;

                // Extract expressions using regex
                var regex = new System.Text.RegularExpressions.Regex(@"\$\{([^}]+)\}", System.Text.RegularExpressions.RegexOptions.Compiled);
                var matches = regex.Matches(textElement.Text);
                
                foreach (System.Text.RegularExpressions.Match match in matches)
                {
                    var expressionContent = match.Groups[1].Value.Trim();
                    if (!string.IsNullOrEmpty(expressionContent))
                    {
                        expressions.Add(expressionContent);
                    }
                }
            }
        }
        catch (Exception ex)
        {
            Logger.Warning($"Error extracting expressions from slide: {ex.Message}");
        }

        return expressions;
    }
}
