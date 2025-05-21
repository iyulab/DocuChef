using DocuChef.Presentation.Processing;

namespace DocuChef.Presentation;

/// <summary>
/// Main entry point for the DocuChef presentation generation engine
/// </summary>
public class PPTChef
{
    private readonly string _templatePath;
    private readonly object _dataSource;
    private List<Models.SlideInfo> _templateSlides;
    private readonly TemplateProcessor _templateProcessor;
    private readonly ContextProcessor _contextProcessor;
    private readonly PlanProcessor _planProcessor;
    private readonly PresentationProcessor _presentationProcessor;

    /// <summary>
    /// Template analysis results
    /// </summary>
    public TemplateAnalysisResult AnalysisResult { get; private set; }

    /// <summary>
    /// Initializes a new instance of the PPTChef engine
    /// </summary>
    public PPTChef(string templatePath, object dataSource)
    {
        _templatePath = templatePath ?? throw new ArgumentNullException(nameof(templatePath));
        _dataSource = dataSource ?? throw new ArgumentNullException(nameof(dataSource));
        _templateSlides = new List<Models.SlideInfo>();
        AnalysisResult = new TemplateAnalysisResult();

        // Initialize processor components
        _templateProcessor = new TemplateProcessor();
        _contextProcessor = new ContextProcessor(dataSource);
        _planProcessor = new PlanProcessor();
        _presentationProcessor = new PresentationProcessor();
    }

    /// <summary>
    /// Analyzes the template to understand its structure
    /// </summary>
    public TemplateAnalysisResult AnalyzeTemplate()
    {
        Logger.Info($"Analyzing template: {_templatePath}");
        _templateSlides.Clear();

        try
        {
            using (PresentationDocument presentationDoc = PresentationDocument.Open(_templatePath, false))
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
            using (PresentationDocument presentationDoc = PresentationDocument.Open(_templatePath, true))
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
    private TemplateAnalysisResult CreateAnalysisResult(List<Models.SlideInfo> slides)
    {
        return new TemplateAnalysisResult
        {
            TotalSlides = slides.Count,
            ForeachSlides = slides.Count(s => s.Type == Models.SlideType.Foreach),
            IfSlides = slides.Count(s => s.Type == Models.SlideType.If),
            RegularSlides = slides.Count(s => s.Type == Models.SlideType.Regular),
            ImplicitDirectives = slides.Count(s => s.HasImplicitDirective)
        };
    }

    /// <summary>
    /// Creates a presentation generation plan with multi-level nesting support
    /// </summary>
    public Models.PresentationPlan CreatePlan()
    {
        if (_templateSlides.Count == 0)
        {
            AnalyzeTemplate();
        }

        Logger.Info("Creating presentation generation plan with multi-level nesting support...");

        // Build the plan using the plan processor
        var plan = _planProcessor.BuildPlan(_templateSlides, _dataSource, _contextProcessor);

        // Log slide details
        LogSlideDetails(plan);

        var summary = plan.GetSummary();
        Logger.Info($"Created presentation plan: {summary}");
        return plan;
    }

    /// <summary>
    /// Logs detailed information about each slide in the plan
    /// </summary>
    private void LogSlideDetails(Models.PresentationPlan plan)
    {
        // Log slide details
        Logger.Info("Slide plan details:");
        for (int i = 0; i < plan.IncludedSlides.Count(); i++)
        {
            var slide = plan.IncludedSlides.ElementAt(i);
            string operation = slide.Operation == Models.SlideOperation.Keep ? "Original" : "Clone";
            string context = slide.HasContext ? slide.Context.GetContextDescription() : "No context";

            Logger.Info($"   - Slide {i + 1}: {operation}, {context}");
        }
    }

    /// <summary>
    /// Generates a presentation based on the template and data
    /// </summary>
    public void GeneratePresentation(string outputPath, Models.PresentationPlan plan = null)
    {
        // Ensure we have a plan
        plan ??= CreatePlan();

        Logger.Info($"Generating presentation to {outputPath}");

        try
        {
            // Create a copy of the template
            File.Copy(_templatePath, outputPath, true);

            // Open the new presentation for editing
            using (PresentationDocument presentationDoc = PresentationDocument.Open(outputPath, true))
            {
                // Execute the presentation generation plan
                var data = _dataSource.GetProperties();
                data.Add("ppt", new PPTMethods());
                _presentationProcessor.ProcessPresentation(presentationDoc, plan, data);
            }

            Logger.Info($"Successfully generated presentation at {outputPath}");
        }
        catch (Exception ex)
        {
            Logger.Error($"Error generating presentation: {ex.Message}", ex);
            throw;
        }
    }

    /// <summary>
    /// Performs the entire process in one step: analyze, plan, and generate
    /// </summary>
    public void GeneratePresentation(string outputPath, bool skipAnalysis = false, bool updateImplicitDirectives = true)
    {
        if (!skipAnalysis || _templateSlides.Count == 0)
        {
            AnalyzeTemplate();
        }

        // Optionally update the template with detected implicit directives
        if (updateImplicitDirectives && AnalysisResult.ImplicitDirectives > 0)
        {
            UpdateTemplateWithImplicitDirectives();
        }

        var plan = CreatePlan();
        GeneratePresentation(outputPath, plan);
    }
}