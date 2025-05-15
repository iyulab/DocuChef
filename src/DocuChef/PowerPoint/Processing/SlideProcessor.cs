using DocuChef.PowerPoint.Helpers;
using DocuChef.PowerPoint.Processing.ArrayProcessing;

namespace DocuChef.PowerPoint.Processing;

/// <summary>
/// Processor responsible for slide analysis and preparation
/// </summary>
internal partial class SlideProcessor
{
    private readonly PowerPointProcessor _mainProcessor;
    private readonly PowerPointContext _context;

    /// <summary>
    /// Initialize slide processor
    /// </summary>
    public SlideProcessor(PowerPointProcessor mainProcessor, PowerPointContext context)
    {
        _mainProcessor = mainProcessor ?? throw new ArgumentNullException(nameof(mainProcessor));
        _context = context ?? throw new ArgumentNullException(nameof(context));
    }

    /// <summary>
    /// Analyze slide and prepare duplicates if needed
    /// </summary>
    public void AnalyzeAndPrepareSlide(PresentationPart presentationPart, SlidePart slidePart)
    {
        string slideId = presentationPart.GetIdOfPart(slidePart);
        Logger.Debug($"Analyzing slide {slideId} for array references and directives");

        var variables = _mainProcessor.PrepareVariables();

        // First check for slide directives in notes
        var directiveProcessor = new DirectiveProcessor(_context, variables);
        var dirResult = directiveProcessor.ProcessDirectives(presentationPart, slidePart);

        if (dirResult.WasProcessed)
        {
            // Slide was already processed by directives, no need for further processing
            Logger.Debug($"Slide {slideId} processed by directives");
            return;
        }

        // If no directives found, use automatic array reference detection
        AutoDetectAndProcessArrayReferences(presentationPart, slidePart, variables);
    }

    /// <summary>
    /// Auto-detect array references and process the slide accordingly
    /// </summary>
    private void AutoDetectAndProcessArrayReferences(
        PresentationPart presentationPart,
        SlidePart slidePart,
        Dictionary<string, object> variables)
    {
        string slideId = presentationPart.GetIdOfPart(slidePart);
        Logger.Debug($"Using automatic array reference detection for slide {slideId}");

        // Find array references in the slide
        var arrayReferences = FindArrayReferencesInSlide(slidePart);
        if (!arrayReferences.Any())
        {
            Logger.Debug("No array references found in slide");
            return;
        }

        // Group references by array name
        foreach (var arrayGroup in arrayReferences.GroupBy(r => r.ArrayName))
        {
            string arrayName = arrayGroup.Key;

            Logger.Debug($"Found array '{arrayName}' references in slide");

            // Create auto-detect parameters
            var parameters = ArrayBatchParameters.CreateAutoDetect(arrayName);

            // Process using the common batch processor
            var processor = new ArrayBatchProcessor(_context, variables);
            var result = processor.ProcessArrayBatch(presentationPart, slidePart, parameters);

            if (result.WasProcessed)
            {
                Logger.Debug($"Successfully processed array references for '{arrayName}'");
            }
        }
    }

    /// <summary>
    /// Find all array references in a slide
    /// </summary>
    private List<ArrayReference> FindArrayReferencesInSlide(SlidePart slidePart)
    {
        var result = new List<ArrayReference>();

        foreach (var shape in slidePart.Slide.Descendants<P.Shape>())
        {
            if (shape.TextBody == null)
                continue;

            var references = PowerPointShapeHelper.FindArrayReferences(shape);
            result.AddRange(references);
        }

        return result;
    }
}