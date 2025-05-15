using DocuChef.PowerPoint.Processing.ArrayProcessing;

namespace DocuChef.PowerPoint.Processing.Directives;

/// <summary>
/// Handler for the foreach directive in PowerPoint templates
/// </summary>
internal class ForeachDirectiveHandler : IDirectiveHandler
{
    /// <summary>
    /// Checks if this handler can process the given directive
    /// </summary>
    public bool CanHandle(SlideDirective directive)
    {
        return directive != null &&
               directive.Name.ToLowerInvariant() == "foreach";
    }

    /// <summary>
    /// Processes the foreach directive
    /// </summary>
    public SlideProcessingResult Process(
        PresentationPart presentationPart,
        SlidePart slidePart,
        SlideDirective directive,
        PowerPointContext context,
        Dictionary<string, object> variables)
    {
        var result = new SlideProcessingResult
        {
            SlidePart = slidePart,
            WasProcessed = false
        };

        string collectionName = directive.Value.Trim();
        if (string.IsNullOrEmpty(collectionName))
        {
            Logger.Warning("Foreach directive missing collection name");
            return result;
        }

        Logger.Debug($"Processing foreach directive for collection: {collectionName}");

        // Extract parameters from directive
        int maxItemsPerSlide = directive.GetParameterAsInt("max", -1);
        int offset = directive.GetParameterAsInt("offset", 0);

        // Create parameters object with explicit values from directive
        var parameters = ArrayBatchParameters.CreateExplicit(collectionName, maxItemsPerSlide, offset);

        // Delegate actual processing to the ArrayBatchProcessor
        var processor = new ArrayBatchProcessor(context, variables);
        return processor.ProcessArrayBatch(presentationPart, slidePart, parameters);
    }
}