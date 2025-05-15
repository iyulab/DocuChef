using DocuChef.PowerPoint.Processing.Directives;

namespace DocuChef.PowerPoint.Processing;

/// <summary>
/// Processor for slide directives in PowerPoint templates
/// </summary>
internal class DirectiveProcessor
{
    private readonly PowerPointContext _context;
    private readonly Dictionary<string, object> _variables;
    private readonly List<IDirectiveHandler> _handlers;

    /// <summary>
    /// Initialize a new instance of DirectiveProcessor
    /// </summary>
    public DirectiveProcessor(PowerPointContext context, Dictionary<string, object> variables)
    {
        _context = context ?? throw new ArgumentNullException(nameof(context));
        _variables = variables ?? throw new ArgumentNullException(nameof(variables));

        // Initialize handlers
        _handlers = new List<IDirectiveHandler>
        {
            new ForeachDirectiveHandler(),
            new IfDirectiveHandler(ProcessorFactory.CreateExpressionEvaluator())
        };
    }

    /// <summary>
    /// Process directives for a slide
    /// </summary>
    public SlideProcessingResult ProcessDirectives(PresentationPart presentationPart, SlidePart slidePart)
    {
        var result = new SlideProcessingResult
        {
            SlidePart = slidePart,
            WasProcessed = false
        };

        try
        {
            // Get slide notes
            string notes = slidePart.GetNotes();
            if (string.IsNullOrEmpty(notes))
                return result;

            // Parse directives from notes
            var directives = DirectiveParser.ParseDirectives(notes);
            if (!directives.Any())
                return result;

            Logger.Debug($"Found {directives.Count} directives in slide");

            // Process each directive
            foreach (var directive in directives)
            {
                var directiveResult = ProcessDirective(presentationPart, slidePart, directive);
                if (directiveResult.WasProcessed)
                {
                    result.WasProcessed = true;
                    result.GeneratedSlides.AddRange(directiveResult.GeneratedSlides);
                }
            }
        }
        catch (Exception ex)
        {
            Logger.Error($"Error processing directives: {ex.Message}", ex);
        }

        return result;
    }

    /// <summary>
    /// Process a single directive
    /// </summary>
    private SlideProcessingResult ProcessDirective(PresentationPart presentationPart, SlidePart slidePart, SlideDirective directive)
    {
        // Find a handler that can process this directive
        var handler = _handlers.FirstOrDefault(h => h.CanHandle(directive));

        if (handler != null)
        {
            Logger.Debug($"Found handler for directive: {directive.Name} => {handler.GetType().Name}");
            return handler.Process(presentationPart, slidePart, directive, _context, _variables);
        }

        Logger.Warning($"No handler found for directive: {directive.Name}");
        return new SlideProcessingResult
        {
            SlidePart = slidePart,
            WasProcessed = false
        };
    }
}