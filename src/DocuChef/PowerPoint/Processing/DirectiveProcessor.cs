using DocuChef.PowerPoint.Helpers;

namespace DocuChef.PowerPoint.Processing;

/// <summary>
/// Interface for hierarchical directive handlers
/// </summary>
internal interface IDirectiveHandler
{
    /// <summary>
    /// Checks if this handler can process the given directive
    /// </summary>
    bool CanHandle(Directive directive);

    /// <summary>
    /// Processes the directive
    /// </summary>
    SlideProcessingResult Process(
        PresentationPart presentationPart,
        SlidePart slidePart,
        Directive directive,
        PowerPointContext context,
        Dictionary<string, object> variables);
}

/// <summary>
/// Processor for slide directives in PowerPoint templates with hierarchical path support
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
            new ForeachHandler(),
            new IfHandler(),
            new RepeatHandler()
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

            Logger.Debug($"Found {directives.Count} hierarchical directives in slide");

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
            Logger.Error($"Error processing hierarchical directives: {ex.Message}", ex);
        }

        return result;
    }

    /// <summary>
    /// Process a single directive
    /// </summary>
    private SlideProcessingResult ProcessDirective(PresentationPart presentationPart, SlidePart slidePart, Directive directive)
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

/// <summary>
/// Handler for hierarchical foreach directive
/// </summary>
internal class ForeachHandler : IDirectiveHandler
{
    /// <summary>
    /// Checks if this handler can process the given directive
    /// </summary>
    public bool CanHandle(Directive directive)
    {
        return directive != null &&
               directive.Name.Equals("foreach", StringComparison.OrdinalIgnoreCase) &&
               !string.IsNullOrEmpty(directive.Value);
    }

    /// <summary>
    /// Processes the hierarchical foreach directive
    /// </summary>
    public SlideProcessingResult Process(
        PresentationPart presentationPart,
        SlidePart slidePart,
        Directive directive,
        PowerPointContext context,
        Dictionary<string, object> variables)
    {
        Logger.Debug($"Processing foreach directive: {directive.Value}");

        // Create path if not already parsed
        if (directive.Path == null || directive.Path.Segments.Count == 0)
        {
            directive.Path = new HierarchicalPath(directive.Value);
            Logger.Debug($"Created hierarchical path from value: {directive.Value}");
        }

        // Delegate to the SlideHierarchyProcessor
        var processor = new SlideHierarchyProcessor(context, variables);
        return processor.ProcessHierarchicalForeach(presentationPart, slidePart, directive);
    }
}

/// <summary>
/// Handler for hierarchical if directive
/// </summary>
internal class IfHandler : IDirectiveHandler
{
    /// <summary>
    /// Checks if this handler can process the given directive
    /// </summary>
    public bool CanHandle(Directive directive)
    {
        return directive != null &&
               directive.Name.Equals("if", StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// Processes the hierarchical if directive
    /// </summary>
    public SlideProcessingResult Process(
        PresentationPart presentationPart,
        SlidePart slidePart,
        Directive directive,
        PowerPointContext context,
        Dictionary<string, object> variables)
    {
        var result = new SlideProcessingResult
        {
            SlidePart = slidePart,
            WasProcessed = false
        };

        if (string.IsNullOrEmpty(directive.Value))
        {
            Logger.Warning("If directive missing condition");
            return result;
        }

        string target = directive.GetParameter("target");
        if (string.IsNullOrEmpty(target))
        {
            Logger.Warning("If directive missing target parameter");
            return result;
        }

        string condition = directive.Value.Trim();
        Logger.Debug($"Processing hierarchical if directive: condition='{condition}', target='{target}'");

        // Evaluate hierarchical condition
        bool conditionResult = EvaluateHierarchicalCondition(condition, context, variables);
        Logger.Debug($"Condition evaluated to: {conditionResult}");

        // Find target shapes
        var targetShapes = FindShapesByName(slidePart, target);
        if (!targetShapes.Any())
        {
            Logger.Warning($"Target shape not found: {target}");
            return result;
        }

        // Set visibility based on condition result
        foreach (var shape in targetShapes)
        {
            if (conditionResult)
            {
                ShapeHelper.ShowShape(shape);
                Logger.Debug($"Showing shape: {target}");
            }
            else
            {
                ShapeHelper.HideShape(shape);
                Logger.Debug($"Hiding shape: {target}");
            }
        }

        // Process visibleWhenFalse parameter if specified
        if (directive.HasParameter("visibleWhenFalse"))
        {
            string visibleWhenFalse = directive.GetParameter("visibleWhenFalse");
            var inverseTargetShapes = FindShapesByName(slidePart, visibleWhenFalse);

            foreach (var shape in inverseTargetShapes)
            {
                if (!conditionResult)
                {
                    ShapeHelper.ShowShape(shape);
                    Logger.Debug($"Showing inverse shape: {visibleWhenFalse}");
                }
                else
                {
                    ShapeHelper.HideShape(shape);
                    Logger.Debug($"Hiding inverse shape: {visibleWhenFalse}");
                }
            }
        }

        result.WasProcessed = true;
        return result;
    }

    /// <summary>
    /// Evaluate a hierarchical condition to boolean
    /// </summary>
    private bool EvaluateHierarchicalCondition(string condition, PowerPointContext context, Dictionary<string, object> variables)
    {
        try
        {
            // Ensure context navigator is initialized
            if (context.Navigator == null)
                context.InitializeNavigator();

            // Check if condition contains a hierarchical path
            if (condition.Contains('.') || condition.Contains('[') || condition.Contains('_'))
            {
                // Try to resolve value directly before wrapping in expression
                var result = context.ResolveHierarchicalValue(condition);

                // Convert result to boolean
                if (result is bool boolResult)
                    return boolResult;

                // Handle string results
                if (result is string stringResult)
                {
                    stringResult = stringResult.Trim().ToLowerInvariant();
                    return stringResult == "true" || stringResult == "yes" || stringResult == "1";
                }

                // Handle numeric results
                if (result is int intResult)
                    return intResult != 0;

                if (result is double doubleResult)
                    return doubleResult != 0;

                // Non-null result is generally considered true
                return result != null;
            }
            else
            {
                // For simple conditions, check variables directly
                if (variables.TryGetValue(condition, out var value))
                {
                    if (value is bool boolValue)
                        return boolValue;

                    if (value is string strValue)
                        return !string.IsNullOrEmpty(strValue) &&
                               strValue.ToLower() != "false" &&
                               strValue.ToLower() != "0";

                    return value != null;
                }

                // Try parsing as boolean constants
                if (bool.TryParse(condition, out bool boolResult))
                    return boolResult;

                // Handle numeric values
                if (int.TryParse(condition, out int intValue))
                    return intValue != 0;

                // Default for unknown conditions
                return false;
            }
        }
        catch (Exception ex)
        {
            Logger.Warning($"Error evaluating hierarchical condition '{condition}': {ex.Message}");
            return false;
        }
    }

    /// <summary>
    /// Find shapes by name
    /// </summary>
    private IEnumerable<P.Shape> FindShapesByName(SlidePart slidePart, string name)
    {
        if (slidePart?.Slide == null)
            return Enumerable.Empty<P.Shape>();

        return slidePart.Slide.Descendants<P.Shape>()
            .Where(s => s.GetShapeName() == name);
    }
}

/// <summary>
/// Handler for repeat directive
/// </summary>
internal class RepeatHandler : IDirectiveHandler
{
    /// <summary>
    /// Checks if this handler can process the given directive
    /// </summary>
    public bool CanHandle(Directive directive)
    {
        return directive != null &&
               directive.Name.Equals("repeat", StringComparison.OrdinalIgnoreCase) &&
               !string.IsNullOrEmpty(directive.Value);
    }

    /// <summary>
    /// Processes the repeat directive
    /// </summary>
    public SlideProcessingResult Process(
        PresentationPart presentationPart,
        SlidePart slidePart,
        Directive directive,
        PowerPointContext context,
        Dictionary<string, object> variables)
    {
        var result = new SlideProcessingResult
        {
            SlidePart = slidePart,
            WasProcessed = false
        };

        try
        {
            // Parse repeat count
            if (!int.TryParse(directive.Value, out int repeatCount) || repeatCount < 1)
            {
                Logger.Warning($"Invalid repeat count: {directive.Value}");
                return result;
            }

            // Maximum repeat limit
            int maxRepeat = context.Options?.MaxSlidesFromTemplate ?? 100;
            repeatCount = Math.Min(repeatCount, maxRepeat);

            Logger.Debug($"Processing repeat directive, count: {repeatCount}");

            // Slides generated so far
            var generatedSlides = new List<SlidePart>();

            // Store the original slide position
            int slidePosition = SlideHelper.FindSlidePosition(presentationPart, slidePart);

            // Repeat the slide
            for (int i = 1; i < repeatCount; i++) // Start from 1 because we already have one slide
            {
                // Clone the original slide
                var newSlidePart = SlideHelper.CloneSlide(presentationPart, slidePart);

                // Insert after the previously generated slides
                int insertPosition = slidePosition + i;
                SlideHelper.InsertSlide(presentationPart, newSlidePart, insertPosition);

                Logger.Debug($"Created repeat slide {i + 1} of {repeatCount} at position {insertPosition}");
                generatedSlides.Add(newSlidePart);
            }

            if (generatedSlides.Count > 0)
            {
                result.WasProcessed = true;
                result.GeneratedSlides.AddRange(generatedSlides);
            }

            return result;
        }
        catch (Exception ex)
        {
            Logger.Error($"Error processing repeat directive: {ex.Message}", ex);
            return result;
        }
    }
}