using DocuChef.PowerPoint.Helpers;

namespace DocuChef.PowerPoint.Processing.Directives;

/// <summary>
/// Handler for the if directive in PowerPoint templates
/// </summary>
internal class IfDirectiveHandler : IDirectiveHandler
{
    private readonly IExpressionEvaluator _expressionEvaluator;

    /// <summary>
    /// Creates a new instance of IfDirectiveHandler
    /// </summary>
    /// <param name="expressionEvaluator">The expression evaluator</param>
    public IfDirectiveHandler(IExpressionEvaluator expressionEvaluator)
    {
        _expressionEvaluator = expressionEvaluator ?? throw new ArgumentNullException(nameof(expressionEvaluator));
    }

    /// <summary>
    /// Checks if this handler can process the given directive
    /// </summary>
    public bool CanHandle(SlideDirective directive)
    {
        return directive != null &&
               directive.Name.ToLowerInvariant() == "if";
    }

    /// <summary>
    /// Processes the if directive
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

        string condition = directive.Value.Trim();
        if (string.IsNullOrEmpty(condition))
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

        Logger.Debug($"Processing if directive: condition='{condition}', target='{target}'");

        // Evaluate condition using ExpressionEvaluator
        bool conditionResult = EvaluateCondition(condition, variables);
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
                PowerPointShapeHelper.ShowShape(shape);
                Logger.Debug($"Showing shape: {target}");
            }
            else
            {
                PowerPointShapeHelper.HideShape(shape);
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
                    PowerPointShapeHelper.ShowShape(shape);
                    Logger.Debug($"Showing inverse shape: {visibleWhenFalse}");
                }
                else
                {
                    PowerPointShapeHelper.HideShape(shape);
                    Logger.Debug($"Hiding inverse shape: {visibleWhenFalse}");
                }
            }
        }

        result.WasProcessed = true;
        return result;
    }

    /// <summary>
    /// Evaluate a condition to boolean
    /// </summary>
    private bool EvaluateCondition(string condition, Dictionary<string, object> variables)
    {
        try
        {
            // Use the expression evaluator to evaluate the condition
            var result = _expressionEvaluator.EvaluateCompleteExpression(condition, variables);

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
        catch (Exception ex)
        {
            Logger.Warning($"Error evaluating condition '{condition}': {ex.Message}");
            return false;
        }
    }

    /// <summary>
    /// Find shapes by name
    /// </summary>
    private IEnumerable<P.Shape> FindShapesByName(SlidePart slidePart, string name)
    {
        var shapes = new List<P.Shape>();

        foreach (var shape in slidePart.Slide.Descendants<P.Shape>())
        {
            string shapeName = shape.GetShapeName();
            if (shapeName == name)
            {
                shapes.Add(shape);
            }
        }

        return shapes;
    }
}