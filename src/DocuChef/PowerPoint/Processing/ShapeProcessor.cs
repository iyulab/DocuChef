namespace DocuChef.PowerPoint.Processing;

/// <summary>
/// Processor responsible for handling shape-related operations with improved functionality
/// </summary>
internal class ShapeProcessor
{
    private readonly IExpressionEvaluator _evaluator;
    private readonly PowerPointContext _context;
    private readonly ExpressionProcessor _expressionProcessor;
    private static readonly Regex PowerPointFunctionPattern = new(@"\${ppt\.(\w+)\(([^)]*)\)}", RegexOptions.Compiled);

    /// <summary>
    /// Initialize shape processor
    /// </summary>
    public ShapeProcessor(IExpressionEvaluator evaluator, PowerPointContext context)
    {
        _evaluator = evaluator ?? throw new ArgumentNullException(nameof(evaluator));
        _context = context ?? throw new ArgumentNullException(nameof(context));
        _expressionProcessor = new ExpressionProcessor(evaluator, context);
    }

    /// <summary>
    /// Process a shape with all its expressions and functions
    /// </summary>
    public bool ProcessShape(P.Shape shape, Dictionary<string, object> variables)
    {
        if (shape.TextBody == null)
            return false;

        bool hasChanges = false;

        // Update shape context
        UpdateShapeContext(shape);

        // Process expression bindings
        if (shape.ContainsExpressions())
        {
            if (_expressionProcessor.ProcessShapeExpressions(shape, variables))
            {
                hasChanges = true;
                Logger.Debug($"Processed expressions in shape '{shape.GetShapeName()}'");
            }
        }

        // Process PowerPoint functions
        if (ProcessPowerPointFunctions(shape, variables))
        {
            hasChanges = true;
            Logger.Debug($"Processed PowerPoint functions in shape '{shape.GetShapeName()}'");
        }

        return hasChanges;
    }

    /// <summary>
    /// Update shape context
    /// </summary>
    private void UpdateShapeContext(P.Shape shape)
    {
        _context.Shape = new ShapeContext
        {
            Name = shape.GetShapeName(),
            Id = shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Id?.Value.ToString(),
            Text = shape.GetText(),
            Type = GetShapeType(shape),
            ShapeObject = shape
        };
    }

    /// <summary>
    /// Get shape type
    /// </summary>
    private string GetShapeType(P.Shape shape)
    {
        var presetGeometry = shape.ShapeProperties?.GetFirstChild<A.PresetGeometry>();
        if (presetGeometry?.Preset != null)
        {
            return presetGeometry.Preset.Value.ToString();
        }

        return shape.TextBody != null ? "TextBox" : "Shape";
    }

    /// <summary>
    /// Process PowerPoint functions in shape
    /// </summary>
    public bool ProcessPowerPointFunctions(P.Shape shape, Dictionary<string, object> variables)
    {
        if (shape.TextBody == null)
            return false;

        bool hasChanges = false;

        foreach (var paragraph in shape.TextBody.Elements<A.Paragraph>().ToList())
        {
            foreach (var run in paragraph.Elements<A.Run>().ToList())
            {
                var textElement = run.GetFirstChild<A.Text>();
                if (textElement == null || string.IsNullOrEmpty(textElement.Text))
                    continue;

                string text = textElement.Text;
                var matches = PowerPointFunctionPattern.Matches(text);

                if (matches.Count == 0)
                    continue;

                // Process function if it's the entire text
                if (matches.Count == 1 && matches[0].Value == text)
                {
                    var match = matches[0];
                    string functionName = match.Groups[1].Value;
                    string parameters = match.Groups[2].Value;

                    if (ProcessPowerPointFunction(functionName, parameters, shape))
                    {
                        hasChanges = true;
                    }
                }
                else
                {
                    // Process mixed content with expressions and functions
                    string processedText = _expressionProcessor.ProcessExpressions(text, variables);
                    if (processedText != text)
                    {
                        textElement.Text = processedText;
                        hasChanges = true;
                    }
                }
            }
        }

        return hasChanges;
    }

    /// <summary>
    /// Process a single PowerPoint function
    /// </summary>
    private bool ProcessPowerPointFunction(string functionName, string parametersString, P.Shape shape)
    {
        Logger.Debug($"Processing PowerPoint function: {functionName}({parametersString})");

        // Find the function
        if (!_context.Functions.TryGetValue(functionName, out var function))
        {
            Logger.Warning($"Function not found: {functionName}");
            return false;
        }

        // Update shape context
        _context.Shape.ShapeObject = shape;

        // Parse parameters
        var parameters = DocuChef.Helpers.ExpressionHelper.ParseFunctionParameters(parametersString);

        try
        {
            // Execute function
            var result = function.Execute(_context, null, parameters);

            // Handle result
            if (result is string resultText)
            {
                var textElements = shape.Descendants<A.Text>()
                    .Where(t => t.Text.Contains($"${{ppt.{functionName}"))
                    .ToList();

                foreach (var textElement in textElements)
                {
                    textElement.Text = resultText;
                }

                return true;
            }

            // Function may have modified shape directly
            return true;
        }
        catch (Exception ex)
        {
            Logger.Error($"Error executing function {functionName}: {ex.Message}", ex);
            return false;
        }
    }
}