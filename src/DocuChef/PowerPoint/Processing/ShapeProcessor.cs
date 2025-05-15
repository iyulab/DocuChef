namespace DocuChef.PowerPoint.Processing;

/// <summary>
/// Processor responsible for handling shape-related operations
/// </summary>
internal class ShapeProcessor
{
    private readonly PowerPointProcessor _mainProcessor;
    private readonly PowerPointContext _context;
    private readonly ExpressionProcessor _expressionProcessor;

    /// <summary>
    /// Initialize shape processor
    /// </summary>
    public ShapeProcessor(PowerPointProcessor mainProcessor, PowerPointContext context)
    {
        _mainProcessor = mainProcessor ?? throw new ArgumentNullException(nameof(mainProcessor));
        _context = context ?? throw new ArgumentNullException(nameof(context));
        _expressionProcessor = new ExpressionProcessor(mainProcessor, context);
    }

    /// <summary>
    /// Process PowerPoint functions in shape
    /// </summary>
    public bool ProcessPowerPointFunctions(P.Shape shape)
    {
        if (shape.TextBody == null)
            return false;

        bool hasChanges = false;
        var variables = _mainProcessor.PrepareVariables();

        foreach (var paragraph in shape.TextBody.Elements<A.Paragraph>().ToList())
        {
            foreach (var run in paragraph.Elements<A.Run>().ToList())
            {
                var textElement = run.GetFirstChild<A.Text>();
                if (textElement == null || string.IsNullOrEmpty(textElement.Text))
                    continue;

                string text = textElement.Text;
                var functions = ExtractPowerPointFunctions(text);

                if (!functions.Any())
                    continue;

                // Process function if it's the entire text
                if (functions.Count == 1 && DocuChef.Helpers.ExpressionHelper.IsSingleExpression(text))
                {
                    var function = functions[0];
                    if (ProcessPowerPointFunction(function, shape))
                    {
                        hasChanges = true;
                    }
                }
                else
                {
                    // Process mixed content
                    string processedText = _expressionProcessor.ProcessExpressions(text);
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
    /// Extract PowerPoint functions from text
    /// </summary>
    private List<PowerPointFunctionCall> ExtractPowerPointFunctions(string text)
    {
        var result = new List<PowerPointFunctionCall>();
        var pattern = new Regex(@"\${ppt\.(\w+)\(([^)]*)\)}", RegexOptions.Compiled);

        var matches = pattern.Matches(text);
        foreach (Match match in matches)
        {
            result.Add(new PowerPointFunctionCall
            {
                FullMatch = match.Value,
                FunctionName = match.Groups[1].Value,
                Parameters = match.Groups[2].Value
            });
        }

        return result;
    }

    /// <summary>
    /// Process a single PowerPoint function
    /// </summary>
    private bool ProcessPowerPointFunction(PowerPointFunctionCall functionCall, P.Shape shape)
    {
        Logger.Debug($"Processing PowerPoint function: {functionCall.FunctionName}({functionCall.Parameters})");

        // Find the function
        if (!_context.Functions.TryGetValue(functionCall.FunctionName, out var function))
        {
            Logger.Warning($"Function not found: {functionCall.FunctionName}");
            return false;
        }

        // Update shape context
        _context.Shape.ShapeObject = shape;

        // Parse parameters
        var parameters = ParseFunctionParameters(functionCall.Parameters);

        try
        {
            // Execute function
            var result = function.Execute(_context, null, parameters);

            // Handle result
            if (result is string resultText)
            {
                var textElements = shape.Descendants<A.Text>()
                    .Where(t => t.Text == functionCall.FullMatch)
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
            Logger.Error($"Error executing function {functionCall.FunctionName}: {ex.Message}", ex);
            return false;
        }
    }

    /// <summary>
    /// Parse function parameters
    /// </summary>
    private string[] ParseFunctionParameters(string parametersString)
    {
        return DocuChef.Helpers.ExpressionHelper.ParseFunctionParameters(parametersString);
    }

    /// <summary>
    /// Represents a PowerPoint function call
    /// </summary>
    private class PowerPointFunctionCall
    {
        public string FullMatch { get; set; }
        public string FunctionName { get; set; }
        public string Parameters { get; set; }
    }
}