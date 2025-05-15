namespace DocuChef.PowerPoint.Processing;

/// <summary>
/// Processor responsible for handling expressions in templates
/// </summary>
internal class ExpressionProcessor
{
    private readonly PowerPointProcessor _mainProcessor;
    private readonly PowerPointContext _context;

    /// <summary>
    /// Initialize expression processor
    /// </summary>
    public ExpressionProcessor(PowerPointProcessor mainProcessor, PowerPointContext context)
    {
        _mainProcessor = mainProcessor ?? throw new ArgumentNullException(nameof(mainProcessor));
        _context = context ?? throw new ArgumentNullException(nameof(context));
    }

    /// <summary>
    /// Process expressions in text
    /// </summary>
    public string ProcessExpressions(string text)
    {
        if (!DocuChef.Helpers.ExpressionHelper.ContainsExpressions(text))
            return text;

        var variables = _mainProcessor.PrepareVariables();
        return DocuChef.Helpers.ExpressionHelper.ProcessExpressions(text, _mainProcessor, variables);
    }

    /// <summary>
    /// Evaluate an expression directly
    /// </summary>
    public object EvaluateExpression(string expression)
    {
        if (string.IsNullOrEmpty(expression))
            return null;

        var variables = _mainProcessor.PrepareVariables();
        return _mainProcessor.EvaluateCompleteExpression(expression, variables);
    }

    /// <summary>
    /// Evaluate a condition to boolean
    /// </summary>
    public bool EvaluateCondition(string condition)
    {
        if (string.IsNullOrEmpty(condition))
            return false;

        object result = EvaluateExpression(condition);
        return ConvertToBoolean(result);
    }

    /// <summary>
    /// Convert a result object to boolean
    /// </summary>
    private bool ConvertToBoolean(object result)
    {
        if (result is bool boolValue)
            return boolValue;

        if (result != null)
        {
            try
            {
                if (result is string stringValue)
                {
                    if (string.IsNullOrEmpty(stringValue))
                        return false;

                    // Handle common string representations
                    stringValue = stringValue.Trim().ToLowerInvariant();
                    return stringValue == "true" || stringValue == "yes" || stringValue == "1";
                }

                return Convert.ToBoolean(result);
            }
            catch
            {
                // If conversion fails, non-null/non-empty values are considered true
                return true;
            }
        }

        return false;
    }
}