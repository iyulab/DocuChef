namespace DocuChef.PowerPoint;

/// <summary>
/// Factory class for creating specialized processors
/// </summary>
internal static class ProcessorFactory
{
    /// <summary>
    /// Create a custom expression evaluator
    /// </summary>
    public static IExpressionEvaluator CreateExpressionEvaluator(PowerPointContext context = null)
    {
        // Use the enhanced DollarSignEngine adapter for expression evaluation
        return context != null
            ? new DollarSignEngine.ExpressionEvaluator(context)
            : new DollarSignEngine.ExpressionEvaluator();
    }
}