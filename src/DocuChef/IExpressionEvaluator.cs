namespace DocuChef;

/// <summary>
/// Interface for expression evaluation in PowerPoint processing
/// </summary>
public interface IExpressionEvaluator
{
    /// <summary>
    /// Evaluates a complete expression with the provided variables
    /// </summary>
    object EvaluateCompleteExpression(string expression, Dictionary<string, object> variables);
}