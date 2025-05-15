using System.Globalization;
using DocuChef.PowerPoint.DollarSignEngine;
using DocuChef.PowerPoint.Processing;

namespace DocuChef.PowerPoint;

/// <summary>
/// Factory class for creating specialized PowerPoint processors
/// </summary>
internal static class ProcessorFactory
{
    /// <summary>
    /// Create a custom expression evaluator with specific culture settings
    /// </summary>
    public static IExpressionEvaluator CreateExpressionEvaluator(CultureInfo cultureInfo = null)
    {
        // Create a lightweight processor just for expression evaluation
        var evaluator = new ExpressionEvaluator(cultureInfo);
        return new ExpressionOnlyProcessor(evaluator);
    }

    /// <summary>
    /// Lightweight processor that only handles expression evaluation
    /// </summary>
    private class ExpressionOnlyProcessor : IExpressionEvaluator
    {
        private readonly ExpressionEvaluator _evaluator;

        public ExpressionOnlyProcessor(ExpressionEvaluator evaluator)
        {
            _evaluator = evaluator ?? throw new ArgumentNullException(nameof(evaluator));
        }

        public object EvaluateCompleteExpression(string expression, Dictionary<string, object> variables)
        {
            try
            {
                return _evaluator.Evaluate(expression, variables);
            }
            catch (Exception ex)
            {
                Logger.Error($"Error evaluating expression '{expression}': {ex.Message}", ex);
                return $"[Error: {ex.Message}]";
            }
        }
    }
}