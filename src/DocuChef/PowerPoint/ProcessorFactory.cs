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
    /// Create a PowerPoint processor with all dependent processors
    /// </summary>
    public static PowerPointProcessor CreateProcessor(PresentationDocument document, PowerPointOptions options)
    {
        if (document == null)
            throw new ArgumentNullException(nameof(document));

        if (options == null)
            throw new ArgumentNullException(nameof(options));

        return new PowerPointProcessor(document, options);
    }

    /// <summary>
    /// Create a custom expression evaluator with specific culture settings
    /// </summary>
    public static IExpressionEvaluator CreateExpressionEvaluator(CultureInfo cultureInfo = null)
    {
        var context = new PowerPointContext();
        var processor = new ExpressionEvaluator(cultureInfo ?? CultureInfo.CurrentCulture);

        // Create a lightweight processor just for expression evaluation
        return new ExpressionOnlyProcessor(processor);
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