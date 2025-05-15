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

}