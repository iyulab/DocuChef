namespace DocuChef.PowerPoint.Processing;

/// <summary>
/// Extension to PowerPointProcessor for utility methods and nested data structure support
/// </summary>
internal partial class PowerPointProcessor
{
    /// <summary>
    /// Process PowerPoint template with nested data structure support
    /// </summary>
    public void ProcessWithNestedData(Dictionary<string, object> variables,
            Dictionary<string, Func<object>> globalVariables,
            Dictionary<string, PowerPointFunction> functions)
    {
        // 1. Initialization
        InitializeContext(variables, globalVariables, functions);

        // 2. Use the NestedDataProcessor to handle nested data structures
        var nestedProcessor = new NestedDataProcessor(this, _context);
        nestedProcessor.Process();
    }
}