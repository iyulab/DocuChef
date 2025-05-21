namespace DocuChef.Presentation.Directives;

/// <summary>
/// Represents an if directive for conditional inclusion of a slide
/// </summary>
internal class IfDirective : Directive
{
    /// <summary>
    /// The condition to evaluate
    /// </summary>
    public string Condition { get; set; }

    /// <summary>
    /// Gets the directive type
    /// </summary>
    public override DirectiveType Type => DirectiveType.If;

    /// <summary>
    /// Evaluates the condition using the provided context
    /// </summary>
    public override bool Evaluate(Models.SlideContext context)
    {
        if (context == null || string.IsNullOrEmpty(Condition))
            return false;

        // Get value from context
        string value = context.GetContextValue(Condition);
        Logger.Debug($"Evaluating if condition: '{Condition}' = '{value}'");

        // Condition is true if the value exists and is truthy
        return !string.IsNullOrEmpty(value) &&
               (value.Equals("true", StringComparison.OrdinalIgnoreCase) ||
               (!value.Equals("false", StringComparison.OrdinalIgnoreCase) &&
               !string.IsNullOrWhiteSpace(value)));
    }

    /// <summary>
    /// Returns a string representation of this directive
    /// </summary>
    public override string ToString()
    {
        return $"If: {Condition}";
    }
}