namespace DocuChef.Presentation.Directives;

/// <summary>
/// Base interface for all directives
/// </summary>
public interface IDirective
{
    /// <summary>
    /// Gets the directive type
    /// </summary>
    DirectiveType Type { get; }

    /// <summary>
    /// Evaluates the directive using the provided context
    /// </summary>
    bool Evaluate(Models.SlideContext context);

    /// <summary>
    /// Returns a string representation of this directive
    /// </summary>
    string ToString();
}

/// <summary>
/// Base class for slide directives
/// </summary>
public abstract class Directive : IDirective
{
    /// <summary>
    /// Gets the directive type
    /// </summary>
    public abstract DirectiveType Type { get; }

    /// <summary>
    /// Evaluates the directive using the provided context
    /// </summary>
    public abstract bool Evaluate(Models.SlideContext context);

    /// <summary>
    /// Returns a string representation of this directive
    /// </summary>
    public override string ToString()
    {
        return $"{Type} Directive";
    }
}

/// <summary>
/// Types of directives
/// </summary>
public enum DirectiveType
{
    /// <summary>
    /// Foreach directive for collections
    /// </summary>
    Foreach,

    /// <summary>
    /// If directive for conditions
    /// </summary>
    If
}