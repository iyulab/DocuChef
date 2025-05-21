using DocuChef.Presentation.Directives;

namespace DocuChef.Presentation.Models;

/// <summary>
/// Represents a slide information from the template
/// </summary>
public class SlideInfo
{
    /// <summary>
    /// Gets or sets the slide ID
    /// </summary>
    public uint SlideId { get; set; }

    /// <summary>
    /// Gets or sets the relationship ID
    /// </summary>
    public string RelationshipId { get; set; }

    /// <summary>
    /// Gets or sets the slide note text
    /// </summary>
    public string NoteText { get; set; }

    /// <summary>
    /// Gets or sets the directive for this slide
    /// </summary>
    public Directive Directive { get; set; }

    /// <summary>
    /// Gets a value indicating whether this slide has a directive
    /// </summary>
    public bool HasDirective => Directive != null;

    /// <summary>
    /// Gets or sets a value indicating whether this slide has an implicit directive derived from expressions
    /// </summary>
    public bool HasImplicitDirective { get; set; }

    /// <summary>
    /// Gets the slide type based on directive
    /// </summary>
    public SlideType Type
    {
        get
        {
            if (Directive == null)
                return SlideType.Regular;

            return Directive.Type == DirectiveType.Foreach
                ? SlideType.Foreach
                : SlideType.If;
        }
    }

    /// <summary>
    /// Returns a string representation of this slide info
    /// </summary>
    public override string ToString()
    {
        if (!HasDirective)
            return $"Slide {SlideId} (Regular): No directive";

        string directiveSource = HasImplicitDirective ? "Implicit" : "Explicit";
        return $"Slide {SlideId} ({Type}): {directiveSource} {Directive}";
    }
}