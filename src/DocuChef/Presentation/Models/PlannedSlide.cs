namespace DocuChef.Presentation.Models;

/// <summary>
/// Represents an operation planned for a slide
/// </summary>
public class PlannedSlide
{
    /// <summary>
    /// Source slide information
    /// </summary>
    public SlideInfo Source { get; set; }

    /// <summary>
    /// Data context for binding
    /// </summary>
    public SlideContext Context { get; set; }

    /// <summary>
    /// Operation to perform on this slide
    /// </summary>
    public SlideOperation Operation { get; set; }

    /// <summary>
    /// The resulting type of this slide after planning
    /// </summary>
    public SlideType ResultingType
    {
        get
        {
            switch (Operation)
            {
                case SlideOperation.Keep:
                    return SlideType.Original;
                case SlideOperation.Clone:
                    return SlideType.Cloned;
                case SlideOperation.Skip:
                    return SlideType.Skipped;
                default:
                    return SlideType.Original;
            }
        }
    }

    /// <summary>
    /// Indicates whether this slide has data context
    /// </summary>
    public bool HasContext => Context != null;

    /// <summary>
    /// Indicates whether this slide will be included in the final presentation
    /// </summary>
    public bool IsIncluded => Operation != SlideOperation.Skip;

    /// <summary>
    /// Creates a new instance with pre-initialized values
    /// </summary>
    public static PlannedSlide Create(
        SlideInfo source,
        SlideContext context,
        SlideOperation operation)
    {
        return new PlannedSlide
        {
            Source = source,
            Context = context,
            Operation = operation
        };
    }

    /// <summary>
    /// Returns a string representation of this planned slide
    /// </summary>
    public override string ToString()
    {
        return $"{Operation} - {(Context != null ? Context.GetContextDescription() : "No context")}";
    }
}