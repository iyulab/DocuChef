namespace DocuChef.Presentation.Models;

/// <summary>
/// Defines the operation to perform on a slide
/// </summary>
public enum SlideOperation
{
    /// <summary>
    /// Keep the original slide
    /// </summary>
    Keep,

    /// <summary>
    /// Clone the slide for a new data item
    /// </summary>
    Clone,

    /// <summary>
    /// Skip/remove the slide
    /// </summary>
    Skip
}
