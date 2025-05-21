namespace DocuChef.Presentation.Utils;

/// <summary>
/// Extension methods for OpenXml elements for slide handling
/// </summary>
public static class SlideExtensions
{
    /// <summary>
    /// Gets the slide part containing the specified OpenXml element
    /// </summary>
    public static SlidePart GetSlidePart(this OpenXmlElement element)
    {
        if (element == null)
            return null;

        // Process according to the PowerPoint element type
        if (element is P.Slide slide)
        {
            return slide.SlidePart;
        }

        // Navigate up through parents to find a slide
        var current = element;
        while (current != null)
        {
            if (current is P.Slide parentSlide)
            {
                return parentSlide.SlidePart;
            }

            // Move to the next parent element
            current = current.Parent;
        }

        return null;
    }
}