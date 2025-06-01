using DocumentFormat.OpenXml.Packaging;
using DocuChef.Logging;
using DocuChef.Presentation.Models;

namespace DocuChef.Presentation.Processors;

/// <summary>
/// Handles slide removal operations
/// Extracted from SlideGenerator to improve code organization
/// </summary>
internal class SlideRemover
{
    /// <summary>
    /// Removes original template slides that have been repositioned
    /// </summary>
    public void RemoveOriginalSlides(PresentationPart presentationPart, List<DocumentFormat.OpenXml.Presentation.SlideId> sourceSlides, HashSet<int> slidesToRemove)
    {
        if (!slidesToRemove.Any())
            return;

        var slideIdList = presentationPart.Presentation.SlideIdList;
        if (slideIdList == null)
            return;

        Logger.Debug($"SlideRemover: Removing {slidesToRemove.Count} original slides that were repositioned");

        // Remove slides in reverse order to maintain indices
        foreach (var slideIndex in slidesToRemove.OrderByDescending(x => x))
        {
            RemoveSlideAtIndex(presentationPart, sourceSlides, slideIndex);
        }
    }

    /// <summary>
    /// Removes a slide at the specified index
    /// </summary>
    private void RemoveSlideAtIndex(PresentationPart presentationPart, List<DocumentFormat.OpenXml.Presentation.SlideId> sourceSlides, int slideIndex)
    {
        if (!IsValidSlideIndex(slideIndex, sourceSlides))
            return;

        var slideToRemove = sourceSlides[slideIndex];
        Logger.Debug($"SlideRemover: Removing original slide {slideIndex} with RelationshipId {slideToRemove.RelationshipId?.Value}");

        // Remove the slide from the presentation
        slideToRemove.Remove();

        // Also remove the corresponding slide part
        RemoveSlidePart(presentationPart, slideToRemove, slideIndex);
    }

    /// <summary>
    /// Validates if the slide index is within bounds
    /// </summary>
    private static bool IsValidSlideIndex(int slideIndex, List<DocumentFormat.OpenXml.Presentation.SlideId> sourceSlides)
    {
        return slideIndex >= 0 && slideIndex < sourceSlides.Count;
    }

    /// <summary>
    /// Removes the slide part from the presentation
    /// </summary>
    private void RemoveSlidePart(PresentationPart presentationPart, DocumentFormat.OpenXml.Presentation.SlideId slideToRemove, int slideIndex)
    {
        if (slideToRemove.RelationshipId?.Value == null)
            return;

        try
        {
            var slidePart = (SlidePart)presentationPart.GetPartById(slideToRemove.RelationshipId.Value);
            presentationPart.DeletePart(slidePart);
            Logger.Debug($"SlideRemover: Deleted slide part for original slide {slideIndex}");
        }
        catch (Exception ex)
        {
            Logger.Warning($"SlideRemover: Failed to delete slide part for slide {slideIndex}: {ex.Message}");
        }
    }
}
