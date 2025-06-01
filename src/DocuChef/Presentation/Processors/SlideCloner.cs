using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocuChef.Logging;
using DocuChef.Presentation.Exceptions;

namespace DocuChef.Presentation.Processors;

/// <summary>
/// Handles slide cloning operations
/// Extracted from SlideGenerator to improve code organization
/// </summary>
internal class SlideCloner
{
    private const uint BaseSlideId = 255;

    /// <summary>
    /// Clones a slide from the template at the specified position
    /// </summary>
    public SlidePart CloneSlideFromTemplate(PresentationPart presentationPart, SlidePart templateSlidePart, int insertPosition = -1)
    {
        if (presentationPart == null)
            throw new ArgumentNullException(nameof(presentationPart));
        if (templateSlidePart == null)
            throw new ArgumentNullException(nameof(templateSlidePart));

        // Create a new slide part
        var newSlidePart = presentationPart.AddNewPart<SlidePart>();        // Clone the slide content
        newSlidePart.Slide = (Slide)templateSlidePart.Slide.CloneNode(true);

        // Clone all image parts and their relationships
        CloneImageRelationships(templateSlidePart, newSlidePart);

        // Add the new slide to the slide ID list at the specified position
        AddSlideToPresentation(presentationPart, newSlidePart, insertPosition);

        // Clone slide layout relationship if it exists
        CloneSlideLayoutRelationship(templateSlidePart, newSlidePart);

        return newSlidePart;
    }

    /// <summary>
    /// Adds the new slide to the presentation at the specified position
    /// </summary>
    private void AddSlideToPresentation(PresentationPart presentationPart, SlidePart newSlidePart, int insertPosition)
    {
        var slideIdList = presentationPart.Presentation.SlideIdList ??
            throw new SlideGenerationException("Slide ID list is missing");

        var newSlideId = CreateNewSlideId(presentationPart, newSlidePart, slideIdList);

        if (ShouldInsertAtPosition(insertPosition, slideIdList))
        {
            InsertSlideAtPosition(slideIdList, newSlideId, insertPosition);
        }
        else
        {
            AppendSlideAtEnd(slideIdList, newSlideId);
        }
    }

    /// <summary>
    /// Creates a new slide ID with a unique identifier
    /// </summary>
    private SlideId CreateNewSlideId(PresentationPart presentationPart, SlidePart newSlidePart, SlideIdList slideIdList)
    {
        var maxSlideId = slideIdList.ChildElements.OfType<SlideId>().Max(s => s.Id?.Value) ?? BaseSlideId;

        return new SlideId
        {
            Id = maxSlideId + 1,
            RelationshipId = presentationPart.GetIdOfPart(newSlidePart)
        };
    }

    /// <summary>
    /// Determines if the slide should be inserted at a specific position
    /// </summary>
    private static bool ShouldInsertAtPosition(int insertPosition, SlideIdList slideIdList)
    {
        return insertPosition >= 0 && insertPosition < slideIdList.ChildElements.Count;
    }

    /// <summary>
    /// Inserts the slide at the specified position
    /// </summary>
    private void InsertSlideAtPosition(SlideIdList slideIdList, SlideId newSlideId, int insertPosition)
    {
        var existingSlides = slideIdList.ChildElements.OfType<SlideId>().ToList();

        if (insertPosition < existingSlides.Count)
        {
            slideIdList.InsertBefore(newSlideId, existingSlides[insertPosition]);
            Logger.Debug($"SlideCloner: Inserted slide at position {insertPosition}, new slide ID: {newSlideId.Id?.Value}");
        }
        else
        {
            AppendSlideAtEnd(slideIdList, newSlideId);
        }
    }

    /// <summary>
    /// Appends the slide at the end of the presentation
    /// </summary>
    private void AppendSlideAtEnd(SlideIdList slideIdList, SlideId newSlideId)
    {
        slideIdList.Append(newSlideId);
        Logger.Debug($"SlideCloner: Appended slide at end, new slide ID: {newSlideId.Id?.Value}");
    }

    /// <summary>
    /// Clones slide layout relationship if it exists
    /// </summary>
    private static void CloneSlideLayoutRelationship(SlidePart templateSlidePart, SlidePart newSlidePart)
    {
        if (templateSlidePart.SlideLayoutPart != null)
        {
            newSlidePart.AddPart(templateSlidePart.SlideLayoutPart);
        }
    }

    /// <summary>
    /// Clones all image parts and their relationships from the template slide to the new slide
    /// </summary>
    private static void CloneImageRelationships(SlidePart templateSlidePart, SlidePart newSlidePart)
    {
        try
        {
            // Get all image parts from the template slide
            var imageParts = templateSlidePart.ImageParts.ToList();

            if (imageParts.Count == 0)
            {
                Logger.Debug("SlideCloner: No image parts found in template slide");
                return;
            }

            Logger.Debug($"SlideCloner: Found {imageParts.Count} image parts to clone");

            foreach (var templateImagePart in imageParts)
            {
                // Get the relationship ID from the template
                var relationshipId = templateSlidePart.GetIdOfPart(templateImagePart);

                if (string.IsNullOrEmpty(relationshipId))
                {
                    Logger.Warning("SlideCloner: Could not find relationship ID for image part");
                    continue;
                }

                // Create a new image part in the new slide
                var newImagePart = newSlidePart.AddNewPart<ImagePart>(templateImagePart.ContentType, relationshipId);

                // Copy the image data
                using (var templateStream = templateImagePart.GetStream())
                using (var newStream = newImagePart.GetStream(FileMode.Create))
                {
                    templateStream.CopyTo(newStream);
                }

                Logger.Debug($"SlideCloner: Successfully cloned image part with relationship ID: {relationshipId}");
            }

            Logger.Debug($"SlideCloner: Successfully cloned all {imageParts.Count} image parts");
        }
        catch (Exception ex)
        {
            Logger.Warning($"SlideCloner: Error cloning image relationships: {ex.Message}");
            // Don't throw - continue with slide creation even if image cloning fails
        }
    }
}
