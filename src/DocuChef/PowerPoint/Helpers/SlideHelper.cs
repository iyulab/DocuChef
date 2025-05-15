namespace DocuChef.PowerPoint.Helpers;

/// <summary>
/// Helper class for slide operations in PowerPoint templates
/// </summary>
internal static class SlideHelper
{
    /// <summary>
    /// Clone a slide with its relationships
    /// </summary>
    public static SlidePart CloneSlide(PresentationPart presentationPart, SlidePart sourceSlidePart)
    {
        var newSlidePart = presentationPart.AddNewPart<SlidePart>();

        // Clone slide content
        using (var sourceReader = new System.IO.StreamReader(sourceSlidePart.GetStream()))
        {
            string slideXml = sourceReader.ReadToEnd();

            using (var writer = new System.IO.StreamWriter(newSlidePart.GetStream(FileMode.Create)))
            {
                writer.Write(slideXml);
            }
        }

        // Clone relationships
        CloneSlideRelationships(sourceSlidePart, newSlidePart);

        // Save the new slide part
        newSlidePart.Slide.Save();

        return newSlidePart;
    }

    /// <summary>
    /// Insert slide in presentation at specified position
    /// </summary>
    public static void InsertSlide(PresentationPart presentationPart, SlidePart slidePart, int position)
    {
        var slideIdList = presentationPart.Presentation.SlideIdList;
        var slideIds = slideIdList.ChildElements.OfType<SlideId>().ToList();

        // Get max slide ID
        uint maxSlideId = slideIds.Any() ? slideIds.Max(id => id.Id.Value) : 0;

        // Create new slide ID
        var newSlideId = new SlideId
        {
            Id = maxSlideId + 1,
            RelationshipId = presentationPart.GetIdOfPart(slidePart)
        };

        // Insert at position
        slideIdList.InsertAt(newSlideId, Math.Min(position, slideIds.Count));

        // Save presentation
        presentationPart.Presentation.Save();
    }

    /// <summary>
    /// Find slide position in presentation
    /// </summary>
    public static int FindSlidePosition(PresentationPart presentationPart, SlidePart slidePart)
    {
        var slideIdList = presentationPart.Presentation.SlideIdList;
        var slideIds = slideIdList.ChildElements.OfType<SlideId>().ToList();
        var relationshipId = presentationPart.GetIdOfPart(slidePart);

        return slideIds.FindIndex(id => id.RelationshipId == relationshipId);
    }

    /// <summary>
    /// Clone slide relationships without circular references
    /// </summary>
    private static void CloneSlideRelationships(SlidePart sourceSlidePart, SlidePart targetSlidePart)
    {
        try
        {
            // Clone slide layout relationship
            if (sourceSlidePart.SlideLayoutPart != null)
            {
                targetSlidePart.CreateRelationshipToPart(sourceSlidePart.SlideLayoutPart);
            }

            // Clone notes slide if exists
            if (sourceSlidePart.NotesSlidePart != null)
            {
                CloneNotesSlidePart(sourceSlidePart.NotesSlidePart, targetSlidePart);
            }

            // Clone image parts
            CloneImageParts(sourceSlidePart, targetSlidePart);

            // Clone external relationships
            foreach (var relationship in sourceSlidePart.ExternalRelationships)
            {
                targetSlidePart.AddExternalRelationship(
                    relationship.RelationshipType,
                    relationship.Uri,
                    relationship.Id);
            }
        }
        catch (Exception ex)
        {
            Logger.Warning($"Error cloning slide relationships: {ex.Message}");
        }
    }

    /// <summary>
    /// Clone notes slide part
    /// </summary>
    private static void CloneNotesSlidePart(NotesSlidePart sourceNotesPart, SlidePart targetSlidePart)
    {
        var targetNotesPart = targetSlidePart.AddNewPart<NotesSlidePart>();

        using (var sourceReader = new System.IO.StreamReader(sourceNotesPart.GetStream()))
        {
            string notesXml = sourceReader.ReadToEnd();

            using (var writer = new System.IO.StreamWriter(targetNotesPart.GetStream(FileMode.Create)))
            {
                writer.Write(notesXml);
            }
        }
    }

    /// <summary>
    /// Clone image parts
    /// </summary>
    private static void CloneImageParts(SlidePart sourceSlidePart, SlidePart targetSlidePart)
    {
        foreach (var idPartPair in sourceSlidePart.Parts)
        {
            if (idPartPair.OpenXmlPart is ImagePart imageSourcePart)
            {
                ImagePart targetImagePart = null;

                // Add image part based on content type
                switch (imageSourcePart.ContentType.ToLowerInvariant())
                {
                    case "image/jpeg":
                    case "image/jpg":
                        targetImagePart = targetSlidePart.AddImagePart(ImagePartType.Jpeg, idPartPair.RelationshipId);
                        break;
                    case "image/png":
                        targetImagePart = targetSlidePart.AddImagePart(ImagePartType.Png, idPartPair.RelationshipId);
                        break;
                    case "image/gif":
                        targetImagePart = targetSlidePart.AddImagePart(ImagePartType.Gif, idPartPair.RelationshipId);
                        break;
                    case "image/bmp":
                        targetImagePart = targetSlidePart.AddImagePart(ImagePartType.Bmp, idPartPair.RelationshipId);
                        break;
                    case "image/tiff":
                        targetImagePart = targetSlidePart.AddImagePart(ImagePartType.Tiff, idPartPair.RelationshipId);
                        break;
                    default:
                        // Default to PNG for unknown types
                        targetImagePart = targetSlidePart.AddImagePart(ImagePartType.Png, idPartPair.RelationshipId);
                        break;
                }

                if (targetImagePart != null)
                {
                    using (var sourceStream = imageSourcePart.GetStream(FileMode.Open, FileAccess.Read))
                    using (var targetStream = targetImagePart.GetStream(FileMode.Create, FileAccess.Write))
                    {
                        sourceStream.CopyTo(targetStream);
                    }
                }
            }
        }
    }
}