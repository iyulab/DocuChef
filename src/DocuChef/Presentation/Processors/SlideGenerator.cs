using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using DocuChef.Presentation.Exceptions;
using DocuChef.Presentation.Models;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Drawing;
using DocuChef.Logging;

namespace DocuChef.Presentation.Processors;

/// <summary>
/// Generates slides based on the slide plan
/// Note: Data binding is handled exclusively in DataBinder.cs
/// </summary>
public class SlideGenerator
{
    // Note: All data binding related fields removed - binding handled exclusively in DataBinder.cs
    private readonly TemplateAnalyzer _templateAnalyzer = new TemplateAnalyzer();    /// <summary>
                                                                                     /// Generates slides according to the slide plan
                                                                                     /// </summary>
                                                                                     /// <param name="presentationDocument">The presentation document</param>
                                                                                     /// <param name="slidePlan">The slide plan to use for generation</param>
                                                                                     /// <param name="slideInfos">The analyzed slide information for auto-generating notes</param>
    public void GenerateSlides(PresentationDocument presentationDocument, SlidePlan slidePlan, List<SlideInfo>? slideInfos = null, object? data = null)
    {
        Logger.Debug($"SlideGenerator: Starting generation with {slidePlan?.SlideInstances?.Count ?? 0} slide instances");

        if (presentationDocument == null)
            throw new ArgumentNullException(nameof(presentationDocument));
        if (slidePlan == null || slidePlan.SlideInstances.Count == 0)
        {
            Logger.Debug("SlideGenerator: No slide instances to generate");
            return;
        }

        // Validate presentation
        ValidatePresentation(presentationDocument);

        // Get the presentation part
        var presentationPart = presentationDocument.PresentationPart;
        if (presentationPart == null)
            throw new SlideGenerationException("Presentation part is missing");

        // Get the slide ID list
        var slideIdList = presentationPart.Presentation.SlideIdList;
        if (slideIdList == null)
            throw new SlideGenerationException("Slide ID list is missing");

        var sourceSlides = slideIdList.ChildElements.OfType<SlideId>().ToList();

        Logger.Debug($"SlideGenerator: Found {sourceSlides.Count} source slides");        // Collect all slides to clone based on their planned position
        var slidesToClone = new List<(SlideInstance instance, int insertPosition)>();
        
        // Track which original template slides are being repositioned and should be removed
        var originalSlidesToRemove = new HashSet<int>();
        var originalSlidesToKeep = new HashSet<int>();

        // Process slide instances in their planned order
        foreach (var slideInstance in slidePlan.SlideInstances)
        {            Logger.Debug($"SlideGenerator: Processing slide instance from template {slideInstance.SourceSlideId} with offset {slideInstance.IndexOffset}");

            // Check if this is an original slide at its original position
            if (slideInstance.Position == slideInstance.SourceSlideId)
            {
                Logger.Debug($"SlideGenerator: Keeping original slide {slideInstance.SourceSlideId} at position {slideInstance.Position}");
                originalSlidesToKeep.Add(slideInstance.SourceSlideId);
                continue;
            }
            else
            {
                // This slide is being repositioned, so mark the original for removal
                if (!originalSlidesToKeep.Contains(slideInstance.SourceSlideId))
                {
                    originalSlidesToRemove.Add(slideInstance.SourceSlideId);
                }
            }

            // Use the planned position from SlidePlan to maintain correct order
            // Position is 0-based, so we use it directly as insert position
            var insertPosition = slideInstance.Position;

            slidesToClone.Add((slideInstance, insertPosition));
        }

        // Sort by insert position to maintain correct order from slide plan
        slidesToClone.Sort((a, b) => a.insertPosition.CompareTo(b.insertPosition));

        // Clone slides in the calculated order
        foreach (var (slideInstance, insertPosition) in slidesToClone)
        {
            // Find the template slide by index (SourceSlideId is 0-based index)
            if (slideInstance.SourceSlideId < 0 || slideInstance.SourceSlideId >= sourceSlides.Count)
            {
                Logger.Warning($"SlideGenerator: Template slide index {slideInstance.SourceSlideId} is out of range (0-{sourceSlides.Count - 1}), skipping");
                continue;
            }

            var templateSlideId = sourceSlides[slideInstance.SourceSlideId];
            if (templateSlideId?.RelationshipId?.Value == null)
            {
                Logger.Warning($"SlideGenerator: Template slide at index {slideInstance.SourceSlideId} has no relationship ID, skipping");
                continue;
            }

            // Get the template slide part
            var templateSlidePart = (SlidePart)presentationPart.GetPartById(templateSlideId.RelationshipId.Value);
            if (templateSlidePart?.Slide == null)
            {
                Logger.Warning($"SlideGenerator: Template slide part for slide {slideInstance.SourceSlideId} is invalid, skipping");
                continue;
            }

            Logger.Debug($"SlideGenerator: Cloning slide from template {slideInstance.SourceSlideId} for additional instance with offset {slideInstance.IndexOffset} at position {insertPosition}");

            // Clone the slide for the new instance
            var newSlidePart = CloneSlideFromTemplate(presentationPart, templateSlidePart, insertPosition);

            if (newSlidePart != null)
            {
                // Generate auto notes if slide info is available
                var slideInfo = slideInfos?.FirstOrDefault(s => s.SlideId == slideInstance.SourceSlideId);
                if (slideInfo != null)
                {
                    GenerateAutoNotesIfNeeded(newSlidePart, slideInfo);
                }

                // Update expressions with index offset - but don't bind data here
                // Data binding will be handled later in DataBinder.cs
                UpdateExpressionsWithIndexOffset(newSlidePart, slideInstance.IndexOffset, data);

                Logger.Debug($"SlideGenerator: Successfully generated additional slide from template {slideInstance.SourceSlideId}");
            }
            else
            {
                Logger.Warning($"SlideGenerator: Failed to clone slide from template {slideInstance.SourceSlideId}");
            }        }

        // Remove original slides that have been repositioned
        RemoveOriginalSlides(presentationPart, sourceSlides, originalSlidesToRemove);

        Logger.Debug("SlideGenerator: Slide generation completed");
    }    /// <summary>
    /// Updates expressions in the slide with the given index offset and hides elements that exceed data bounds
    /// Note: This only adjusts array indices, actual data binding happens in DataBinder.cs
    /// </summary>
    private void UpdateExpressionsWithIndexOffset(SlidePart slidePart, int indexOffset, object? data)    {
        if (slidePart?.Slide == null || indexOffset <= 0)
            return;

        Logger.Debug($"SlideGenerator: Updating expressions with index offset {indexOffset}");

        try
        {
            // Find all text elements in the slide
            var textElements = slidePart.Slide.Descendants<DocumentFormat.OpenXml.Drawing.Text>().ToList();
            var elementsToHide = new List<OpenXmlElement>();

            foreach (var textElement in textElements)
            {
                if (string.IsNullOrEmpty(textElement.Text))
                    continue;

                // Check if this element contains expressions that exceed data bounds
                if (ShouldHideElement(textElement.Text, indexOffset, data))
                {
                    // Find the parent shape to hide
                    var parentShape = FindParentShape(textElement);
                    if (parentShape != null && !elementsToHide.Contains(parentShape))
                    {
                        elementsToHide.Add(parentShape);
                        Logger.Debug($"SlideGenerator: Marking shape for hiding due to data overflow in expression: '{textElement.Text}'");
                    }
                    continue;
                }

                // Only adjust array indices in expressions, don't evaluate them
                var updatedText = AdjustArrayIndicesInText(textElement.Text, indexOffset);
                if (updatedText != textElement.Text)
                {
                    Logger.Debug($"SlideGenerator: Updated expression from '{textElement.Text}' to '{updatedText}'");
                    textElement.Text = updatedText;
                }
            }

            // Hide elements that contain data overflow
            foreach (var element in elementsToHide)
            {
                HideElement(element);
            }
        }
        catch (Exception ex)
        {
            Logger.Warning($"SlideGenerator: Error updating expressions with index offset: {ex.Message}");
        }
    }

    /// <summary>
    /// Adjusts array indices in text expressions
    /// Example: "${Items[0].Name}" becomes "${Items[2].Name}" with offset 2
    /// </summary>
    private string AdjustArrayIndicesInText(string text, int indexOffset)
    {
        if (string.IsNullOrEmpty(text) || indexOffset <= 0)
            return text;

        // Pattern to match array indices in expressions like ${Items[0].Name} or Items[1].Property
        var arrayIndexPattern = new Regex(@"(\w+)\[(\d+)\]", RegexOptions.Compiled);

        return arrayIndexPattern.Replace(text, match =>
        {
            var arrayName = match.Groups[1].Value;
            var currentIndex = int.Parse(match.Groups[2].Value);
            var newIndex = currentIndex + indexOffset;
            return $"{arrayName}[{newIndex}]";
        });
    }

    /// <summary>
    /// Validates that the presentation document is properly structured
    /// </summary>
    private void ValidatePresentation(PresentationDocument presentationDocument)
    {
        if (presentationDocument.PresentationPart == null)
            throw new SlideGenerationException("Presentation part is missing.");

        if (presentationDocument.PresentationPart.Presentation == null)
            throw new SlideGenerationException("Presentation object is missing.");

        if (presentationDocument.PresentationPart.Presentation.SlideIdList == null)
            throw new SlideGenerationException("Slide ID list is missing.");
    }

    /// <summary>
    /// Validates that slide generation is possible
    /// </summary>
    public void ValidateSlideGeneration(PresentationDocument presentationDocument, int sourceSlideId)
    {
        ValidatePresentation(presentationDocument);

        var presentationPart = presentationDocument.PresentationPart;
        var slideIdList = presentationPart!.Presentation.SlideIdList;

        var sourceSlides = slideIdList!.ChildElements.OfType<SlideId>().ToList();
        if (!sourceSlides.Any(s => s.Id?.Value == sourceSlideId))
            throw new SlideGenerationException($"Source slide with ID {sourceSlideId} does not exist.");
    }    /// <summary>
         /// Clones a slide from the template
         /// </summary>
    private SlidePart CloneSlideFromTemplate(PresentationPart presentationPart, SlidePart templateSlidePart, int insertPosition = -1)
    {
        // Create a new slide part
        var newSlidePart = presentationPart.AddNewPart<SlidePart>();

        // Clone the slide content
        newSlidePart.Slide = (Slide)templateSlidePart.Slide.CloneNode(true);

        // Add the new slide to the slide ID list at the specified position
        var slideIdList = presentationPart.Presentation.SlideIdList;
        var maxSlideId = slideIdList!.ChildElements.OfType<SlideId>().Max(s => s.Id?.Value) ?? 255;
        var newSlideId = new SlideId { Id = maxSlideId + 1, RelationshipId = presentationPart.GetIdOfPart(newSlidePart) };

        if (insertPosition >= 0 && insertPosition < slideIdList.ChildElements.Count)
        {
            // Insert at the specified position
            var existingSlides = slideIdList.ChildElements.OfType<SlideId>().ToList();
            if (insertPosition < existingSlides.Count)
            {
                slideIdList.InsertBefore(newSlideId, existingSlides[insertPosition]);
                Logger.Debug($"SlideGenerator: Inserted slide at position {insertPosition}, new slide ID: {newSlideId.Id?.Value}");
            }
            else
            {
                slideIdList.Append(newSlideId);
                Logger.Debug($"SlideGenerator: Appended slide at end, new slide ID: {newSlideId.Id?.Value}");
            }
        }
        else
        {
            // Append at the end if no valid position specified
            slideIdList.Append(newSlideId);
            Logger.Debug($"SlideGenerator: Cloned slide, new slide ID: {newSlideId.Id?.Value}");
        }

        // Clone slide layout relationship if it exists
        if (templateSlidePart.SlideLayoutPart != null)
        {
            newSlidePart.AddPart(templateSlidePart.SlideLayoutPart);
        }

        return newSlidePart;
    }

    /// <summary>
    /// Generates automatic notes for the slide if needed
    /// </summary>
    private void GenerateAutoNotesIfNeeded(SlidePart slidePart, SlideInfo slideInfo)
    {
        if (slidePart?.Slide == null || slideInfo == null)
            return;

        try
        {
            // Generate notes content based on slide info
            var autoNotesContent = _templateAnalyzer.GenerateSlideNotes(slideInfo, slidePart.Slide);

            if (!string.IsNullOrEmpty(autoNotesContent))
            {
                // Add notes to the slide if they don't already exist
                if (slidePart.NotesSlidePart == null)
                {
                    var notesSlidePart = slidePart.AddNewPart<NotesSlidePart>(); var notesSlide = new NotesSlide(
                        new CommonSlideData(
                            new ShapeTree(
                                new DocumentFormat.OpenXml.Presentation.NonVisualGroupShapeProperties(
                                    new DocumentFormat.OpenXml.Presentation.NonVisualDrawingProperties { Id = 1, Name = "" },
                                    new DocumentFormat.OpenXml.Presentation.NonVisualGroupShapeDrawingProperties(),
                                    new ApplicationNonVisualDrawingProperties()),
                                new GroupShapeProperties(new TransformGroup()))),
                        new ColorMapOverride(new MasterColorMapping()));

                    notesSlidePart.NotesSlide = notesSlide;
                }

                Logger.Debug($"SlideGenerator: Generated auto notes for slide {slideInfo.SlideId}");
            }
        }
        catch (Exception ex)
        {
            Logger.Warning($"SlideGenerator: Error generating auto notes: {ex.Message}");        }
    }

    /// <summary>
    /// Removes original template slides that have been repositioned
    /// </summary>
    private static void RemoveOriginalSlides(PresentationPart presentationPart, List<SlideId> sourceSlides, HashSet<int> slidesToRemove)
    {
        if (!slidesToRemove.Any()) return;

        var slideIdList = presentationPart.Presentation.SlideIdList;
        if (slideIdList == null) return;

        Logger.Debug($"SlideGenerator: Removing {slidesToRemove.Count} original slides that were repositioned");

        // Remove slides in reverse order to maintain indices
        foreach (var slideIndex in slidesToRemove.OrderByDescending(x => x))
        {
            if (slideIndex >= 0 && slideIndex < sourceSlides.Count)
            {
                var slideToRemove = sourceSlides[slideIndex];
                Logger.Debug($"SlideGenerator: Removing original slide {slideIndex} with RelationshipId {slideToRemove.RelationshipId?.Value}");
                
                // Remove the slide from the presentation
                slideToRemove.Remove();
                
                // Also remove the corresponding slide part
                if (slideToRemove.RelationshipId?.Value != null)
                {
                    try
                    {
                        var slidePart = (SlidePart)presentationPart.GetPartById(slideToRemove.RelationshipId.Value);
                        presentationPart.DeletePart(slidePart);
                        Logger.Debug($"SlideGenerator: Deleted slide part for original slide {slideIndex}");
                    }
                    catch (Exception ex)
                    {
                        Logger.Warning($"SlideGenerator: Failed to delete slide part for slide {slideIndex}: {ex.Message}");
                    }
                }
            }
        }
    }

    /// <summary>
    /// Checks if an element should be hidden due to data index overflow
    /// </summary>
    private bool ShouldHideElement(string text, int indexOffset, object? data)
    {
        if (string.IsNullOrEmpty(text) || data == null)
            return false;

        // Pattern to match array indices in expressions like ${Items[0].Name}
        var arrayIndexPattern = new Regex(@"\$\{(\w+)\[(\d+)\]", RegexOptions.Compiled);
        var matches = arrayIndexPattern.Matches(text);

        foreach (Match match in matches)
        {
            var arrayName = match.Groups[1].Value;
            var currentIndex = int.Parse(match.Groups[2].Value);
            var finalIndex = currentIndex + indexOffset;

            if (!IsIndexValid(arrayName, finalIndex, data))
            {
                return true; // Should hide this element
            }
        }

        return false;
    }

    /// <summary>
    /// Checks if the specified array index is valid for the given data
    /// </summary>
    private bool IsIndexValid(string arrayName, int index, object data)
    {
        try
        {
            if (data is Dictionary<string, object> dict && dict.TryGetValue(arrayName, out var arrayValue))
            {
                if (arrayValue is System.Collections.IList list)
                {
                    return index >= 0 && index < list.Count;
                }
                else if (arrayValue is System.Collections.IEnumerable enumerable)
                {
                    var count = enumerable.Cast<object>().Count();
                    return index >= 0 && index < count;
                }
            }
            
            // Use reflection to check properties
            var property = data.GetType().GetProperty(arrayName);
            if (property != null)
            {
                var value = property.GetValue(data);
                if (value is System.Collections.IList list)
                {
                    return index >= 0 && index < list.Count;
                }
                else if (value is System.Collections.IEnumerable enumerable)
                {
                    var count = enumerable.Cast<object>().Count();
                    return index >= 0 && index < count;
                }
            }
        }
        catch (Exception ex)
        {
            Logger.Warning($"SlideGenerator: Error checking array bounds for {arrayName}[{index}]: {ex.Message}");
        }

        return true; // Default to not hiding if we can't determine bounds
    }

    /// <summary>
    /// Finds the parent shape element for a text element
    /// </summary>
    private OpenXmlElement? FindParentShape(OpenXmlElement element)
    {
        var current = element.Parent;
        while (current != null)
        {
            if (current is DocumentFormat.OpenXml.Presentation.Shape ||
                current is DocumentFormat.OpenXml.Presentation.Picture ||
                current is DocumentFormat.OpenXml.Presentation.GraphicFrame)
            {
                return current;
            }
            current = current.Parent;
        }
        return null;
    }

    /// <summary>
    /// Hides an element by setting its visibility to hidden
    /// </summary>
    private void HideElement(OpenXmlElement element)
    {
        try
        {
            // For shapes, we can set them to be hidden by making them very small or transparent
            if (element is DocumentFormat.OpenXml.Presentation.Shape shape)
            {
                // Find the shape properties
                var spPr = shape.ShapeProperties;
                if (spPr != null)
                {
                    // Set the shape to be invisible by setting width and height to 0
                    var transform = spPr.Transform2D;
                    if (transform != null)
                    {
                        var extents = transform.Extents;
                        if (extents != null)
                        {
                            extents.Cx = 0; // Width = 0
                            extents.Cy = 0; // Height = 0
                            Logger.Debug($"SlideGenerator: Hidden shape by setting dimensions to 0");
                        }
                    }
                }
            }
            else if (element is DocumentFormat.OpenXml.Presentation.Picture picture)
            {
                // Similar logic for pictures
                var spPr = picture.ShapeProperties;
                if (spPr != null)
                {
                    var transform = spPr.Transform2D;
                    if (transform != null)
                    {
                        var extents = transform.Extents;
                        if (extents != null)
                        {
                            extents.Cx = 0;
                            extents.Cy = 0;
                            Logger.Debug($"SlideGenerator: Hidden picture by setting dimensions to 0");
                        }
                    }
                }
            }
        }
        catch (Exception ex)
        {
            Logger.Warning($"SlideGenerator: Error hiding element: {ex.Message}");
        }
    }
}
