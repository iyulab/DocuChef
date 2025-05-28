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
    private readonly TemplateAnalyzer _templateAnalyzer = new TemplateAnalyzer();

    /// <summary>
    /// Generates slides according to the slide plan
    /// </summary>
    /// <param name="presentationDocument">The presentation document</param>
    /// <param name="slidePlan">The slide plan to use for generation</param>
    /// <param name="slideInfos">The analyzed slide information for auto-generating notes</param>
    public void GenerateSlides(PresentationDocument presentationDocument, SlidePlan slidePlan, List<SlideInfo>? slideInfos = null)
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

        Logger.Debug($"SlideGenerator: Found {sourceSlides.Count} source slides");        // Generate slides based on the plan
        // Only clone slides for instances with IndexOffset > 0 (original slides are kept as-is)
        foreach (var slideInstance in slidePlan.SlideInstances)
        {
            Logger.Debug($"SlideGenerator: Processing slide instance from template {slideInstance.SourceSlideId} with offset {slideInstance.IndexOffset}");

            // Skip original slides (IndexOffset = 0) - they stay as-is
            if (slideInstance.IndexOffset == 0)
            {
                Logger.Debug($"SlideGenerator: Skipping original slide {slideInstance.SourceSlideId} (offset=0)");
                continue;
            }

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

            Logger.Debug($"SlideGenerator: Cloning slide from template {slideInstance.SourceSlideId} for additional instance with offset {slideInstance.IndexOffset}");

            // Clone the slide for the new instance
            var newSlidePart = CloneSlideFromTemplate(presentationPart, templateSlidePart);

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
                UpdateExpressionsWithIndexOffset(newSlidePart, slideInstance.IndexOffset);

                Logger.Debug($"SlideGenerator: Successfully generated additional slide from template {slideInstance.SourceSlideId}");
            }
            else
            {
                Logger.Warning($"SlideGenerator: Failed to clone slide from template {slideInstance.SourceSlideId}");
            }
        }

        Logger.Debug("SlideGenerator: Slide generation completed");
    }

    /// <summary>
    /// Updates expressions in the slide with the given index offset
    /// Note: This only adjusts array indices, actual data binding happens in DataBinder.cs
    /// </summary>
    private void UpdateExpressionsWithIndexOffset(SlidePart slidePart, int indexOffset)
    {
        if (slidePart?.Slide == null || indexOffset <= 0)
            return;

        Logger.Debug($"SlideGenerator: Updating expressions with index offset {indexOffset}");

        try
        {
            // Find all text elements in the slide
            var textElements = slidePart.Slide.Descendants<DocumentFormat.OpenXml.Drawing.Text>().ToList();

            foreach (var textElement in textElements)
            {
                if (string.IsNullOrEmpty(textElement.Text))
                    continue;

                // Only adjust array indices in expressions, don't evaluate them
                var updatedText = AdjustArrayIndicesInText(textElement.Text, indexOffset);
                if (updatedText != textElement.Text)
                {
                    Logger.Debug($"SlideGenerator: Updated expression from '{textElement.Text}' to '{updatedText}'");
                    textElement.Text = updatedText;
                }
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
    }

    /// <summary>
    /// Clones a slide from the template
    /// </summary>
    private SlidePart CloneSlideFromTemplate(PresentationPart presentationPart, SlidePart templateSlidePart)
    {
        // Create a new slide part
        var newSlidePart = presentationPart.AddNewPart<SlidePart>();

        // Clone the slide content
        newSlidePart.Slide = (Slide)templateSlidePart.Slide.CloneNode(true);

        // Add the new slide to the slide ID list
        var slideIdList = presentationPart.Presentation.SlideIdList;
        var maxSlideId = slideIdList!.ChildElements.OfType<SlideId>().Max(s => s.Id?.Value) ?? 255;
        var newSlideId = new SlideId { Id = maxSlideId + 1, RelationshipId = presentationPart.GetIdOfPart(newSlidePart) };
        slideIdList.Append(newSlideId);

        // Clone slide layout relationship if it exists
        if (templateSlidePart.SlideLayoutPart != null)
        {
            newSlidePart.AddPart(templateSlidePart.SlideLayoutPart);
        }

        Logger.Debug($"SlideGenerator: Cloned slide, new slide ID: {newSlideId.Id?.Value}");

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
            Logger.Warning($"SlideGenerator: Error generating auto notes: {ex.Message}");
        }
    }
}
