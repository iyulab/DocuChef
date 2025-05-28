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
/// Refactored to improve code organization and maintainability
/// </summary>
public class SlideGenerator
{
    private readonly TemplateAnalyzer _templateAnalyzer;
    private readonly SlideCloner _slideCloner;
    private readonly ExpressionUpdater _expressionUpdater;
    private readonly SlideRemover _slideRemover;

    public SlideGenerator()
    {
        _templateAnalyzer = new TemplateAnalyzer();
        _slideCloner = new SlideCloner();
        _expressionUpdater = new ExpressionUpdater();
        _slideRemover = new SlideRemover();
    }/// <summary>
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
        {
            Logger.Debug($"SlideGenerator: Processing slide instance from template {slideInstance.SourceSlideId} with offset {slideInstance.IndexOffset}");

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
            var newSlidePart = _slideCloner.CloneSlideFromTemplate(presentationPart, templateSlidePart, insertPosition);

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
                _expressionUpdater.UpdateExpressionsWithIndexOffset(newSlidePart, slideInstance.IndexOffset, data);

                Logger.Debug($"SlideGenerator: Successfully generated additional slide from template {slideInstance.SourceSlideId}");
            }
            else
            {
                Logger.Warning($"SlideGenerator: Failed to clone slide from template {slideInstance.SourceSlideId}");
            }
        }        // Remove original slides that have been repositioned
        _slideRemover.RemoveOriginalSlides(presentationPart, sourceSlides, originalSlidesToRemove); Logger.Debug("SlideGenerator: Slide generation completed");
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
