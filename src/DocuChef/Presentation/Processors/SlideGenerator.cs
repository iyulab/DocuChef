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
    }    /// <summary>
         /// Generates slides according to the slide plan
         /// </summary>
         /// <param name="presentationDocument">The presentation document</param>
         /// <param name="slidePlan">The slide plan to use for generation</param>
         /// <param name="slideInfos">The analyzed slide information for auto-generating notes</param>
         /// <param name="data">The data object</param>
         /// <param name="aliasMap">Alias mappings for expression transformation</param>
    public void GenerateSlides(PresentationDocument presentationDocument, SlidePlan slidePlan, List<SlideInfo>? slideInfos = null, object? data = null, Dictionary<string, string>? aliasMap = null)
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
            Logger.Debug($"SlideGenerator: Processing slide instance from template {slideInstance.SourceSlideId} with offset {slideInstance.IndexOffset}");            // Check if this is an original slide at its original position
            if (slideInstance.Position == slideInstance.SourceSlideId)
            {
                Logger.Debug($"SlideGenerator: Keeping original slide {slideInstance.SourceSlideId} at position {slideInstance.Position}");
                originalSlidesToKeep.Add(slideInstance.SourceSlideId);

                // Apply aliases to original slide if needed
                if (aliasMap != null && aliasMap.Count > 0)
                {
                    // Get the original slide part
                    var originalSlideId = sourceSlides[slideInstance.SourceSlideId];
                    if (originalSlideId?.RelationshipId?.Value != null)
                    {
                        var originalSlidePart = (SlidePart)presentationPart.GetPartById(originalSlideId.RelationshipId.Value);
                        if (originalSlidePart != null)
                        {
                            ApplyAliasesToSlide(originalSlidePart, aliasMap);
                        }
                    }
                }
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
            Logger.Debug($"SlideGenerator: Cloning slide from template {slideInstance.SourceSlideId} for additional instance with offset {slideInstance.IndexOffset} at position {insertPosition}");            // Clone the slide for the new instance
            var newSlidePart = _slideCloner.CloneSlideFromTemplate(presentationPart, templateSlidePart, insertPosition);

            if (newSlidePart != null)
            {
                // Apply alias transformations before other operations
                if (aliasMap != null && aliasMap.Count > 0)
                {
                    ApplyAliasesToSlide(newSlidePart, aliasMap);
                }

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
    }    /// <summary>
         /// Applies alias transformations to all text elements in a slide
         /// </summary>    /// <summary>
         /// Applies alias transformations to all text elements in a slide using 2-stage processing:
         /// Stage 1: Process complete expressions at span (text element) level
         /// Stage 2: Process incomplete expressions at paragraph level by combining text elements
         /// </summary>
    private void ApplyAliasesToSlide(SlidePart slidePart, Dictionary<string, string> aliasMap)
    {
        if (slidePart?.Slide == null || aliasMap == null || aliasMap.Count == 0)
            return;

        Logger.Debug($"SlideGenerator: Applying aliases to slide with 2-stage processing, {aliasMap.Count} alias mappings available");

        // Log alias mappings for debugging
        foreach (var alias in aliasMap)
        {
            Logger.Debug($"SlideGenerator: Available alias mapping: '{alias.Key}' -> '{alias.Value}'");
        }

        // Get all paragraphs in the slide
        var paragraphs = slidePart.Slide.Descendants<DocumentFormat.OpenXml.Drawing.Paragraph>().ToList();
        Logger.Debug($"SlideGenerator: Found {paragraphs.Count} paragraphs to process");

        foreach (var paragraph in paragraphs)
        {
            ProcessParagraphWithTwoStages(paragraph, aliasMap);
        }
    }

    /// <summary>
    /// Processes a paragraph using 2-stage alias transformation
    /// </summary>
    private void ProcessParagraphWithTwoStages(DocumentFormat.OpenXml.Drawing.Paragraph paragraph, Dictionary<string, string> aliasMap)
    {
        var textElements = paragraph.Descendants<DocumentFormat.OpenXml.Drawing.Text>().ToList();
        if (textElements.Count == 0)
            return;

        Logger.Debug($"SlideGenerator: Processing paragraph with {textElements.Count} text elements");

        // Stage 1: Process complete expressions at span (text element) level
        bool hasCompleteExpressions = ProcessCompleteExpressionsAtSpanLevel(textElements, aliasMap);

        // Stage 2: Process incomplete expressions at paragraph level if no complete expressions were found
        if (!hasCompleteExpressions)
        {
            ProcessIncompleteExpressionsAtParagraphLevel(textElements, aliasMap);
        }
    }

    /// <summary>
    /// Stage 1: Process complete expressions at individual text element level
    /// Returns true if any complete expressions were found and processed
    /// </summary>
    private bool ProcessCompleteExpressionsAtSpanLevel(List<DocumentFormat.OpenXml.Drawing.Text> textElements, Dictionary<string, string> aliasMap)
    {
        bool foundCompleteExpressions = false;
        var completeExpressionPattern = new Regex(@"\$\{[^}]+\}", RegexOptions.Compiled);

        Logger.Debug($"SlideGenerator: Stage 1 - Processing complete expressions at span level");

        for (int i = 0; i < textElements.Count; i++)
        {
            var text = textElements[i].Text;
            if (string.IsNullOrEmpty(text))
                continue;

            // Check if this text element contains complete expressions
            if (completeExpressionPattern.IsMatch(text))
            {
                Logger.Debug($"SlideGenerator: Stage 1 - Found complete expression in span {i}: '{text}'");

                var transformedText = _expressionUpdater.ApplyAliases(text, aliasMap);
                if (transformedText != text)
                {
                    Logger.Debug($"SlideGenerator: Stage 1 - Applied alias transformation: '{text}' -> '{transformedText}'");
                    textElements[i].Text = transformedText;
                    foundCompleteExpressions = true;
                }
            }
        }

        Logger.Debug($"SlideGenerator: Stage 1 - Found complete expressions: {foundCompleteExpressions}");
        return foundCompleteExpressions;
    }

    /// <summary>
    /// Stage 2: Process incomplete expressions by combining text elements at paragraph level
    /// </summary>
    private void ProcessIncompleteExpressionsAtParagraphLevel(List<DocumentFormat.OpenXml.Drawing.Text> textElements, Dictionary<string, string> aliasMap)
    {
        Logger.Debug($"SlideGenerator: Stage 2 - Processing incomplete expressions at paragraph level");

        // Combine all text elements in the paragraph to form complete expressions
        var combinedText = string.Join("", textElements.Select(t => t.Text));
        Logger.Debug($"SlideGenerator: Stage 2 - Combined text: '{combinedText}'");

        // Check if the combined text contains expressions that need transformation
        var expressionPattern = new Regex(@"\$\{[^}]+\}", RegexOptions.Compiled);
        if (!expressionPattern.IsMatch(combinedText))
        {
            Logger.Debug($"SlideGenerator: Stage 2 - No expressions found in combined text");
            return;
        }

        // Apply alias transformations to the combined text
        var transformedText = _expressionUpdater.ApplyAliases(combinedText, aliasMap);

        if (transformedText != combinedText)
        {
            Logger.Debug($"SlideGenerator: Stage 2 - Applied alias transformation: '{combinedText}' -> '{transformedText}'");

            // Update the first text element with the transformed text and clear others
            // This preserves the paragraph structure while applying the transformation
            textElements[0].Text = transformedText;

            // Clear the remaining text elements in this paragraph
            for (int i = 1; i < textElements.Count; i++)
            {
                textElements[i].Text = "";
            }
        }
        else
        {
            Logger.Debug($"SlideGenerator: Stage 2 - No alias transformation applied");
        }
    }
}
