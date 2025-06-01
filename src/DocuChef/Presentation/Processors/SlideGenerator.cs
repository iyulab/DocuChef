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

        Logger.Debug($"SlideGenerator: Found {sourceSlides.Count} source slides");        // ==== PHASE 1: Clone all slides first (preserving original expressions) ====
        var slidesToClone = new List<(SlideInstance instance, int insertPosition)>();
        var originalSlidesToRemove = new HashSet<int>();
        var originalSlidesToKeep = new HashSet<int>();
        var generatedSlides = new List<(SlidePart slidePart, SlideInstance instance)>();

        // Collect slides to clone and track originals
        foreach (var slideInstance in slidePlan.SlideInstances)
        {
            Logger.Debug($"SlideGenerator: Processing slide instance from template {slideInstance.SourceSlideId} with offset {slideInstance.IndexOffset}");

            // Check if this is an original slide at its original position
            if (slideInstance.Position == slideInstance.SourceSlideId)
            {
                Logger.Debug($"SlideGenerator: Keeping original slide {slideInstance.SourceSlideId} at position {slideInstance.Position}");
                originalSlidesToKeep.Add(slideInstance.SourceSlideId);

                // Add original slide to processing list for Phase 2
                var originalSlideId = sourceSlides[slideInstance.SourceSlideId];
                if (originalSlideId?.RelationshipId?.Value != null)
                {
                    var originalSlidePart = (SlidePart)presentationPart.GetPartById(originalSlideId.RelationshipId.Value);
                    if (originalSlidePart != null)
                    {
                        generatedSlides.Add((originalSlidePart, slideInstance));
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

            var insertPosition = slideInstance.Position;
            slidesToClone.Add((slideInstance, insertPosition));
        }

        // Sort by insert position to maintain correct order from slide plan
        slidesToClone.Sort((a, b) => a.insertPosition.CompareTo(b.insertPosition));

        // Clone slides without any expression modifications
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

            Logger.Debug($"SlideGenerator: Cloning slide from template {slideInstance.SourceSlideId} at position {insertPosition}");

            // Clone the slide WITHOUT any expression modifications
            var newSlidePart = _slideCloner.CloneSlideFromTemplate(presentationPart, templateSlidePart, insertPosition);

            if (newSlidePart != null)
            {
                // Generate auto notes if slide info is available
                var slideInfo = slideInfos?.FirstOrDefault(s => s.SlideId == slideInstance.SourceSlideId);
                if (slideInfo != null)
                {
                    GenerateAutoNotesIfNeeded(newSlidePart, slideInfo);
                }

                // Add to processing list for Phase 2
                generatedSlides.Add((newSlidePart, slideInstance));
                Logger.Debug($"SlideGenerator: Successfully cloned slide from template {slideInstance.SourceSlideId}");
            }
            else
            {
                Logger.Warning($"SlideGenerator: Failed to clone slide from template {slideInstance.SourceSlideId}");
            }
        }

        // ==== PHASE 2: Apply expression corrections to all slides ====
        Logger.Debug($"SlideGenerator: Starting Phase 2 - Expression correction for {generatedSlides.Count} slides");

        foreach (var (slidePart, slideInstance) in generatedSlides)
        {
            Logger.Debug($"SlideGenerator: Applying expression updates to slide from template {slideInstance.SourceSlideId} with offset {slideInstance.IndexOffset}");

            // Apply alias transformations first
            if (aliasMap != null && aliasMap.Count > 0)
            {
                ApplyAliasesToSlide(slidePart, aliasMap);
            }

            // Apply expression updates with correct context path and index offset
            var contextPathString = slideInstance.ContextPath.Count > 0 ? string.Join(">", slideInstance.ContextPath) : null;
            _expressionUpdater.UpdateExpressionsWithIndexOffset(slidePart, slideInstance.IndexOffset, data, contextPathString);

            Logger.Debug($"SlideGenerator: Applied expression updates to slide from template {slideInstance.SourceSlideId} with context: {contextPathString}, offset: {slideInstance.IndexOffset}");
        }

        // Remove original slides that have been repositioned
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
    }    /// <summary>
         /// Processes a paragraph using 2-stage alias transformation
         /// Enhanced to handle mixed complete and incomplete expressions
         /// </summary>
    private void ProcessParagraphWithTwoStages(DocumentFormat.OpenXml.Drawing.Paragraph paragraph, Dictionary<string, string> aliasMap)
    {
        var textElements = paragraph.Descendants<DocumentFormat.OpenXml.Drawing.Text>().ToList();
        if (textElements.Count == 0)
            return;

        Logger.Debug($"SlideGenerator: Processing paragraph with {textElements.Count} text elements");

        // Stage 1: Process complete expressions at span (text element) level
        var processedElements = new HashSet<int>();
        ProcessCompleteExpressionsAtSpanLevel(textElements, aliasMap, processedElements);

        // Stage 2: Process incomplete expressions at paragraph level for unprocessed elements
        ProcessIncompleteExpressionsAtParagraphLevel(textElements, aliasMap, processedElements);
    }

    /// <summary>
    /// Stage 1: Process complete expressions at individual text element level
    /// Tracks which elements were processed to avoid double-processing
    /// </summary>
    private void ProcessCompleteExpressionsAtSpanLevel(List<DocumentFormat.OpenXml.Drawing.Text> textElements, Dictionary<string, string> aliasMap, HashSet<int> processedElements)
    {
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
                    processedElements.Add(i);
                }
            }
        }

        Logger.Debug($"SlideGenerator: Stage 1 - Processed {processedElements.Count} elements with complete expressions");
    }    /// <summary>
         /// Stage 2: Process incomplete expressions by combining text elements at paragraph level
         /// Uses smart text distribution to preserve formatting where possible
         /// Only processes elements that weren't already handled in Stage 1
         /// </summary>
    private void ProcessIncompleteExpressionsAtParagraphLevel(List<DocumentFormat.OpenXml.Drawing.Text> textElements, Dictionary<string, string> aliasMap, HashSet<int> processedElements)
    {
        Logger.Debug($"SlideGenerator: Stage 2 - Processing incomplete expressions at paragraph level");
        Logger.Debug($"SlideGenerator: Stage 2 - Elements processed in Stage 1: {string.Join(", ", processedElements)}");

        // Check if there are any unprocessed elements that might contain partial expressions
        var unprocessedElements = new List<(int index, DocumentFormat.OpenXml.Drawing.Text element)>();
        for (int i = 0; i < textElements.Count; i++)
        {
            if (!processedElements.Contains(i) && !string.IsNullOrEmpty(textElements[i].Text))
            {
                unprocessedElements.Add((i, textElements[i]));
            }
        }

        if (unprocessedElements.Count == 0)
        {
            Logger.Debug($"SlideGenerator: Stage 2 - No unprocessed elements found");
            return;
        }

        Logger.Debug($"SlideGenerator: Stage 2 - Found {unprocessedElements.Count} unprocessed elements");

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

            // Use smart text distribution to preserve formatting where possible
            DistributeTextSmartly(textElements, combinedText, transformedText);

            // Mark all elements as processed
            for (int i = 0; i < textElements.Count; i++)
            {
                processedElements.Add(i);
            }
        }
        else
        {
            Logger.Debug($"SlideGenerator: Stage 2 - No alias transformation applied");
        }
    }

    /// <summary>
    /// Distributes transformed text across original text elements to preserve formatting
    /// Prioritizes keeping expressions intact within single elements when possible
    /// </summary>
    private void DistributeTextSmartly(List<DocumentFormat.OpenXml.Drawing.Text> textElements, string originalText, string transformedText)
    {
        Logger.Debug($"SlideGenerator: Smart distribution - Original: '{originalText}' -> Transformed: '{transformedText}'");

        // If transformation didn't change the structure significantly, try to preserve original distribution
        if (TryPreserveOriginalDistribution(textElements, originalText, transformedText))
        {
            Logger.Debug($"SlideGenerator: Smart distribution - Preserved original text distribution");
            return;
        }

        // If we have expressions that span multiple elements, use intelligent distribution
        if (TryIntelligentDistribution(textElements, transformedText))
        {
            Logger.Debug($"SlideGenerator: Smart distribution - Used intelligent distribution");
            return;
        }

        // Fallback: Place all text in the first element and clear others
        Logger.Debug($"SlideGenerator: Smart distribution - Using fallback distribution");
        textElements[0].Text = transformedText;
        for (int i = 1; i < textElements.Count; i++)
        {
            textElements[i].Text = "";
        }
    }

    /// <summary>
    /// Attempts to preserve the original text distribution pattern
    /// Returns true if successful, false if transformation changed structure too much
    /// </summary>
    private bool TryPreserveOriginalDistribution(List<DocumentFormat.OpenXml.Drawing.Text> textElements, string originalText, string transformedText)
    {
        // Calculate the length change ratio
        double lengthRatio = originalText.Length > 0 ? (double)transformedText.Length / originalText.Length : 1.0;

        // If the text length changed dramatically (more than 50%), don't try to preserve distribution
        if (lengthRatio < 0.5 || lengthRatio > 2.0)
        {
            Logger.Debug($"SlideGenerator: Length ratio {lengthRatio:F2} too extreme, cannot preserve distribution");
            return false;
        }

        // Try to distribute text proportionally based on original lengths
        var originalLengths = textElements.Select(t => t.Text.Length).ToList();
        var totalOriginalLength = originalLengths.Sum();

        if (totalOriginalLength == 0)
            return false;

        int currentPos = 0;
        bool success = true;

        for (int i = 0; i < textElements.Count; i++)
        {
            if (currentPos >= transformedText.Length)
            {
                textElements[i].Text = "";
                continue;
            }

            // Calculate proportional length for this element
            double proportion = (double)originalLengths[i] / totalOriginalLength;
            int targetLength = Math.Max(1, (int)(transformedText.Length * proportion));

            // Don't break expressions - find a safe break point
            int actualLength = FindSafeBreakPoint(transformedText, currentPos, Math.Min(currentPos + targetLength, transformedText.Length));

            if (actualLength <= currentPos && i < textElements.Count - 1)
            {
                // Cannot find a safe break point, fallback
                success = false;
                break;
            }

            // For the last element, take all remaining text
            if (i == textElements.Count - 1)
                actualLength = transformedText.Length;

            textElements[i].Text = transformedText.Substring(currentPos, actualLength - currentPos);
            currentPos = actualLength;
        }

        return success && currentPos == transformedText.Length;
    }

    /// <summary>
    /// Attempts intelligent distribution that keeps expressions intact
    /// </summary>
    private bool TryIntelligentDistribution(List<DocumentFormat.OpenXml.Drawing.Text> textElements, string transformedText)
    {
        // Find all expressions in the transformed text
        var expressionPattern = new Regex(@"\$\{[^}]+\}", RegexOptions.Compiled);
        var matches = expressionPattern.Matches(transformedText).Cast<Match>().ToList();

        if (matches.Count == 0)
        {
            // No expressions, distribute text evenly
            return DistributeTextEvenly(textElements, transformedText);
        }

        // If we have more expressions than text elements, fallback
        if (matches.Count > textElements.Count)
        {
            Logger.Debug($"SlideGenerator: {matches.Count} expressions > {textElements.Count} text elements, using fallback");
            return false;
        }

        // Try to place each expression in its own text element
        int currentPos = 0;
        int elementIndex = 0;

        foreach (var match in matches)
        {
            // Place text before the expression in current element if there's space
            if (match.Index > currentPos && elementIndex < textElements.Count)
            {
                string beforeText = transformedText.Substring(currentPos, match.Index - currentPos);
                if (elementIndex > 0 && string.IsNullOrEmpty(textElements[elementIndex].Text))
                {
                    textElements[elementIndex].Text = beforeText;
                    elementIndex++;
                }
                else if (elementIndex == 0)
                {
                    textElements[elementIndex].Text = beforeText;
                    elementIndex++;
                }
            }

            // Place the expression in the next available element
            if (elementIndex < textElements.Count)
            {
                textElements[elementIndex].Text = match.Value;
                elementIndex++;
                currentPos = match.Index + match.Length;
            }
            else
            {
                // No more elements available
                return false;
            }
        }

        // Place any remaining text in the last available element
        if (currentPos < transformedText.Length && elementIndex < textElements.Count)
        {
            textElements[elementIndex].Text = transformedText.Substring(currentPos);
            elementIndex++;
        }

        // Clear any remaining elements
        for (int i = elementIndex; i < textElements.Count; i++)
        {
            textElements[i].Text = "";
        }

        return true;
    }

    /// <summary>
    /// Distributes text evenly across all text elements
    /// </summary>
    private bool DistributeTextEvenly(List<DocumentFormat.OpenXml.Drawing.Text> textElements, string text)
    {
        if (textElements.Count == 0 || string.IsNullOrEmpty(text))
            return false;

        int charsPerElement = text.Length / textElements.Count;
        int remainder = text.Length % textElements.Count;
        int currentPos = 0;

        for (int i = 0; i < textElements.Count; i++)
        {
            int length = charsPerElement + (i < remainder ? 1 : 0);
            if (currentPos + length > text.Length)
                length = text.Length - currentPos;

            if (length > 0)
            {
                textElements[i].Text = text.Substring(currentPos, length);
                currentPos += length;
            }
            else
            {
                textElements[i].Text = "";
            }
        }

        return true;
    }

    /// <summary>
    /// Finds a safe break point that doesn't break expressions
    /// </summary>
    private int FindSafeBreakPoint(string text, int start, int preferredEnd)
    {
        if (preferredEnd >= text.Length)
            return text.Length;

        // Don't break inside expressions
        bool insideExpression = false;
        int braceCount = 0;

        for (int i = start; i < preferredEnd; i++)
        {
            if (text[i] == '$' && i + 1 < text.Length && text[i + 1] == '{')
            {
                insideExpression = true;
                braceCount = 1;
                i++; // skip the '{'
                continue;
            }

            if (insideExpression)
            {
                if (text[i] == '{')
                    braceCount++;
                else if (text[i] == '}')
                {
                    braceCount--;
                    if (braceCount == 0)
                    {
                        insideExpression = false;
                        // This is a safe break point after the expression
                        if (i + 1 <= preferredEnd)
                            return i + 1;
                    }
                }
            }
        }

        // If we're inside an expression, we can't break here
        if (insideExpression)
        {
            // Try to find the start of the expression and break before it
            for (int i = preferredEnd - 1; i >= start; i--)
            {
                if (text[i] == '$' && i + 1 < text.Length && text[i + 1] == '{')
                    return i;
            }
            // If we can't find a safe break, return the start
            return start;
        }

        return preferredEnd;
    }
}
