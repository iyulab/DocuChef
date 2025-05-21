using DocuChef.Presentation.Models;

namespace DocuChef.Presentation.Core;

public static partial class SlideManager
{
    /// <summary>
    /// Clones the notes slide part and adjusts expressions (only for directive notes)
    /// </summary>
    private static void CloneNotesPartWithContext(SlidePart sourceSlidePart, SlidePart newSlidePart, SlideContext context)
    {
        // Only clone notes if they contain directives
        string noteText = GetSlideNoteText(sourceSlidePart);
        if (!IsDirective(noteText))
        {
            return;
        }

        NotesSlidePart sourceNotesPart = sourceSlidePart.NotesSlidePart;
        NotesSlidePart newNotesPart = newSlidePart.AddNewPart<NotesSlidePart>();

        // Clone notes content
        newNotesPart.NotesSlide = (P.NotesSlide)sourceNotesPart.NotesSlide.CloneNode(true);

        // Adjust expressions in notes text elements
        if (context != null)
        {
            try
            {
                var noteTextElements = newNotesPart.NotesSlide.Descendants<D.Text>().ToList();
                foreach (var textElement in noteTextElements)
                {
                    string originalText = textElement.Text;
                    if (string.IsNullOrEmpty(originalText) || !originalText.Contains("${") || !originalText.Contains('['))
                        continue;

                    string adjustedText = ExpressionAdjuster.AdjustExpressionIndices(originalText, context);
                    if (adjustedText != originalText)
                    {
                        Logger.Debug($"Adjusted notes text expression: '{originalText}' -> '{adjustedText}'");
                        textElement.Text = adjustedText;
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Error($"Error adjusting notes expressions: {ex.Message}", ex);
            }
        }
    }

    /// <summary>
    /// Clones slide relationships (images, charts)
    /// </summary>
    private static void CloneSlideRelationships(SlidePart sourceSlidePart, SlidePart newSlidePart)
    {
        // Handle common relationships like images
        foreach (var idPartPair in sourceSlidePart.Parts)
        {
            OpenXmlPart part = idPartPair.OpenXmlPart;

            // Skip layout and notes parts (already handled)
            if (part is SlideLayoutPart || part is NotesSlidePart)
                continue;

            // Handle image parts
            if (part is ImagePart)
            {
                var imagePart = part as ImagePart;
                var newImagePart = newSlidePart.AddImagePart(imagePart.ContentType);

                using (Stream sourceStream = imagePart.GetStream(FileMode.Open, FileAccess.Read))
                using (Stream targetStream = newImagePart.GetStream(FileMode.Create, FileAccess.Write))
                {
                    sourceStream.CopyTo(targetStream);
                }

                // Update references in the new slide
                string oldId = sourceSlidePart.GetIdOfPart(imagePart);
                string newId = newSlidePart.GetIdOfPart(newImagePart);
                UpdateImageReferences(newSlidePart.Slide, oldId, newId);
            }
        }
    }

    /// <summary>
    /// Clones a slide and adjusts expression indices based on slide context
    /// </summary>
    public static SlidePart CloneSlideWithContext(PresentationPart presentationPart, SlidePart sourceSlidePart, SlideContext context)
    {
        if (presentationPart == null || sourceSlidePart == null)
            throw new ArgumentNullException("Source or presentation part is null");

        // Add a new slide part
        SlidePart newSlidePart = presentationPart.AddNewPart<SlidePart>();

        // Clone slide content
        newSlidePart.Slide = (P.Slide)sourceSlidePart.Slide.CloneNode(true);

        // Adjust expressions in text elements
        AdjustSlideExpressions(newSlidePart, context);

        // Clone slide layout relationship
        if (sourceSlidePart.SlideLayoutPart != null)
        {
            newSlidePart.AddPart(sourceSlidePart.SlideLayoutPart);
        }

        // Clone other important relationships (images, charts, etc.)
        CloneSlideRelationships(sourceSlidePart, newSlidePart);

        // Don't automatically clone notes - only if there are actual directives
        if (sourceSlidePart.NotesSlidePart != null)
        {
            string noteText = GetSlideNoteText(sourceSlidePart);
            if (IsDirective(noteText))
            {
                CloneNotesPartWithContext(sourceSlidePart, newSlidePart, context);
            }
        }

        return newSlidePart;
    }

    /// <summary>
    /// Adjusts expression indices in all paragraphs of a slide
    /// </summary>
    public static void AdjustSlideExpressions(SlidePart slidePart, SlideContext context)
    {
        if (slidePart == null || context == null || context.Offset == 0)
            return;

        Logger.Debug($"Adjusting expressions in slide for context: {context.GetContextDescription()}");

        try
        {
            // Process expressions at the Run level first to preserve formatting
            AdjustRunLevelExpressions(slidePart, context);

            // Then process any remaining expressions at the paragraph level
            AdjustParagraphLevelExpressions(slidePart, context);
        }
        catch (Exception ex)
        {
            Logger.Error($"Error adjusting slide expressions: {ex.Message}", ex);
        }
    }

    /// <summary>
    /// Adjusts expressions at the Run level to preserve formatting
    /// </summary>
    private static void AdjustRunLevelExpressions(SlidePart slidePart, SlideContext context)
    {
        Logger.Debug("Adjusting expressions at Run level to preserve formatting");

        // Get all text runs in the slide
        var runs = slidePart.Slide.Descendants<D.Run>().ToList();
        Logger.Debug($"Found {runs.Count} runs to examine");

        foreach (var run in runs)
        {
            // Get all text elements in this run
            var textElements = run.Descendants<D.Text>().ToList();
            if (textElements.Count == 0)
                continue;

            // Process each text element in this run
            foreach (var textElement in textElements)
            {
                string originalText = textElement.Text;
                if (string.IsNullOrEmpty(originalText) || !originalText.Contains("${") || !originalText.Contains('['))
                    continue;

                // Check if this contains a complete expression
                if (ContainsCompleteExpression(originalText))
                {
                    string adjustedText = ExpressionAdjuster.AdjustExpressionIndices(originalText, context);
                    if (adjustedText != originalText)
                    {
                        Logger.Debug($"Adjusted run-level text expression: '{originalText}' -> '{adjustedText}'");
                        textElement.Text = adjustedText;
                    }
                }
            }
        }
    }

    /// <summary>
    /// Adjusts expressions at the Paragraph level for split expressions
    /// </summary>
    private static void AdjustParagraphLevelExpressions(SlidePart slidePart, SlideContext context)
    {
        Logger.Debug("Adjusting expressions at Paragraph level for split expressions");

        // Process expressions by paragraphs for cases where expressions span multiple runs
        var paragraphs = slidePart.Slide.Descendants<D.Paragraph>().ToList();
        Logger.Debug($"Found {paragraphs.Count} paragraphs to adjust");

        foreach (var paragraph in paragraphs)
        {
            // Skip paragraphs that have already been processed at the Run level
            if (paragraph.Descendants<D.Run>().All(r => r.Descendants<D.Text>().All(t =>
                string.IsNullOrEmpty(t.Text) || !t.Text.Contains("${") || ContainsCompleteExpression(t.Text))))
                continue;

            // Get all text elements within this paragraph
            var textElements = paragraph.Descendants<D.Text>().ToList();
            if (textElements.Count == 0)
                continue;

            // Combine all text in this paragraph
            StringBuilder combinedText = new StringBuilder();
            foreach (var textElement in textElements)
            {
                if (textElement.Text != null)
                    combinedText.Append(textElement.Text);
            }

            string originalText = combinedText.ToString();

            // Skip if no potential expressions
            if (string.IsNullOrEmpty(originalText) ||
                !(originalText.Contains("${") && originalText.Contains('[')))
                continue;

            // First, check for hierarchical paths using the configured delimiter
            string delimiter = PowerPointOptions.Current.HierarchyDelimiter;

            // If there's a hierarchy delimiter in the context collection name, this requires special handling
            if (context.CollectionName.Contains(delimiter))
            {
                // Log the context info for debugging
                Logger.Debug($"Adjusting expressions for hierarchical context: {context.CollectionName}, Offset: {context.Offset}");

                string[] contextSegments = context.CollectionName.Split(delimiter);

                // For expressions that match the full hierarchical path, adjust directly
                if (originalText.Contains(context.CollectionName))
                {
                    string adjustedText = ExpressionAdjuster.AdjustExpressionIndices(originalText, context);

                    // Only update if the text was changed
                    if (adjustedText != originalText)
                    {
                        Logger.Debug($"Adjusting hierarchical path expression: '{originalText}' -> '{adjustedText}'");
                        UpdateParagraphTextPreservingRuns(paragraph, textElements, originalText, adjustedText);
                        continue;
                    }
                }

                // For expressions with just the last segment, adjust based on context
                string lastSegment = contextSegments[contextSegments.Length - 1];
                if (originalText.Contains($"{lastSegment}["))
                {
                    // Create a specialized context just for the last segment
                    var segmentContext = new SlideContext
                    {
                        CollectionName = lastSegment,
                        Offset = context.Offset,
                        TotalItems = context.TotalItems,
                        CurrentItem = context.CurrentItem,
                        RootData = context.RootData,
                        ParentContext = context.ParentContext
                    };

                    string adjustedText = ExpressionAdjuster.AdjustExpressionIndices(originalText, segmentContext);

                    // Only update if the text was changed
                    if (adjustedText != originalText)
                    {
                        Logger.Debug($"Adjusting last segment expression: '{originalText}' -> '{adjustedText}'");
                        UpdateParagraphTextPreservingRuns(paragraph, textElements, originalText, adjustedText);
                        continue;
                    }
                }
            }

            // Default adjustment
            string adjustedFullText = ExpressionAdjuster.AdjustExpressionIndices(originalText, context);

            // Only update if the text was changed
            if (adjustedFullText != originalText)
            {
                Logger.Debug($"Adjusting paragraph text: '{originalText}' -> '{adjustedFullText}'");

                // Update text in the paragraph while preserving runs
                UpdateParagraphTextPreservingRuns(paragraph, textElements, originalText, adjustedFullText);
            }
        }
    }

    /// <summary>
    /// Updates paragraph text while preserving the original run structure and formatting
    /// </summary>
    private static void UpdateParagraphTextPreservingRuns(D.Paragraph paragraph, List<D.Text> textElements,
                                                         string originalText, string newText)
    {
        // Special case: if there's only one text element, update it directly
        if (textElements.Count == 1)
        {
            textElements[0].Text = newText;
            return;
        }

        // Create maps of the original run structure
        var runMap = new List<RunInfo>();
        int currentPosition = 0;

        foreach (var textElement in textElements)
        {
            if (textElement.Text == null)
                continue;

            int length = textElement.Text.Length;
            if (length > 0)
            {
                runMap.Add(new RunInfo
                {
                    TextElement = textElement,
                    StartPosition = currentPosition,
                    EndPosition = currentPosition + length - 1,
                    Run = textElement.Parent as D.Run
                });
                currentPosition += length;
            }
        }

        // If the text length hasn't changed, we can distribute it proportionally
        if (originalText.Length == newText.Length)
        {
            for (int i = 0; i < runMap.Count; i++)
            {
                var run = runMap[i];
                run.TextElement.Text = newText.Substring(run.StartPosition, run.EndPosition - run.StartPosition + 1);
            }
        }
        else
        {
            // For changed lengths, try to distribute based on expression boundaries
            // Find all expressions in the adjusted text
            var expressionRanges = FindExpressionRanges(newText);

            if (expressionRanges.Count > 0)
            {
                DistributeTextWithExpressions(runMap, newText, expressionRanges);
            }
            else
            {
                // Fallback: put all in first text element, clear others
                textElements[0].Text = newText;
                for (int i = 1; i < textElements.Count; i++)
                {
                    textElements[i].Text = string.Empty;
                }
            }
        }
    }

    /// <summary>
    /// Distributes text containing expressions across run elements
    /// </summary>
    private static void DistributeTextWithExpressions(List<RunInfo> runMap, string text, List<ExpressionRange> expressionRanges)
    {
        // Simple distribution: try to keep whole expressions in single runs when possible

        // Strategy: assign each expression to a run if possible,
        // and distribute remaining text proportionally

        if (runMap.Count == 0)
            return;

        // Clear all text elements
        foreach (var run in runMap)
        {
            run.TextElement.Text = string.Empty;
        }

        // Find non-expression text segments
        var textSegments = new List<TextSegment>();
        int lastEnd = 0;

        foreach (var range in expressionRanges)
        {
            // Add segment before this expression
            if (range.Start > lastEnd)
            {
                textSegments.Add(new TextSegment
                {
                    Start = lastEnd,
                    End = range.Start - 1,
                    Text = text.Substring(lastEnd, range.Start - lastEnd)
                });
            }

            // Add the expression segment
            textSegments.Add(new TextSegment
            {
                Start = range.Start,
                End = range.End,
                Text = range.Expression,
                IsExpression = true
            });

            lastEnd = range.End + 1;
        }

        // Add any text after the last expression
        if (lastEnd < text.Length)
        {
            textSegments.Add(new TextSegment
            {
                Start = lastEnd,
                End = text.Length - 1,
                Text = text.Substring(lastEnd)
            });
        }

        // Now assign segments to runs
        int currentRun = 0;

        foreach (var segment in textSegments)
        {
            if (currentRun >= runMap.Count)
                currentRun = runMap.Count - 1; // Use last run if we run out

            // For expressions, try to keep them in a single run
            if (segment.IsExpression)
            {
                // If this run already has text and we have more runs available, go to next run
                if (runMap[currentRun].TextElement.Text.Length > 0 && currentRun < runMap.Count - 1)
                    currentRun++;

                runMap[currentRun].TextElement.Text += segment.Text;

                // Move to next run after an expression
                if (currentRun < runMap.Count - 1)
                    currentRun++;
            }
            else
            {
                // For non-expression text, add to current run
                runMap[currentRun].TextElement.Text += segment.Text;
            }
        }
    }
}