using DocuChef.PowerPoint.Helpers;

namespace DocuChef.PowerPoint.Processing;

/// <summary>
/// Processor responsible for applying variable bindings to slides
/// </summary>
internal class BindingProcessor
{
    private readonly PowerPointProcessor _mainProcessor;
    private readonly PowerPointContext _context;
    private readonly ExpressionProcessor _expressionProcessor;
    private readonly ShapeProcessor _shapeProcessor;

    /// <summary>
    /// Initialize binding processor
    /// </summary>
    public BindingProcessor(PowerPointProcessor mainProcessor, PowerPointContext context)
    {
        _mainProcessor = mainProcessor ?? throw new ArgumentNullException(nameof(mainProcessor));
        _context = context ?? throw new ArgumentNullException(nameof(context));

        _expressionProcessor = new ExpressionProcessor(mainProcessor, context);
        _shapeProcessor = new ShapeProcessor(mainProcessor, context);
    }

    /// <summary>
    /// Apply bindings to all visible shapes in slide
    /// </summary>
    public void ApplyBindings(SlidePart slidePart)
    {
        string slideId = _context.Variables.ContainsKey("_document") ?
            ((PresentationDocument)_context.Variables["_document"]).PresentationPart.GetIdOfPart(slidePart) :
            "unknown";

        Logger.Debug($"Applying bindings to slide {slideId}");

        // Set slide context
        _context.SlidePart = slidePart;
        _context.Slide = new SlideContext
        {
            Id = slideId,
            Notes = slidePart.GetNotes()
        };

        // Prepare variable context
        var variables = _mainProcessor.PrepareVariables();

        // Process directives again to ensure proper context for binding
        CheckSlideDirectives(slidePart, variables);

        // Apply bindings to all visible shapes
        var shapes = slidePart.Slide.Descendants<P.Shape>()
                    .Where(s => !PowerPointShapeHelper.IsShapeHidden(s))
                    .ToList();

        Logger.Debug($"Processing {shapes.Count} visible shapes");

        int processedShapeCount = 0;
        foreach (var shape in shapes)
        {
            bool shapeUpdated = false;

            // Update shape context
            UpdateShapeContext(shape);

            // Process expression bindings
            if (shape.TextBody != null && shape.ContainsExpressions())
            {
                if (ProcessShapeExpressions(shape, variables))
                {
                    shapeUpdated = true;
                    Logger.Debug($"Processed expressions in shape '{shape.GetShapeName()}'");
                }
            }

            // Process PowerPoint functions
            if (_shapeProcessor.ProcessPowerPointFunctions(shape))
            {
                shapeUpdated = true;
                Logger.Debug($"Processed PowerPoint functions in shape '{shape.GetShapeName()}'");
            }

            if (shapeUpdated)
                processedShapeCount++;
        }

        Logger.Debug($"Processed {processedShapeCount} shapes in slide");
        slidePart.Slide.Save();
    }

    /// <summary>
    /// Check slide directives to ensure proper context is set
    /// </summary>
    private void CheckSlideDirectives(SlidePart slidePart, Dictionary<string, object> variables)
    {
        string notes = slidePart.GetNotes();
        if (string.IsNullOrEmpty(notes))
            return;

        // Parse directives
        var directives = DirectiveParser.ParseDirectives(notes);
        if (!directives.Any())
            return;

        // Process foreach directives to ensure proper current index is set
        foreach (var directive in directives)
        {
            if (directive.Name.ToLowerInvariant() == "foreach")
            {
                string collectionName = directive.Value.Trim();
                if (string.IsNullOrEmpty(collectionName))
                    continue;

                // If this is a nested collection (Parent_Child format)
                if (collectionName.Contains('_'))
                {
                    string[] parts = collectionName.Split('_');
                    if (parts.Length >= 2)
                    {
                        string parentName = parts[0];
                        string childName = parts[1];

                        // Look for batch index information in processed slide IDs
                        string slideId = ((PresentationDocument)_context.Variables["_document"]).PresentationPart.GetIdOfPart(slidePart);
                        foreach (var processedId in _context.ProcessedArraySlides)
                        {
                            var match = System.Text.RegularExpressions.Regex.Match(processedId, $@"{parentName}_(\d+)_{childName}_(\d+)");
                            if (match.Success && match.Groups.Count > 2 && slideId.Contains(processedId))
                            {
                                if (int.TryParse(match.Groups[1].Value, out int parentIndex) &&
                                    int.TryParse(match.Groups[2].Value, out int childBatchIndex))
                                {
                                    // Set parent index to ensure proper context
                                    _context.SetCollectionIndex(parentName, parentIndex);

                                    // Calculate batch start index for child collection
                                    int maxItemsPerSlide = directive.GetParameterAsInt("max", int.MaxValue);
                                    int childStartIndex = childBatchIndex * maxItemsPerSlide;

                                    // Store in context for reference during binding
                                    if (!_context.CurrentIndices.ContainsKey($"{childName}_batch_start"))
                                    {
                                        _context.CurrentIndices[$"{childName}_batch_start"] = childStartIndex;
                                    }

                                    Logger.Debug($"Set context for nested collection: {parentName}[{parentIndex}].{childName}, batch start: {childStartIndex}");
                                    break;
                                }
                            }
                        }
                    }
                }
                // Simple collection (non-nested)
                else
                {
                    // Look for batch information
                    string slideId = ((PresentationDocument)_context.Variables["_document"]).PresentationPart.GetIdOfPart(slidePart);
                    var batchMatch = System.Text.RegularExpressions.Regex.Match(slideId, $@"batch_{collectionName}_(\d+)");
                    if (batchMatch.Success && batchMatch.Groups.Count > 1)
                    {
                        if (int.TryParse(batchMatch.Groups[1].Value, out int batchIndex))
                        {
                            // Calculate batch start index
                            int maxItemsPerSlide = directive.GetParameterAsInt("max", int.MaxValue);
                            int startIndex = batchIndex * maxItemsPerSlide;

                            // Store in context
                            _context.CurrentIndices[collectionName] = startIndex;
                            Logger.Debug($"Set context for collection batch: {collectionName}[{startIndex}]");
                        }
                    }
                }
            }
        }
    }

    /// <summary>
    /// Update shape context
    /// </summary>
    private void UpdateShapeContext(P.Shape shape)
    {
        _context.Shape = new ShapeContext
        {
            Name = shape.GetShapeName(),
            Id = shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Id?.Value.ToString(),
            Text = shape.GetText(),
            Type = GetShapeType(shape),
            ShapeObject = shape
        };
    }

    /// <summary>
    /// Get shape type
    /// </summary>
    private string GetShapeType(P.Shape shape)
    {
        var presetGeometry = shape.ShapeProperties?.GetFirstChild<A.PresetGeometry>();
        if (presetGeometry?.Preset != null)
        {
            return presetGeometry.Preset.Value.ToString();
        }

        return shape.TextBody != null ? "TextBox" : "Shape";
    }

    /// <summary>
    /// Process expression bindings in shape text
    /// </summary>
    private bool ProcessShapeExpressions(P.Shape shape, Dictionary<string, object> variables)
    {
        if (shape.TextBody == null)
            return false;

        bool hasChanges = false;

        // Check if expressions span across runs
        bool crossRunExpressions = ContainsCrossRunExpressions(shape);

        if (crossRunExpressions)
        {
            // Process expressions that span across multiple runs
            foreach (var paragraph in shape.TextBody.Elements<A.Paragraph>().ToList())
            {
                // Reconstruct paragraph text
                var (paragraphText, runMappings) = ReconstructParagraphText(paragraph);

                if (!DocuChef.Helpers.ExpressionHelper.ContainsExpressions(paragraphText))
                    continue;

                // Process expressions
                string processedText = DocuChef.Helpers.ExpressionHelper.ProcessExpressions(paragraphText, _mainProcessor, variables);
                if (processedText == paragraphText)
                    continue;

                // Map processed text back to runs
                if (MapProcessedTextToRuns(paragraph, runMappings, processedText))
                    hasChanges = true;
            }
        }
        else
        {
            // Process individual runs
            foreach (var paragraph in shape.TextBody.Elements<A.Paragraph>())
            {
                foreach (var run in paragraph.Elements<A.Run>())
                {
                    var textElement = run.GetFirstChild<A.Text>();
                    if (textElement == null || !DocuChef.Helpers.ExpressionHelper.ContainsExpressions(textElement.Text))
                        continue;

                    string processedText = DocuChef.Helpers.ExpressionHelper.ProcessExpressions(textElement.Text, _mainProcessor, variables);
                    if (processedText != textElement.Text)
                    {
                        textElement.Text = processedText;
                        hasChanges = true;
                    }
                }
            }
        }

        return hasChanges;
    }

    /// <summary>
    /// Check if expressions span across runs
    /// </summary>
    private bool ContainsCrossRunExpressions(P.Shape shape)
    {
        foreach (var paragraph in shape.TextBody.Elements<A.Paragraph>())
        {
            bool hasOpenBrace = false;

            foreach (var run in paragraph.Elements<A.Run>())
            {
                var textElement = run.GetFirstChild<A.Text>();
                if (textElement == null || string.IsNullOrEmpty(textElement.Text))
                    continue;

                string text = textElement.Text;

                // Check for complete expressions
                if (text.Contains("${") && text.Contains("}"))
                    continue;

                // Check for partial expressions
                if (text.Contains("${"))
                    hasOpenBrace = true;

                if (text.Contains("}") && hasOpenBrace)
                    return true;
            }

            if (hasOpenBrace)
                return true;
        }

        return false;
    }

    /// <summary>
    /// Reconstruct complete paragraph text
    /// </summary>
    private (string Text, List<(A.Run Run, int StartPos, int Length)> RunMappings) ReconstructParagraphText(A.Paragraph paragraph)
    {
        var sb = new StringBuilder();
        var runMappings = new List<(A.Run Run, int StartPos, int Length)>();

        foreach (var run in paragraph.Elements<A.Run>())
        {
            var textElement = run.GetFirstChild<A.Text>();
            if (textElement != null && !string.IsNullOrEmpty(textElement.Text))
            {
                int startPos = sb.Length;
                string text = textElement.Text;
                sb.Append(text);
                runMappings.Add((run, startPos, text.Length));
            }
        }

        return (sb.ToString(), runMappings);
    }

    /// <summary>
    /// Map processed text back to runs
    /// </summary>
    private bool MapProcessedTextToRuns(A.Paragraph paragraph, List<(A.Run Run, int StartPos, int Length)> runMappings, string processedText)
    {
        if (runMappings.Count == 0)
            return false;

        // Simple case: single run
        if (runMappings.Count == 1)
        {
            var textElement = runMappings[0].Run.GetFirstChild<A.Text>();
            if (textElement != null)
            {
                textElement.Text = processedText;
                return true;
            }
            return false;
        }

        // Complex case: distribute text across runs
        DistributeTextAcrossRuns(runMappings, processedText);
        return true;
    }

    /// <summary>
    /// Distribute text across multiple runs
    /// </summary>
    private void DistributeTextAcrossRuns(List<(A.Run Run, int StartPos, int Length)> runMappings, string processedText)
    {
        // Calculate distribution ratio
        double ratio = (double)processedText.Length / runMappings.Sum(r => r.Length);
        int currentPos = 0;

        for (int i = 0; i < runMappings.Count; i++)
        {
            var runInfo = runMappings[i];
            runInfo.Run.RemoveAllChildren<A.Text>();

            // Calculate text portion for this run
            string runText;
            if (i == runMappings.Count - 1)
            {
                // Last run gets remaining text
                runText = currentPos < processedText.Length
                    ? processedText.Substring(currentPos)
                    : string.Empty;
            }
            else
            {
                // Distribute proportionally
                int newLength = (int)Math.Ceiling(runInfo.Length * ratio);
                newLength = Math.Min(newLength, processedText.Length - currentPos);

                runText = newLength > 0 && currentPos < processedText.Length
                    ? processedText.Substring(currentPos, newLength)
                    : string.Empty;

                currentPos += runText.Length;
            }

            runInfo.Run.AppendChild(new A.Text(runText));
        }
    }
}