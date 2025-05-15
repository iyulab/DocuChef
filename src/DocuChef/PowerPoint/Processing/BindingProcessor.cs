using DocuChef.PowerPoint.Helpers;

namespace DocuChef.PowerPoint.Processing
{
    /// <summary>
    /// Processor responsible for applying variable bindings to slides
    /// </summary>
    internal class BindingProcessor
    {
        private readonly PowerPointProcessor _mainProcessor;
        private readonly PowerPointContext _context;
        private readonly DirectiveProcessor _directiveProcessor;
        private readonly ShapeProcessor _shapeProcessor;

        /// <summary>
        /// Initialize binding processor
        /// </summary>
        public BindingProcessor(PowerPointProcessor mainProcessor, PowerPointContext context)
        {
            _mainProcessor = mainProcessor ?? throw new ArgumentNullException(nameof(mainProcessor));
            _context = context ?? throw new ArgumentNullException(nameof(context));

            _directiveProcessor = new DirectiveProcessor(mainProcessor, context);
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
            _context.Slide.Notes = slidePart.GetNotes();

            // Process directives from notes
            var directives = DirectiveParser.ParseDirectives(_context.Slide.Notes);
            if (directives.Count > 0)
            {
                Logger.Debug($"Processing {directives.Count} directives from slide notes");
                foreach (var directive in directives)
                {
                    _directiveProcessor.ProcessShapeDirective(slidePart, directive);
                }
            }

            // Prepare variable context
            var variables = _mainProcessor.PrepareVariables();
            var textProcessor = new TextBindingProcessor(_mainProcessor, variables);

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
                _shapeProcessor.UpdateShapeContext(shape);

                // Process expression bindings
                if (shape.TextBody != null && shape.ContainsExpressions())
                {
                    if (textProcessor.ProcessShape(shape))
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
    }

    /// <summary>
    /// Processor for binding text expressions in shapes
    /// </summary>
    internal class TextBindingProcessor
    {
        private readonly PowerPointProcessor _mainProcessor;
        private readonly Dictionary<string, object> _variables;

        /// <summary>
        /// Initialize text binding processor
        /// </summary>
        public TextBindingProcessor(PowerPointProcessor processor, Dictionary<string, object> variables)
        {
            _mainProcessor = processor ?? throw new ArgumentNullException(nameof(processor));
            _variables = variables ?? new Dictionary<string, object>();
        }

        /// <summary>
        /// Process shape text with formatting preservation
        /// </summary>
        public bool ProcessShape(P.Shape shape)
        {
            if (shape?.TextBody == null)
                return false;

            // Check if shape contains expressions
            if (!shape.ContainsExpressions())
                return false;

            // Check if expressions span across runs
            bool crossRunExpressions = ContainsCrossRunExpressions(shape);

            return crossRunExpressions
                ? ProcessCrossRunExpressions(shape)
                : ProcessIndividualRuns(shape);
        }

        /// <summary>
        /// Process individual runs for expressions
        /// </summary>
        private bool ProcessIndividualRuns(P.Shape shape)
        {
            bool hasChanges = false;

            foreach (var paragraph in shape.TextBody.Elements<A.Paragraph>())
            {
                foreach (var run in paragraph.Elements<A.Run>())
                {
                    var textElement = run.GetFirstChild<A.Text>();
                    if (textElement == null || !DocuChef.Helpers.ExpressionHelper.ContainsExpressions(textElement.Text))
                        continue;

                    string processedText = DocuChef.Helpers.ExpressionHelper.ProcessExpressions(textElement.Text, _mainProcessor, _variables);
                    if (processedText != textElement.Text)
                    {
                        textElement.Text = processedText;
                        hasChanges = true;
                    }
                }
            }

            return hasChanges;
        }

        /// <summary>
        /// Process expressions that span across multiple runs
        /// </summary>
        private bool ProcessCrossRunExpressions(P.Shape shape)
        {
            bool hasChanges = false;

            foreach (var paragraph in shape.TextBody.Elements<A.Paragraph>().ToList())
            {
                // Reconstruct paragraph text
                var (paragraphText, runMappings) = ReconstructParagraphText(paragraph);

                if (!DocuChef.Helpers.ExpressionHelper.ContainsExpressions(paragraphText))
                    continue;

                // Process expressions
                string processedText = DocuChef.Helpers.ExpressionHelper.ProcessExpressions(paragraphText, _mainProcessor, _variables);
                if (processedText == paragraphText)
                    continue;

                // Map processed text back to runs
                if (MapProcessedTextToRuns(paragraph, runMappings, processedText))
                    hasChanges = true;
            }

            return hasChanges;
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
    }
}