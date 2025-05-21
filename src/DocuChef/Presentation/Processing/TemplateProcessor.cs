using DocuChef.Presentation.Core;
using DocuChef.Presentation.Directives;
using DocuChef.Presentation.Models;

namespace DocuChef.Presentation.Processing;

/// <summary>
/// Processes template slides to extract directives and structure
/// </summary>
internal partial class TemplateProcessor
{
    private readonly DirectiveProcessor _directiveProcessor;

    /// <summary>
    /// Initializes a new instance of the TemplateProcessor class
    /// </summary>
    public TemplateProcessor()
    {
        _directiveProcessor = new DirectiveProcessor();
    }

    /// <summary>
    /// Initializes a new instance of the TemplateProcessor class with a custom directive processor
    /// </summary>
    public TemplateProcessor(DirectiveProcessor directiveProcessor)
    {
        _directiveProcessor = directiveProcessor ?? throw new ArgumentNullException(nameof(directiveProcessor));
    }

    /// <summary>
    /// Analyzes all slides in the template to identify directives and structure while preserving original order
    /// </summary>
    public List<SlideInfo> AnalyzeTemplateSlides(PresentationPart presentationPart)
    {
        var result = new List<SlideInfo>();

        if (presentationPart == null)
            return result;

        // Get the slide ID list
        SlideIdList slideIdList = presentationPart.Presentation.SlideIdList;
        if (slideIdList == null)
        {
            Logger.Debug("No slides found in template");
            return result;
        }

        // Process each slide in the template in the order they appear in the presentation
        // This is critical for maintaining the correct slide order
        foreach (SlideId slideId in slideIdList.Elements<SlideId>())
        {
            var slideInfo = ProcessTemplateSlide(presentationPart, slideId);
            if (slideInfo != null)
            {
                // Set appropriate slide type based on directive
                if (slideInfo.HasDirective)
                {
                    slideInfo.Type = SlideType.Source; // Slides with directives are typically sources for cloning
                }
                else
                {
                    slideInfo.Type = SlideType.Original; // Regular slides with no directives
                }

                Logger.Debug($"Processed slide {slideId.Id}: {slideInfo.Type}, Directive: {(slideInfo.HasDirective ? slideInfo.DirectiveType.ToString() : "None")}");
                result.Add(slideInfo);
            }
        }

        // Log summary of analysis
        Logger.Info($"Analyzed {result.Count} slides in template");
        LogDirectiveSummary(result);

        return result;
    }

    /// <summary>
    /// Logs a summary of all directives found in the template
    /// </summary>
    private void LogDirectiveSummary(List<SlideInfo> slides)
    {
        int foreachCount = 0;
        int ifCount = 0;
        int regularCount = 0;
        int implicitForeachCount = 0;

        var collections = new HashSet<string>();

        foreach (var slide in slides)
        {
            if (slide.DirectiveType == DirectiveType.Foreach)
            {
                foreachCount++;
                if (slide.Directive is ForeachDirective foreachDirective)
                {
                    collections.Add(foreachDirective.CollectionName);
                }
            }
            else if (slide.DirectiveType == DirectiveType.If)
            {
                ifCount++;
            }
            else
            {
                regularCount++;
            }

            // Track implicit directives
            if (slide.HasImplicitDirective)
            {
                implicitForeachCount++;
            }
        }

        Logger.Info($"Template contains: {regularCount} regular slides, {foreachCount} foreach slides ({implicitForeachCount} implicit), {ifCount} if slides");
        if (collections.Count > 0)
        {
            Logger.Info($"Collections used: {string.Join(", ", collections)}");
        }
    }

    /// <summary>
    /// Processes a single template slide to extract its information and directives
    /// </summary>
    private SlideInfo ProcessTemplateSlide(PresentationPart presentationPart, SlideId slideId)
    {
        // Extract the relationship ID
        string relationshipId = slideId.RelationshipId?.Value;
        if (string.IsNullOrEmpty(relationshipId))
        {
            Logger.Debug($"Slide ID {slideId.Id} has no relationship ID, skipping");
            return null;
        }

        try
        {
            SlidePart slidePart = (SlidePart)presentationPart.GetPartById(relationshipId);
            if (slidePart == null)
                return null;

            // Extract slide note text to check for directives
            string noteText = SlideManager.GetSlideNoteText(slidePart);

            // Parse directive from note text
            Directive directive = _directiveProcessor.Parse(noteText);

            // If no explicit directive found, try to derive implicit directives from expressions
            bool hasImplicitDirective = false;
            if (directive == null)
            {
                directive = TryCreateImplicitDirective(slidePart);
                hasImplicitDirective = directive != null;

                if (hasImplicitDirective)
                {
                    Logger.Debug($"Created implicit directive for slide {slideId.Id}: {directive}");
                }
            }

            // Create slide info
            var slideInfo = new SlideInfo
            {
                SlideId = slideId.Id,
                RelationshipId = relationshipId,
                NoteText = noteText,
                Directive = directive,
                HasImplicitDirective = hasImplicitDirective,
                // Default to Original, will be adjusted after analysis
                Type = SlideType.Original
            };

            Logger.Debug($"Processed template slide {slideId.Id}: Directive type: {slideInfo.DirectiveType}");
            return slideInfo;
        }
        catch (Exception ex)
        {
            Logger.Warning($"Error processing slide {slideId.Id}: {ex.Message}");
            return null;
        }
    }

    /// <summary>
    /// Tries to create an implicit directive by analyzing expressions in the slide using paragraph-based processing
    /// </summary>
    private Directive TryCreateImplicitDirective(SlidePart slidePart)
    {
        try
        {
            Logger.Debug("Starting paragraph-based implicit directive analysis");

            // First extract all paragraphs directly instead of text elements
            var paragraphs = slidePart.Slide.Descendants<D.Paragraph>().ToList();
            if (paragraphs.Count == 0)
            {
                Logger.Debug("No paragraphs found in slide");
                return null;
            }

            Logger.Debug($"Found {paragraphs.Count} paragraphs in slide");

            // Analyze each paragraph as a coherent unit
            var allText = new StringBuilder();
            foreach (var paragraph in paragraphs)
            {
                // Get all text elements within this paragraph and reconstruct the full text
                var paragraphText = new StringBuilder();

                // Use the order of elements as they appear in the XML
                foreach (var textElement in paragraph.Descendants<D.Text>())
                {
                    if (!string.IsNullOrEmpty(textElement.Text))
                    {
                        paragraphText.Append(textElement.Text);
                    }
                }

                // Only process non-empty paragraphs
                string fullParagraphText = paragraphText.ToString();
                if (!string.IsNullOrEmpty(fullParagraphText))
                {
                    allText.AppendLine(fullParagraphText);
                    Logger.Debug($"Reconstructed paragraph text: '{fullParagraphText}'");
                }
            }

            // Try to extract implicit directives from all the reconstructed text
            var combinedText = allText.ToString();
            if (string.IsNullOrEmpty(combinedText))
            {
                Logger.Debug("No text content found in paragraphs");
                return null;
            }

            Logger.Debug($"Combined paragraph-based text for analysis: '{combinedText}'");

            var implicitDirectives = ExpressionParser.CreateImplicitDirectives(combinedText);
            if (implicitDirectives.Count == 0)
            {
                Logger.Debug("No implicit directives found in paragraph text");
                return null;
            }

            // Log all found directives
            foreach (var dir in implicitDirectives)
            {
                Logger.Debug($"Found implicit directive: {dir}");
            }

            // Find the directive with the highest priority
            // Prioritize: 1) Most collection segments 2) Highest max items 3) Uses hierarchy delimiter
            var priorityDirective = implicitDirectives
                .OrderByDescending(d => d.CollectionPath.Count(c => c == PowerPointOptions.Current.HierarchyDelimiter[0]) + 1)
                .ThenByDescending(d => d.MaxItems)
                .ThenByDescending(d => d.CollectionPath.Contains(PowerPointOptions.Current.HierarchyDelimiter))
                .FirstOrDefault();

            if (priorityDirective == null)
            {
                Logger.Debug("No priority directive selected");
                return null;
            }

            Logger.Debug($"Selected priority directive: {priorityDirective}");

            // Create a ForeachDirective based on the implicit directive info
            return new ForeachDirective
            {
                CollectionName = priorityDirective.CollectionPath,
                MaxItems = priorityDirective.MaxItems
            };
        }
        catch (Exception ex)
        {
            Logger.Warning($"Error creating implicit directive: {ex.Message}", ex);
            return null;
        }
    }

    /// <summary>
    /// Analyzes slide hierarchy to determine collection dependencies and nesting levels
    /// </summary>
    public HierarchyInfo AnalyzeSlideHierarchy(List<SlideInfo> slides)
    {
        var result = new HierarchyInfo();

        // Find all foreach slides
        var foreachSlides = GetForeachSlides(slides);

        // Gather all collection names
        var allCollections = foreachSlides
            .Select(x => x.Directive.CollectionName)
            .Distinct()
            .ToList();

        Logger.Debug($"Found {allCollections.Count} unique collection names in template");

        // Process each collection to build hierarchy
        foreach (var collection in allCollections)
        {
            ProcessCollectionHierarchy(collection, foreachSlides, result);
        }

        // Store max items per slide for each collection
        foreach (var entry in foreachSlides)
        {
            var directive = entry.Directive as ForeachDirective;
            if (directive != null)
            {
                string collectionName = directive.CollectionName;
                int maxItems = directive.MaxItems;

                // Only store if not already stored or if the new value is smaller
                if (!result.MaxItemsPerSlide.ContainsKey(collectionName) ||
                    result.MaxItemsPerSlide[collectionName] > maxItems)
                {
                    result.MaxItemsPerSlide[collectionName] = maxItems;
                    Logger.Debug($"Set max items for collection '{collectionName}' to {maxItems}");
                }
            }
        }

        // Validate hierarchy correctness
        ValidateHierarchy(result);

        // Log hierarchy structure
        LogHierarchyStructure(result);

        return result;
    }

    /// <summary>
    /// Gets all foreach slides from the template
    /// </summary>
    private List<ForeachSlideInfo> GetForeachSlides(List<SlideInfo> slides)
    {
        return slides
            .Where(s => s.DirectiveType == DirectiveType.Foreach)
            .Select(s => new ForeachSlideInfo
            {
                Slide = s,
                Directive = s.Directive as ForeachDirective
            })
            .Where(x => x.Directive != null)
            .ToList();
    }

    /// <summary>
    /// Validates the hierarchy information for potential issues
    /// </summary>
    private void ValidateHierarchy(HierarchyInfo hierarchy)
    {
        // Check that all nested collections have valid parent collections
        foreach (var collection in hierarchy.CollectionLevel.Keys)
        {
            int level = hierarchy.CollectionLevel[collection];

            if (level > 0) // Nested collection
            {
                if (!hierarchy.ParentCollections.ContainsKey(collection))
                {
                    Logger.Warning($"Nested collection '{collection}' has no parent defined");
                    continue;
                }

                string parent = hierarchy.ParentCollections[collection];

                // Verify parent exists in our collections
                if (!hierarchy.CollectionLevel.ContainsKey(parent))
                {
                    Logger.Warning($"Collection '{collection}' references unknown parent '{parent}'");
                }

                // Verify parent is at the correct level
                int parentLevel = hierarchy.CollectionLevel[parent];
                if (parentLevel != level - 1)
                {
                    Logger.Warning($"Collection '{collection}' at level {level} has parent '{parent}' at incorrect level {parentLevel}");
                }
            }
        }
    }

    /// <summary>
    /// Logs the hierarchy structure
    /// </summary>
    private void LogHierarchyStructure(HierarchyInfo hierarchy)
    {
        Logger.Debug("Collection hierarchy structure:");
        foreach (var level in hierarchy.CollectionsByLevel.Keys.OrderBy(k => k))
        {
            Logger.Debug($"Level {level}: {string.Join(", ", hierarchy.CollectionsByLevel[level])}");

            // Log parent-child relationships for this level
            foreach (var collection in hierarchy.CollectionsByLevel[level])
            {
                if (level > 0)
                {
                    string parent = hierarchy.ParentCollections[collection];
                    Logger.Debug($"  - {collection} -> parent: {parent}");
                }

                if (hierarchy.ChildCollections.ContainsKey(collection))
                {
                    var children = hierarchy.ChildCollections[collection];
                    Logger.Debug($"  - {collection} -> children: {string.Join(", ", children)}");
                }
            }
        }
    }

    /// <summary>
    /// Helper class to store foreach slide information
    /// </summary>
    private class ForeachSlideInfo
    {
        public SlideInfo Slide { get; set; }
        public ForeachDirective Directive { get; set; }
    }
}