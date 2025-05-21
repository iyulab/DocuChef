using DocuChef.Presentation.Directives;
using DocumentFormat.OpenXml.Presentation;

namespace DocuChef.Presentation.Processing;

/// <summary>
/// Processes template slides to extract directives and structure
/// </summary>
internal class TemplateProcessor
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
    /// Analyzes all slides in the template to identify directives and structure
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

        // Process each slide in the template
        foreach (SlideId slideId in slideIdList.Elements<SlideId>())
        {
            var slideInfo = ProcessTemplateSlide(presentationPart, slideId);
            if (slideInfo != null)
            {
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
            switch (slide.Type)
            {
                case SlideType.Foreach:
                    foreachCount++;
                    if (slide.Directive is ForeachDirective foreachDirective)
                    {
                        collections.Add(foreachDirective.CollectionName);
                    }
                    break;
                case SlideType.If:
                    ifCount++;
                    break;
                default:
                    regularCount++;
                    break;
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
                HasImplicitDirective = hasImplicitDirective
            };

            Logger.Debug($"Processed template slide {slideId.Id}: {(slideInfo.HasDirective ? slideInfo.Type.ToString() : "Regular slide")}");
            return slideInfo;
        }
        catch (Exception ex)
        {
            Logger.Warning($"Error processing slide {slideId.Id}: {ex.Message}");
            return null;
        }
    }

    /// <summary>
    /// Tries to create an implicit directive by analyzing expressions in the slide
    /// </summary>
    private Directive TryCreateImplicitDirective(SlidePart slidePart)
    {
        try
        {
            // Extract all text elements from the slide
            var textElements = slidePart.Slide.Descendants<D.Text>().ToList();
            if (textElements.Count == 0)
                return null;

            // Concatenate all text to analyze expressions
            var allText = new StringBuilder();
            foreach (var text in textElements)
            {
                if (!string.IsNullOrEmpty(text.Text))
                {
                    allText.AppendLine(text.Text);
                }
            }

            // Try to extract implicit directives from expressions
            var implicitDirectives = ExpressionParser.CreateImplicitDirectives(allText.ToString());
            if (implicitDirectives.Count == 0)
                return null;

            // Find the directive with the highest priority (typically the one with the most complex path)
            // For now we'll use a simple approach: take the one with the most segments in collection path
            var priorityDirective = implicitDirectives
                .OrderByDescending(d => d.CollectionPath.Count(c => c == '_') + 1)
                .ThenByDescending(d => d.MaxItems)
                .FirstOrDefault();

            if (priorityDirective == null)
                return null;

            // Create a ForeachDirective based on the implicit directive info
            return new ForeachDirective
            {
                CollectionName = priorityDirective.CollectionPath,
                MaxItems = priorityDirective.MaxItems
            };
        }
        catch (Exception ex)
        {
            Logger.Warning($"Error creating implicit directive: {ex.Message}");
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
            string collectionName = entry.Directive.CollectionName;
            int maxItems = entry.Directive.MaxItems;

            // Only store if not already stored or if the new value is smaller
            if (!result.MaxItemsPerSlide.ContainsKey(collectionName) ||
                result.MaxItemsPerSlide[collectionName] > maxItems)
            {
                result.MaxItemsPerSlide[collectionName] = maxItems;
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
            .Where(s => s.Type == SlideType.Foreach)
            .Select(s => new ForeachSlideInfo
            {
                Slide = s,
                Directive = s.Directive as ForeachDirective
            })
            .Where(x => x.Directive != null)
            .ToList();
    }

    /// <summary>
    /// Processes a collection to determine its hierarchy
    /// </summary>
    private void ProcessCollectionHierarchy(
        string collection,
        List<ForeachSlideInfo> foreachSlides,
        HierarchyInfo result)
    {
        var segments = collection.Split('_');
        int level = segments.Length - 1; // 0 for top-level, 1+ for nested

        // Add to level dictionary
        if (!result.CollectionsByLevel.ContainsKey(level))
        {
            result.CollectionsByLevel[level] = new List<string>();
        }

        result.CollectionsByLevel[level].Add(collection);
        result.CollectionLevel[collection] = level;

        // Map slides to collections
        var slidesUsingCollection = foreachSlides
            .Where(x => x.Directive.CollectionName == collection)
            .Select(x => x.Slide)
            .ToList();

        result.SlidesByCollection[collection] = slidesUsingCollection;

        // Find parent-child relationships
        if (level > 0)
        {
            AddParentChildRelationship(collection, segments, result);
        }
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
    /// Adds parent-child relationship for a collection
    /// </summary>
    private void AddParentChildRelationship(
        string collection,
        string[] segments,
        HierarchyInfo result)
    {
        var parentPath = string.Join("_", segments.Take(segments.Length - 1));
        result.ParentCollections[collection] = parentPath;

        // Add this collection to the child collections of the parent
        if (!result.ChildCollections.ContainsKey(parentPath))
        {
            result.ChildCollections[parentPath] = new List<string>();
        }

        result.ChildCollections[parentPath].Add(collection);
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

/// <summary>
/// Contains information about collection hierarchy
/// </summary>
internal class HierarchyInfo
{
    /// <summary>
    /// Collections grouped by their nesting level (0 for top-level)
    /// </summary>
    public Dictionary<int, List<string>> CollectionsByLevel { get; set; } = new Dictionary<int, List<string>>();

    /// <summary>
    /// Maps collection name to its hierarchy level
    /// </summary>
    public Dictionary<string, int> CollectionLevel { get; set; } = new Dictionary<string, int>();

    /// <summary>
    /// Maps each collection to its slides in the template
    /// </summary>
    public Dictionary<string, List<SlideInfo>> SlidesByCollection { get; set; } = new Dictionary<string, List<SlideInfo>>();

    /// <summary>
    /// Maps each nested collection to its parent collection
    /// </summary>
    public Dictionary<string, string> ParentCollections { get; set; } = new Dictionary<string, string>();

    /// <summary>
    /// Maps each collection to its child collections
    /// </summary>
    public Dictionary<string, List<string>> ChildCollections { get; set; } = new Dictionary<string, List<string>>();

    /// <summary>
    /// Maps each collection to its max items per slide
    /// </summary>
    public Dictionary<string, int> MaxItemsPerSlide { get; set; } = new Dictionary<string, int>();
}