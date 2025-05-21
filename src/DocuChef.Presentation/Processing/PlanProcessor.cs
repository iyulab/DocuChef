using DocuChef.Presentation.Directives;

namespace DocuChef.Presentation.Processing;

/// <summary>
/// Processes template slides to build a presentation plan
/// </summary>
internal class PlanProcessor
{
    private readonly TemplateProcessor _templateProcessor;
    private readonly DirectiveProcessor _directiveProcessor;

    /// <summary>
    /// Initializes a new instance of the PlanProcessor class
    /// </summary>
    public PlanProcessor()
    {
        _templateProcessor = new TemplateProcessor();
        _directiveProcessor = new DirectiveProcessor();
    }

    /// <summary>
    /// Builds a complete presentation plan with multi-level nesting support
    /// </summary>
    public PresentationPlan BuildPlan(
        List<SlideInfo> templateSlides,
        object dataSource,
        ContextProcessor contextProcessor)
    {
        var plan = new PresentationPlan();

        // Analyze slides to identify their hierarchy levels
        var hierarchyInfo = _templateProcessor.AnalyzeSlideHierarchy(templateSlides);

        // Group slides by type for easier processing
        var slidesMap = GroupSlidesByType(templateSlides);

        // Process regular slides first (no directives)
        ProcessRegularSlides(slidesMap.RegularSlides, plan, dataSource);

        // Process if slides
        ProcessIfSlides(slidesMap.IfSlides, plan, dataSource);

        // Generate collection contexts
        var collectionContexts = GenerateCollectionContexts(hierarchyInfo, contextProcessor);

        // Process foreach slides with interleaved structure
        ProcessInterleavedHierarchy(hierarchyInfo, plan, collectionContexts);

        return plan;
    }

    /// <summary>
    /// Groups slides by their type for easier processing
    /// </summary>
    private SlideTypeMap GroupSlidesByType(List<SlideInfo> templateSlides)
    {
        return new SlideTypeMap
        {
            RegularSlides = templateSlides.Where(s => s.Type == SlideType.Regular).ToList(),
            ForeachSlides = templateSlides.Where(s => s.Type == SlideType.Foreach).ToList(),
            IfSlides = templateSlides.Where(s => s.Type == SlideType.If).ToList()
        };
    }

    /// <summary>
    /// Processes regular slides (without directives)
    /// </summary>
    private void ProcessRegularSlides(List<SlideInfo> regularSlides, PresentationPlan plan, object dataSource)
    {
        foreach (var slide in regularSlides)
        {
            // Create root context for regular slides
            var slideContext = SlideContext.Create(
                "Root",
                0,
                dataSource,
                dataSource,
                1);

            // Keep regular slides in the plan
            plan.AddSlide(PlannedSlide.Create(
                slide,
                slideContext,
                SlideOperation.Keep
            ));
        }
    }

    /// <summary>
    /// Generates collection contexts for all collections
    /// </summary>
    private Dictionary<string, List<SlideContext>> GenerateCollectionContexts(
        HierarchyInfo hierarchy,
        ContextProcessor contextProcessor)
    {
        var result = new Dictionary<string, List<SlideContext>>();

        // Generate top-level collection contexts first
        if (hierarchy.CollectionsByLevel.ContainsKey(0))
        {
            foreach (var collection in hierarchy.CollectionsByLevel[0])
            {
                var contexts = contextProcessor.GenerateContextsForCollection(collection);
                result[collection] = contexts;

                Logger.Debug($"Generated {contexts.Count} contexts for top-level collection '{collection}'");
            }
        }

        // Now generate nested collection contexts based on parent contexts
        int maxLevel = hierarchy.CollectionsByLevel.Keys.Any() ?
            hierarchy.CollectionsByLevel.Keys.Max() : 0;

        for (int level = 1; level <= maxLevel; level++)
        {
            if (!hierarchy.CollectionsByLevel.ContainsKey(level))
                continue;

            foreach (var collection in hierarchy.CollectionsByLevel[level])
            {
                string parentCollection = hierarchy.ParentCollections[collection];

                if (!result.ContainsKey(parentCollection))
                {
                    Logger.Warning($"Parent collection '{parentCollection}' contexts not found for '{collection}'");
                    continue;
                }

                var nestedContexts = new List<SlideContext>();

                // For each parent context, generate nested contexts
                foreach (var parentContext in result[parentCollection])
                {
                    // Get max items parameter from directive
                    int maxItems = hierarchy.MaxItemsPerSlide.ContainsKey(collection) ?
                        hierarchy.MaxItemsPerSlide[collection] : 0;

                    var childContexts = contextProcessor.GenerateNestedContexts(
                        collection,
                        parentContext,
                        maxItems);

                    nestedContexts.AddRange(childContexts);
                }

                result[collection] = nestedContexts;
                Logger.Debug($"Generated {nestedContexts.Count} contexts for nested collection '{collection}'");
            }
        }

        return result;
    }

    /// <summary>
    /// Processes hierarchy with interleaved parent-child structure
    /// </summary>
    private void ProcessInterleavedHierarchy(
        HierarchyInfo hierarchy,
        PresentationPlan plan,
        Dictionary<string, List<SlideContext>> collectionContexts)
    {
        // We only have interleaved structure for top-level and their immediate children
        if (!hierarchy.CollectionsByLevel.ContainsKey(0))
            return;

        // Get top-level collections
        var topLevelCollections = hierarchy.CollectionsByLevel[0];

        foreach (var topCollection in topLevelCollections)
        {
            // Get templates for this collection
            if (!hierarchy.SlidesByCollection.ContainsKey(topCollection) ||
                !collectionContexts.ContainsKey(topCollection))
                continue;

            var topSlideTemplates = hierarchy.SlidesByCollection[topCollection];
            var topContexts = collectionContexts[topCollection];

            // Get child collections for this top-level collection
            List<string> childCollections = hierarchy.ChildCollections.ContainsKey(topCollection) ?
                hierarchy.ChildCollections[topCollection] : new List<string>();

            // Process each parent item with its children interleaved
            foreach (var topContext in topContexts)
            {
                // First add the parent item slide
                foreach (var topSlide in topSlideTemplates)
                {
                    // Determine operation
                    bool isFirst = !plan.IncludedSlides.Any(s =>
                        s.Source.SlideId == topSlide.SlideId);

                    SlideOperation operation = isFirst ?
                        SlideOperation.Keep : SlideOperation.Clone;

                    plan.AddSlide(PlannedSlide.Create(
                        topSlide,
                        topContext,
                        operation
                    ));

                    Logger.Debug($"Added {operation} for {topCollection} with context {topContext.GetContextDescription()}");
                }

                // Then add all child items for this parent
                ProcessChildrenForParent(
                    hierarchy,
                    plan,
                    collectionContexts,
                    topContext,
                    childCollections);
            }
        }
    }

    /// <summary>
    /// Processes child collection items for a specific parent context
    /// </summary>
    private void ProcessChildrenForParent(
        HierarchyInfo hierarchy,
        PresentationPlan plan,
        Dictionary<string, List<SlideContext>> collectionContexts,
        SlideContext parentContext,
        List<string> childCollections)
    {
        foreach (var childCollection in childCollections)
        {
            // Skip if no templates or contexts
            if (!hierarchy.SlidesByCollection.ContainsKey(childCollection) ||
                !collectionContexts.ContainsKey(childCollection))
                continue;

            var childSlideTemplates = hierarchy.SlidesByCollection[childCollection];
            var allChildContexts = collectionContexts[childCollection];

            // Filter child contexts for this specific parent
            var childContextsForParent = allChildContexts
                .Where(c => c.ParentContext?.Offset == parentContext.Offset &&
                           c.ParentContext?.CollectionName == parentContext.CollectionName)
                .ToList();

            // Process each child context
            foreach (var childContext in childContextsForParent)
            {
                foreach (var childSlide in childSlideTemplates)
                {
                    // Determine operation
                    bool isFirst = !plan.IncludedSlides.Any(s =>
                        s.Source.SlideId == childSlide.SlideId);

                    SlideOperation operation = isFirst ?
                        SlideOperation.Keep : SlideOperation.Clone;

                    plan.AddSlide(PlannedSlide.Create(
                        childSlide,
                        childContext,
                        operation
                    ));

                    Logger.Debug($"Added {operation} for {childCollection} with context {childContext.GetContextDescription()}");
                }
            }
        }
    }

    /// <summary>
    /// Processes slides with if directives
    /// </summary>
    private void ProcessIfSlides(List<SlideInfo> ifSlides, PresentationPlan plan, object dataSource)
    {
        foreach (var slideInfo in ifSlides)
        {
            var directive = slideInfo.Directive as IfDirective;
            if (directive == null)
            {
                Logger.Warning($"Invalid if directive for slide {slideInfo.SlideId}");
                plan.AddSlide(PlannedSlide.Create(slideInfo, null, SlideOperation.Keep));
                continue;
            }

            Logger.Debug($"Processing if directive: {directive}");

            // Create context for evaluation
            var slideContext = SlideContext.Create(
                "Root",
                0,
                dataSource,
                dataSource,
                1);

            // Evaluate condition
            bool conditionMet = directive.Evaluate(slideContext);
            Logger.Debug($"Condition '{directive.Condition}' evaluated to {conditionMet}");

            // Add to plan based on condition
            SlideOperation operation = conditionMet ? SlideOperation.Keep : SlideOperation.Skip;
            plan.AddSlide(PlannedSlide.Create(slideInfo, slideContext, operation));
        }
    }
}

/// <summary>
/// Helper class to group slides by type
/// </summary>
internal class SlideTypeMap
{
    /// <summary>
    /// Regular slides without directives
    /// </summary>
    public List<SlideInfo> RegularSlides { get; set; } = new List<SlideInfo>();

    /// <summary>
    /// Slides with foreach directives
    /// </summary>
    public List<SlideInfo> ForeachSlides { get; set; } = new List<SlideInfo>();

    /// <summary>
    /// Slides with if directives
    /// </summary>
    public List<SlideInfo> IfSlides { get; set; } = new List<SlideInfo>();
}