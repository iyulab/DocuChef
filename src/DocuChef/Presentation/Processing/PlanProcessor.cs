using DocuChef.Presentation.Core;
using DocuChef.Presentation.Directives;
using DocuChef.Presentation.Models;

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

        // Create a list to keep track of processed slides
        var processedSlideIds = new HashSet<uint>();

        // Important: Process slides in their presentation order (as they appear in the template)
        // This ensures the final presentation maintains the original slide order
        foreach (var slide in templateSlides)
        {
            // Skip if already processed
            if (processedSlideIds.Contains(slide.SlideId))
                continue;

            // Process slide based on its directive type
            if (slide.DirectiveType == DirectiveType.Foreach)
            {
                // Process foreach slide with its collection
                var directive = slide.Directive as ForeachDirective;
                if (directive != null)
                {
                    ProcessForeachSlideAndAddToPosition(slide, directive, plan, dataSource, contextProcessor,
                        hierarchyInfo, processedSlideIds);
                }
            }
            else if (slide.DirectiveType == DirectiveType.If)
            {
                // Process if slide
                ProcessIfSlide(slide, plan, dataSource, processedSlideIds);
            }
            else
            {
                // Process regular slide with no directive
                ProcessRegularSlide(slide, plan, dataSource, processedSlideIds);
            }
        }

        return plan;
    }

    /// <summary>
    /// Process a regular slide
    /// </summary>
    private void ProcessRegularSlide(
        SlideInfo slide,
        PresentationPlan plan,
        object dataSource,
        HashSet<uint> processedSlideIds)
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

        // Mark as processed
        processedSlideIds.Add(slide.SlideId);
    }

    /// <summary>
    /// Process an if slide
    /// </summary>
    private void ProcessIfSlide(
        SlideInfo slideInfo,
        PresentationPlan plan,
        object dataSource,
        HashSet<uint> processedSlideIds)
    {
        var directive = slideInfo.Directive as IfDirective;
        if (directive == null)
        {
            Logger.Warning($"Invalid if directive for slide {slideInfo.SlideId}");
            plan.AddSlide(PlannedSlide.Create(slideInfo, null, SlideOperation.Keep));
            processedSlideIds.Add(slideInfo.SlideId);
            return;
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

        // Mark as processed
        processedSlideIds.Add(slideInfo.SlideId);
    }

    /// <summary>
    /// Process a foreach slide and its collection, adding the results at the current position in the plan
    /// </summary>
    private void ProcessForeachSlideAndAddToPosition(
        SlideInfo slide,
        ForeachDirective directive,
        PresentationPlan plan,
        object dataSource,
        ContextProcessor contextProcessor,
        HierarchyInfo hierarchyInfo,
        HashSet<uint> processedSlideIds)
    {
        string collectionName = directive.CollectionName;
        Logger.Debug($"Processing foreach slide for collection: {collectionName}");

        // Mark the foreach slide as processed 
        processedSlideIds.Add(slide.SlideId);

        // Check if this collection uses grouping
        bool useGrouping = directive.ShouldUseGroupedMode();
        int maxItemsPerSlide = directive.MaxItems;

        Logger.Debug($"Collection '{collectionName}' uses grouping: {useGrouping}, maxItemsPerSlide: {maxItemsPerSlide}");

        // Generate contexts for this collection - with support for grouping
        List<SlideContext> contexts;
        if (useGrouping)
        {
            // Use the grouping context generator for max items mode
            contexts = new List<SlideContext>();

            // Get the collection data
            IEnumerable collection = dataSource.GetCollection(collectionName);
            if (collection == null)
            {
                Logger.Warning($"Collection '{collectionName}' not found in data source");
                // Skip this slide if collection not found
                plan.AddSlide(PlannedSlide.Create(slide, null, SlideOperation.Skip));
                return;
            }

            // Count items in the collection
            int totalItems = collection.Count();
            Logger.Debug($"Collection '{collectionName}' has {totalItems} total items");

            if (totalItems == 0)
            {
                // Skip this slide if collection is empty
                plan.AddSlide(PlannedSlide.Create(slide, null, SlideOperation.Skip));
                return;
            }

            // Calculate number of groups
            int groupCount = (int)Math.Ceiling((double)totalItems / maxItemsPerSlide);
            Logger.Debug($"Collection '{collectionName}' will be divided into {groupCount} groups");

            // Create a context for each group
            for (int groupIndex = 0; groupIndex < groupCount; groupIndex++)
            {
                // Calculate start index and count for this group
                int startIndex = groupIndex * maxItemsPerSlide;
                int itemsToTake = Math.Min(maxItemsPerSlide, totalItems - startIndex);

                // Get items for this group
                var groupItems = collection.Cast<object>().Skip(startIndex).Take(itemsToTake).ToList();

                // Create context for this group
                var groupContext = SlideContext.Create(
                    collectionName,
                    startIndex, // Important - use correct offset for the group
                    groupItems,
                    dataSource,
                    totalItems);

                // Set number of items in the group
                groupContext.ItemsInGroup = groupItems.Count;

                contexts.Add(groupContext);
                Logger.Debug($"Created group context: Collection={collectionName}, Offset={startIndex}, Items={groupItems.Count}");
            }
        }
        else
        {
            // Use standard context generator for one-item-per-slide mode
            contexts = contextProcessor.GenerateContextsForCollection(collectionName);
            Logger.Debug($"Created {contexts.Count} individual contexts for collection '{collectionName}'");
        }

        if (contexts.Count == 0)
        {
            Logger.Debug($"No contexts generated for collection: {collectionName}");
            // Skip this slide if no contexts
            plan.AddSlide(PlannedSlide.Create(slide, null, SlideOperation.Skip));
            return;
        }

        // For each context of the collection
        for (int i = 0; i < contexts.Count; i++)
        {
            var context = contexts[i];
            bool isFirstContext = i == 0;

            // Add the foreach slide with this context
            SlideOperation operation = isFirstContext ? SlideOperation.Keep : SlideOperation.Clone;
            plan.AddSlide(PlannedSlide.Create(slide, context, operation));
            Logger.Debug($"Added {operation} for slide {slide.SlideId} with context {context.GetContextDescription()}");

            // Process any child collections for this context
            ProcessNestedCollectionsForContext(context, hierarchyInfo, plan, dataSource, contextProcessor, isFirstContext);
        }
    }

    /// <summary>
    /// Process nested collections for a specific parent context
    /// </summary>
    private void ProcessNestedCollectionsForContext(
        SlideContext parentContext,
        HierarchyInfo hierarchyInfo,
        PresentationPlan plan,
        object dataSource,
        ContextProcessor contextProcessor,
        bool isFirstParentContext)
    {
        string parentCollectionName = parentContext.CollectionName;

        // Find child collections of this parent
        if (!hierarchyInfo.ChildCollections.TryGetValue(parentCollectionName, out var childCollections) || childCollections.Count == 0)
        {
            Logger.Debug($"No child collections found for: {parentCollectionName}");
            return;
        }

        // Process each child collection
        foreach (var childCollectionName in childCollections)
        {
            Logger.Debug($"Processing nested collection: {childCollectionName} for parent: {parentCollectionName}");

            // Skip if no slides for this child collection
            if (!hierarchyInfo.SlidesByCollection.TryGetValue(childCollectionName, out var slidesForChild) || slidesForChild.Count == 0)
            {
                Logger.Debug($"No slides found for child collection: {childCollectionName}");
                continue;
            }

            // Get directive for this child collection
            var directive = slidesForChild[0].Directive as ForeachDirective;
            if (directive == null)
            {
                Logger.Warning($"Invalid directive for child collection: {childCollectionName}");
                continue;
            }

            int maxItemsPerSlide = directive.MaxItems;

            // Generate nested contexts for this child collection using the parent context
            var nestedContexts = contextProcessor.GenerateNestedContexts(childCollectionName, parentContext, maxItemsPerSlide);
            if (nestedContexts.Count == 0)
            {
                Logger.Debug($"No nested contexts generated for child collection: {childCollectionName}");
                continue;
            }

            // For each nested context
            for (int i = 0; i < nestedContexts.Count; i++)
            {
                var nestedContext = nestedContexts[i];
                bool isFirstNestedContext = i == 0 && isFirstParentContext;

                // Add slide for this nested context
                foreach (var slideInfo in slidesForChild)
                {
                    SlideOperation operation = isFirstNestedContext ? SlideOperation.Keep : SlideOperation.Clone;
                    plan.AddSlide(PlannedSlide.Create(slideInfo, nestedContext, operation));
                    Logger.Debug($"Added {operation} for {childCollectionName} with context {nestedContext.GetContextDescription()}");
                }

                // Process any deeper nested collections (if needed)
                ProcessNestedCollectionsForContext(nestedContext, hierarchyInfo, plan, dataSource, contextProcessor, isFirstNestedContext);
            }
        }
    }
}