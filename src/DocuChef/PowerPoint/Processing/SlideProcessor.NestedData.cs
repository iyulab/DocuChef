using DocuChef.PowerPoint.Helpers;
using DocuChef.PowerPoint.Processing.ArrayProcessing;

namespace DocuChef.PowerPoint.Processing;

/// <summary>
/// Extension to SlideProcessor for handling nested data structures
/// </summary>
internal partial class SlideProcessor
{
    /// <summary>
    /// Analyze slide and prepare it for nested data processing
    /// </summary>
    public void AnalyzeAndPrepareNestedData(PresentationPart presentationPart, SlidePart slidePart)
    {
        string slideId = presentationPart.GetIdOfPart(slidePart);
        Logger.Debug($"Analyzing slide {slideId} for nested data structures");

        var variables = _mainProcessor.PrepareVariables();

        // First check for slide notes to find directives
        string notes = slidePart.GetNotes();
        if (string.IsNullOrEmpty(notes))
        {
            Logger.Debug("No notes found in slide, skipping nested data processing");
            return;
        }

        // Parse directives
        var directives = DirectiveParser.ParseDirectives(notes);
        if (!directives.Any())
        {
            Logger.Debug("No directives found in slide notes");
            return;
        }

        // Look for nested foreach directives (foreach-nested or with Parent_Child notation)
        var nestedForEachDirectives = directives.Where(d =>
            d.Name.Equals("foreach-nested", StringComparison.OrdinalIgnoreCase) ||
            (d.Name.Equals("foreach", StringComparison.OrdinalIgnoreCase) && d.Value.Contains('_'))
        ).ToList();

        if (!nestedForEachDirectives.Any())
        {
            Logger.Debug("No nested foreach directives found");
            return;
        }

        // Process each nested foreach directive
        foreach (var directive in nestedForEachDirectives)
        {
            ProcessNestedForEachDirective(presentationPart, slidePart, directive, variables);
        }
    }

    /// <summary>
    /// Process a nested foreach directive
    /// </summary>
    private void ProcessNestedForEachDirective(
        PresentationPart presentationPart,
        SlidePart slidePart,
        SlideDirective directive,
        Dictionary<string, object> variables)
    {
        try
        {
            string collectionPath = directive.Value.Trim();
            Logger.Debug($"Processing nested foreach directive for collection: {collectionPath}");

            // Handle different formats of nested collection references
            string parentName, childName;

            if (directive.Name.Equals("foreach-nested", StringComparison.OrdinalIgnoreCase))
            {
                // Format: #foreach-nested: Collection, parent: "Parent", child: "Child"
                parentName = directive.GetParameter("parent");
                childName = directive.GetParameter("child");

                if (string.IsNullOrEmpty(parentName) || string.IsNullOrEmpty(childName))
                {
                    Logger.Warning("Nested foreach directive missing parent or child parameter");
                    return;
                }
            }
            else
            {
                // Format: #foreach: Parent_Child
                var parts = collectionPath.Split('_');
                if (parts.Length < 2)
                {
                    Logger.Warning($"Invalid nested collection path: {collectionPath}");
                    return;
                }

                parentName = parts[0];
                childName = parts[1];
            }

            // Get extra parameters
            int maxItemsPerSlide = directive.GetParameterAsInt("max", -1);
            int offset = directive.GetParameterAsInt("offset", 0);

            Logger.Debug($"Nested foreach parameters: parent={parentName}, child={childName}, " +
                         $"max={maxItemsPerSlide}, offset={offset}");

            // Ensure the parent collection exists
            if (!variables.TryGetValue(parentName, out var parentObj) || parentObj == null)
            {
                Logger.Warning($"Parent collection not found: {parentName}");
                return;
            }

            // Get the parent collection count
            int parentCount = CollectionHelper.GetCollectionCount(parentObj);
            if (parentCount == 0)
            {
                Logger.Warning($"Parent collection is empty: {parentName}");
                return;
            }

            Logger.Debug($"Parent collection {parentName} has {parentCount} items");

            // Process each parent item
            for (int parentIndex = 0; parentIndex < parentCount; parentIndex++)
            {
                ProcessNestedCollectionItem(
                    presentationPart,
                    slidePart,
                    parentName,
                    childName,
                    parentIndex,
                    maxItemsPerSlide,
                    offset);
            }
        }
        catch (Exception ex)
        {
            Logger.Error($"Error processing nested foreach directive: {ex.Message}", ex);
        }
    }

    /// <summary>
    /// Process a single parent item with its nested collection
    /// </summary>
    private void ProcessNestedCollectionItem(
        PresentationPart presentationPart,
        SlidePart slidePart,
        string parentName,
        string childName,
        int parentIndex,
        int maxItemsPerSlide,
        int offset)
    {
        try
        {
            Logger.Debug($"Processing {parentName}[{parentIndex}].{childName}");

            // Set the current parent index in context
            _context.SetCollectionIndex(parentName, parentIndex);

            // Get the child collection for this parent
            var childCollection = _context.GetNestedCollection($"{parentName}[{parentIndex}].{childName}");
            if (childCollection == null)
            {
                Logger.Warning($"Child collection not found: {parentName}[{parentIndex}].{childName}");
                return;
            }

            // Get the child collection count
            int childCount = CollectionHelper.GetCollectionCount(childCollection);
            if (childCount == 0)
            {
                Logger.Warning($"Child collection is empty: {parentName}[{parentIndex}].{childName}");
                return;
            }

            Logger.Debug($"Child collection {childName} has {childCount} items");

            // Create the combined variable name (Parent_Child)
            string combinedName = $"{parentName}_{childName}";

            // Ensure the combined variable exists in the context
            if (!_context.Variables.ContainsKey(combinedName))
            {
                _context.Variables[combinedName] = childCollection;
                Logger.Debug($"Added combined variable: {combinedName}");
            }

            // Create array batch parameters
            var batchParams = ArrayBatchParameters.CreateNestedCollection(
                new[] { parentName, childName },
                new[] { parentIndex },
                maxItemsPerSlide,
                offset);

            // Clone the slide for this parent's child collection
            var clonedSlidePart = SlideHelper.CloneSlide(presentationPart, slidePart);

            // Insert it after the original
            int originalPosition = SlideHelper.FindSlidePosition(presentationPart, slidePart);
            SlideHelper.InsertSlide(presentationPart, clonedSlidePart, originalPosition + 1 + parentIndex);

            Logger.Debug($"Created slide for {parentName}[{parentIndex}].{childName}");

            // Process the cloned slide with the array batch processor
            var processor = new ArrayBatchProcessor(_context, _context.Variables);
            var result = processor.ProcessArrayBatch(presentationPart, clonedSlidePart, batchParams);

            if (result.WasProcessed)
            {
                // Track this processed slide for later reference
                string processedId = $"{parentName}_{parentIndex}_{childName}_0";
                _context.ProcessedArraySlides.Add(processedId);

                // Track additional slides generated for this batch if any
                foreach (var generatedSlide in result.GeneratedSlides)
                {
                    if (generatedSlide != clonedSlidePart)
                    {
                        int batchIndex = result.GeneratedSlides.IndexOf(generatedSlide);
                        string batchId = $"{parentName}_{parentIndex}_{childName}_{batchIndex}";
                        _context.ProcessedArraySlides.Add(batchId);
                    }
                }

                Logger.Debug($"Successfully processed nested collection {parentName}[{parentIndex}].{childName}");
            }
            else
            {
                Logger.Warning($"Failed to process nested collection {parentName}[{parentIndex}].{childName}");
            }
        }
        catch (Exception ex)
        {
            Logger.Error($"Error processing nested collection item: {ex.Message}", ex);
        }
    }
}