using DocuChef.Presentation.Core;
using DocuChef.Presentation.Models;

namespace DocuChef.Presentation.Processing;

/// <summary>
/// Processes and generates slide contexts for collections and nested hierarchies
/// </summary>
internal class ContextProcessor
{
    private readonly object _dataSource;

    /// <summary>
    /// Initializes a new instance of the ContextProcessor
    /// </summary>
    public ContextProcessor(object dataSource)
    {
        _dataSource = dataSource ?? throw new ArgumentNullException(nameof(dataSource));
    }

    /// <summary>
    /// Generates contexts for a top-level collection
    /// </summary>
    public List<SlideContext> GenerateContextsForCollection(string collectionName)
    {
        Logger.Debug($"Generating contexts for top-level collection: {collectionName}");
        var result = new List<SlideContext>();

        // Get the collection from data source
        IEnumerable collection = _dataSource.GetCollection(collectionName);
        if (collection == null)
        {
            Logger.Warning($"Collection '{collectionName}' not found in data source");
            return result;
        }

        // Count items in the collection
        int totalItems = collection.Count();
        Logger.Debug($"Collection '{collectionName}' has {totalItems} items");

        if (totalItems == 0)
            return result;

        // Generate a context for each item
        int offset = 0;
        foreach (var item in collection)
        {
            // Create slide context for the current item
            var slideContext = CreateContext(collectionName, offset, item, totalItems);
            result.Add(slideContext);
            offset++;
        }

        return result;
    }

    /// <summary>
    /// Creates a new slide context with the specified parameters
    /// </summary>
    private SlideContext CreateContext(
        string collectionName,
        int offset,
        object item,
        int totalItems,
        SlideContext parentContext = null)
    {
        // Default parameters for non-nested items
        int hierarchyLevel = 0;
        object[] ancestors = null;
        var levelOffsets = new Dictionary<int, int> { [0] = offset };

        // Check if this is a nested collection item
        if (item is NestedCollectionItem nestedItem)
        {
            // Extract hierarchy information
            hierarchyLevel = nestedItem.NestingLevel;

            // Create ancestors array without the item itself (last element)
            ancestors = nestedItem.Ancestry?.Length > 1
                ? nestedItem.Ancestry.Take(nestedItem.Ancestry.Length - 1).ToArray()
                : null;

            // The actual item is inside the nested wrapper
            item = nestedItem.Item;

            // Extract hierarchy level offsets if available
            if (nestedItem.HierarchyPath != null)
            {
                for (int i = 0; i < nestedItem.HierarchyPath.Length; i++)
                {
                    // Default all level offsets to 0, will be populated elsewhere
                    levelOffsets[i] = 0;
                }
            }
        }

        return new SlideContext
        {
            CollectionName = collectionName,
            Offset = offset,
            CurrentItem = item,
            RootData = _dataSource,
            TotalItems = totalItems,
            HierarchyLevel = hierarchyLevel,
            Ancestors = ancestors,
            LevelOffsets = levelOffsets,
            ParentContext = parentContext
        };
    }

    /// <summary>
    /// Generates contexts for a nested collection
    /// </summary>
    public List<SlideContext> GenerateNestedContexts(
        string nestedCollection,
        SlideContext parentContext,
        int maxItemsPerSlide = 0)
    {
        Logger.Debug($"Generating nested contexts for collection: {nestedCollection} based on parent context");
        var result = new List<SlideContext>();

        if (parentContext == null || parentContext.CurrentItem == null)
        {
            Logger.Warning("Parent context or current item is null");
            return result;
        }

        // Get the configured hierarchy delimiter
        string delimiter = PowerPointOptions.Current.HierarchyDelimiter;

        // Split collection name using the configured delimiter
        string[] segments = nestedCollection.Split(new[] { delimiter }, StringSplitOptions.None);
        if (segments.Length < 2)
        {
            Logger.Warning($"Invalid nested collection name: {nestedCollection}");
            return result;
        }

        string childPropertyName = segments[segments.Length - 1];

        // Get the child collection from the parent item
        var childCollection = GetNestedCollection(parentContext.CurrentItem, childPropertyName);
        if (childCollection == null)
        {
            Logger.Debug($"No child collection '{childPropertyName}' found in parent item");
            return result;
        }

        // Count child items
        int totalChildItems = childCollection.Count();
        if (totalChildItems == 0)
            return result;

        Logger.Debug($"Found {totalChildItems} items in nested collection {childPropertyName}");

        // Always use grouping if maxItemsPerSlide is specified, even for small collections
        if (maxItemsPerSlide > 0)
        {
            return GenerateGroupedNestedContexts(nestedCollection, parentContext, maxItemsPerSlide, childCollection, totalChildItems);
        }

        // Generate context for each child item individually if no grouping
        int childOffset = 0;
        foreach (var childItem in childCollection)
        {
            // Create ancestry chain
            var ancestors = parentContext.Ancestors != null
                ? parentContext.Ancestors.ToList()
                : new List<object>();

            // Add the parent item to the ancestry chain
            ancestors.Add(parentContext.CurrentItem);

            // Create level offsets dictionary
            var levelOffsets = new Dictionary<int, int>(parentContext.LevelOffsets);
            int childLevel = segments.Length - 1;
            levelOffsets[childLevel] = childOffset;

            // Create nested context
            var childContext = new SlideContext
            {
                CollectionName = nestedCollection,
                Offset = childOffset,
                CurrentItem = childItem,
                RootData = _dataSource,
                TotalItems = totalChildItems,
                HierarchyLevel = childLevel,
                Ancestors = ancestors.ToArray(),
                LevelOffsets = levelOffsets,
                ParentContext = parentContext
            };

            result.Add(childContext);
            childOffset++;
        }

        Logger.Debug($"Generated {result.Count} nested contexts for collection: {nestedCollection}");
        return result;
    }

    /// <summary>
    /// Gets a nested collection from a parent item using property name
    /// </summary>
    private IEnumerable GetNestedCollection(object parentItem, string childPropertyName)
    {
        try
        {
            // Try to get property by reflection first
            var parentType = parentItem.GetType();
            var childProperty = parentType.GetProperty(childPropertyName,
                System.Reflection.BindingFlags.Public |
                System.Reflection.BindingFlags.Instance |
                System.Reflection.BindingFlags.IgnoreCase);

            if (childProperty != null)
            {
                var childValue = childProperty.GetValue(parentItem);
                if (childValue is IEnumerable enumerable && !(childValue is string))
                {
                    Logger.Debug($"Found child collection '{childPropertyName}' using reflection from parent item of type {parentType.Name}");
                    return enumerable;
                }
            }

            // Try using GetCollection extension method
            var collection = parentItem.GetCollection(childPropertyName);
            if (collection != null)
            {
                Logger.Debug($"Found child collection '{childPropertyName}' using GetCollection method");
                return collection;
            }

            // Try to access dictionary
            if (parentItem is IDictionary<string, object> dict && dict.TryGetValue(childPropertyName, out var dictValue))
            {
                if (dictValue is IEnumerable enumerable && !(dictValue is string))
                {
                    Logger.Debug($"Found child collection '{childPropertyName}' in dictionary");
                    return enumerable;
                }
            }

            // Try using dynamic if everything else fails
            try
            {
                dynamic dynamicParent = parentItem;
                var dynamicResult = dynamicParent[childPropertyName];

                if (dynamicResult is IEnumerable enumerable && !(dynamicResult is string))
                {
                    Logger.Debug($"Found child collection '{childPropertyName}' using dynamic access");
                    return enumerable;
                }
            }
            catch
            {
                // Dynamic access failed, continue with other methods
            }

            Logger.Warning($"Could not find child collection '{childPropertyName}' in parent item of type {parentItem.GetType().Name}");
            return null;
        }
        catch (Exception ex)
        {
            Logger.Error($"Error accessing nested collection '{childPropertyName}': {ex.Message}", ex);
            return null;
        }
    }

    /// <summary>
    /// Generates nested contexts with grouping based on max items per slide
    /// </summary>
    public List<SlideContext> GenerateGroupedNestedContexts(
        string nestedCollection,
        SlideContext parentContext,
        int maxItemsPerSlide,
        IEnumerable childItems = null,
        int? totalItemsCount = null)
    {
        Logger.Debug($"Generating grouped nested contexts for collection: {nestedCollection}, max items: {maxItemsPerSlide}");
        var result = new List<SlideContext>();

        // Get the configured hierarchy delimiter
        string delimiter = PowerPointOptions.Current.HierarchyDelimiter;

        // Extract child property name using the configured delimiter
        string[] segments = nestedCollection.Split(new[] { delimiter }, StringSplitOptions.None);
        if (segments.Length < 2)
            return result;

        string childPropertyName = segments[segments.Length - 1];

        // Get child items from parent context if not provided
        if (childItems == null)
        {
            childItems = GetNestedCollection(parentContext.CurrentItem, childPropertyName);
            if (childItems == null)
                return result;
        }

        // Count total items if not provided
        int totalItems = totalItemsCount ?? childItems.Count();
        if (totalItems == 0)
            return result;

        // Group items by max items per slide
        var itemsList = childItems.Cast<object>().ToList();
        int groupCount = CalculateGroupCount(totalItems, maxItemsPerSlide);

        Logger.Debug($"Grouping {totalItems} items into {groupCount} groups with max {maxItemsPerSlide} items per group");
        Logger.Debug($"Parent context: {parentContext.GetContextDescription()}");

        for (int groupIndex = 0; groupIndex < groupCount; groupIndex++)
        {
            // Get items for this group
            int startIndex = groupIndex * maxItemsPerSlide;
            int itemsToTake = Math.Min(maxItemsPerSlide, totalItems - startIndex);
            var groupItems = itemsList.Skip(startIndex).Take(itemsToTake).ToList();

            // Create ancestry chain
            var ancestors = parentContext.Ancestors != null
                ? parentContext.Ancestors.ToList()
                : new List<object>();

            // Add the parent item to the ancestry chain
            ancestors.Add(parentContext.CurrentItem);

            // Create level offsets dictionary - copy from parent first
            var levelOffsets = new Dictionary<int, int>(parentContext.LevelOffsets);
            int childLevel = segments.Length - 1;

            // Set correct offset for this group based on its position in the collection
            levelOffsets[childLevel] = startIndex;

            // Create context for this group
            var groupContext = new SlideContext
            {
                CollectionName = nestedCollection,
                Offset = startIndex, // Correct offset for each group
                CurrentItem = groupItems,
                RootData = _dataSource,
                TotalItems = totalItems,
                HierarchyLevel = childLevel,
                Ancestors = ancestors.ToArray(),
                LevelOffsets = levelOffsets,
                ParentContext = parentContext,
                ItemsInGroup = groupItems.Count
            };

            result.Add(groupContext);

            Logger.Debug($"Created nested group context: Collection={nestedCollection}, Offset={startIndex}, Items={groupItems.Count}, ParentContext={parentContext.CollectionName}[{parentContext.Offset}]");
        }

        return result;
    }

    /// <summary>
    /// Calculates the number of groups needed based on total items and max items per group
    /// </summary>
    private int CalculateGroupCount(int totalItems, int maxItemsPerGroup)
    {
        if (maxItemsPerGroup <= 0)
            maxItemsPerGroup = int.MaxValue;

        return (int)Math.Ceiling((double)totalItems / maxItemsPerGroup);
    }
}