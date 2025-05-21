namespace DocuChef.Presentation.Processing;

internal partial class TemplateProcessor
{
    /// <summary>
    /// Processes a collection to determine its hierarchy
    /// </summary>
    private void ProcessCollectionHierarchy(
        string collection,
        List<ForeachSlideInfo> foreachSlides,
        HierarchyInfo result)
    {
        // Get the configured hierarchy delimiter
        string delimiter = PowerPointOptions.Current.HierarchyDelimiter;

        // Split the collection name using the configured delimiter
        var segments = collection.Split(new[] { delimiter }, StringSplitOptions.None);
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
    /// Adds parent-child relationship for a collection
    /// </summary>
    private void AddParentChildRelationship(
        string collection,
        string[] segments,
        HierarchyInfo result)
    {
        // Get the configured hierarchy delimiter
        string delimiter = PowerPointOptions.Current.HierarchyDelimiter;

        // Create the parent path by joining all segments except the last one
        var parentPath = string.Join(delimiter, segments.Take(segments.Length - 1));
        result.ParentCollections[collection] = parentPath;

        // Add this collection to the child collections of the parent
        if (!result.ChildCollections.ContainsKey(parentPath))
        {
            result.ChildCollections[parentPath] = new List<string>();
        }

        result.ChildCollections[parentPath].Add(collection);
    }
}