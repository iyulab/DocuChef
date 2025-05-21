using DocuChef.Presentation.Core;

namespace DocuChef.Presentation.Models;

public partial class SlideContext
{
    /// <summary>
    /// Processes array notation like Categories[0], Categories>Products[1], etc.
    /// </summary>
    private string ProcessArrayNotation(string propertyPath)
    {
        int openBracket = propertyPath.IndexOf('[');
        int closeBracket = propertyPath.IndexOf(']', openBracket);

        if (openBracket <= 0 || closeBracket <= openBracket)
            return null;

        string collectionName = propertyPath.Substring(0, openBracket);
        string indexStr = propertyPath.Substring(openBracket + 1, closeBracket - openBracket - 1);

        if (!int.TryParse(indexStr, out int index))
            return null;

        // Get remaining property path after the bracket if any
        string remainingPath = "";
        if (closeBracket + 1 < propertyPath.Length && propertyPath[closeBracket + 1] == '.')
        {
            remainingPath = propertyPath.Substring(closeBracket + 2);
        }

        Logger.Debug($"Processing array notation: collection={collectionName}, index={index}, remainingPath={remainingPath}");

        // Check if current context matches the collection and index
        if (CollectionName.Equals(collectionName, StringComparison.OrdinalIgnoreCase) && Offset == index)
        {
            if (string.IsNullOrEmpty(remainingPath))
                return CurrentItem?.ToString() ?? string.Empty;
            else
                return ObjectExtensions.GetPropertyValue(CurrentItem, remainingPath);
        }

        // Check if parent context matches
        if (ParentContext != null &&
            ParentContext.CollectionName.Equals(collectionName, StringComparison.OrdinalIgnoreCase) &&
            ParentContext.Offset == index)
        {
            if (string.IsNullOrEmpty(remainingPath))
                return ParentContext.CurrentItem?.ToString() ?? string.Empty;
            else
                return ObjectExtensions.GetPropertyValue(ParentContext.CurrentItem, remainingPath);
        }

        // Check ancestors (for deeply nested hierarchies)
        if (Ancestors != null && HierarchyLevel > 0)
        {
            for (int i = 0; i < Ancestors.Length; i++)
            {
                // Since we don't have collection names for ancestors directly, 
                // we need to check parent context's parent contexts
                var currentParentContext = ParentContext;
                int level = 1;

                while (currentParentContext != null)
                {
                    if (currentParentContext.CollectionName.Equals(collectionName, StringComparison.OrdinalIgnoreCase) &&
                        currentParentContext.Offset == index)
                    {
                        if (string.IsNullOrEmpty(remainingPath))
                            return currentParentContext.CurrentItem?.ToString() ?? string.Empty;
                        else
                            return ObjectExtensions.GetPropertyValue(currentParentContext.CurrentItem, remainingPath);
                    }

                    currentParentContext = currentParentContext.ParentContext;
                    level++;

                    if (level > 10) break; // Safety limit to prevent infinite loops
                }
            }
        }

        return null;
    }

    /// <summary>
    /// Gets the collection and hierarchical segment names from a collection path
    /// </summary>
    public string[] GetHierarchySegments()
    {
        if (string.IsNullOrEmpty(CollectionName))
            return new string[0];

        // Get the hierarchy delimiter from options
        string delimiter = PowerPointOptions.Current.HierarchyDelimiter;

        // Split using the configured delimiter
        return CollectionName.Split(new[] { delimiter }, StringSplitOptions.None);
    }

    /// <summary>
    /// Checks if the current context represents a nested collection
    /// </summary>
    public bool IsNestedCollection()
    {
        string delimiter = PowerPointOptions.Current.HierarchyDelimiter;
        return CollectionName.Contains(delimiter);
    }

    /// <summary>
    /// Gets the parent collection name for a nested collection
    /// </summary>
    public string GetParentCollectionName()
    {
        if (!IsNestedCollection())
            return string.Empty;

        string delimiter = PowerPointOptions.Current.HierarchyDelimiter;
        int lastDelimiterIndex = CollectionName.LastIndexOf(delimiter);

        if (lastDelimiterIndex > 0)
            return CollectionName.Substring(0, lastDelimiterIndex);

        return string.Empty;
    }

    /// <summary>
    /// Gets the last segment name for a nested collection
    /// </summary>
    public string GetLastSegmentName()
    {
        if (!IsNestedCollection())
            return CollectionName;

        string delimiter = PowerPointOptions.Current.HierarchyDelimiter;
        int lastDelimiterIndex = CollectionName.LastIndexOf(delimiter);

        if (lastDelimiterIndex >= 0 && lastDelimiterIndex < CollectionName.Length - 1)
            return CollectionName.Substring(lastDelimiterIndex + delimiter.Length);

        return CollectionName;
    }
}