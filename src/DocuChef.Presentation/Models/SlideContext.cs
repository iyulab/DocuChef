using DocumentFormat.OpenXml.Office2010.CustomUI;
using System.Reflection;

namespace DocuChef.Presentation.Models;

/// <summary>
/// Manages slide context information during processing
/// </summary>
public class SlideContext
{
    /// <summary>
    /// Current collection name being processed
    /// </summary>
    public string CollectionName { get; set; }

    /// <summary>
    /// Current item index within the collection (0-based)
    /// </summary>
    public int Offset { get; set; }

    /// <summary>
    /// Current collection item (single item or group of items)
    /// </summary>
    public object CurrentItem { get; set; }

    /// <summary>
    /// Reference to the root data object
    /// </summary>
    public object RootData { get; set; }

    /// <summary>
    /// Total number of items in the collection
    /// </summary>
    public int TotalItems { get; set; }

    /// <summary>
    /// Hierarchy level for nested collections (0 for top level)
    /// </summary>
    public int HierarchyLevel { get; set; }

    /// <summary>
    /// Array of ancestor items in the hierarchy (if nested)
    /// </summary>
    public object[] Ancestors { get; set; }

    /// <summary>
    /// Dictionary of level-specific offsets for nested hierarchies
    /// </summary>
    public Dictionary<int, int> LevelOffsets { get; set; }

    /// <summary>
    /// Parent context for nested collections
    /// </summary>
    public SlideContext ParentContext { get; set; }

    /// <summary>
    /// Number of items in a group (for grouped contexts)
    /// </summary>
    public int ItemsInGroup { get; set; } = 1;

    /// <summary>
    /// Whether this context contains a group of items
    /// </summary>
    public bool IsGroupContext => CurrentItem is IList || ItemsInGroup > 1;

    /// <summary>
    /// Whether the current item is the first in the collection
    /// </summary>
    public bool IsFirst => Offset == 0;

    /// <summary>
    /// Whether the current item is the last in the collection
    /// </summary>
    public bool IsLast => Offset + ItemsInGroup >= TotalItems;

    /// <summary>
    /// Whether the current item has an even index (0, 2, 4, ...)
    /// </summary>
    public bool IsEven => Offset % 2 == 0;

    /// <summary>
    /// Whether the current item has an odd index (1, 3, 5, ...)
    /// </summary>
    public bool IsOdd => !IsEven;

    /// <summary>
    /// The 1-based index of the current item (for display purposes)
    /// </summary>
    public int ItemNumber => Offset + 1;

    /// <summary>
    /// The 1-based index of the last item in this group
    /// </summary>
    public int LastItemNumber => Offset + ItemsInGroup;

    /// <summary>
    /// Gets the parent item, if available
    /// </summary>
    public object Parent => ParentContext?.CurrentItem ??
        (Ancestors != null && Ancestors.Length > 0 ? Ancestors[Ancestors.Length - 1] : null);

    /// <summary>
    /// Gets an item from the group by index (0-based)
    /// </summary>
    public object GetGroupItem(int index)
    {
        if (!IsGroupContext || CurrentItem == null)
            return null;

        if (CurrentItem is IList list && index >= 0 && index < list.Count)
            return list[index];

        return null;
    }

    /// <summary>
    /// Creates a context description for slide notes or logging
    /// </summary>
    public string GetContextDescription()
    {
        string itemInfo;

        if (IsGroupContext)
        {
            // For group contexts, show the number of items
            itemInfo = $"Group with {ItemsInGroup} items";
        }
        else
        {
            // For single item contexts, show the item description
            itemInfo = CurrentItem?.ToString() ?? "null";

            // Truncate long item descriptions
            if (itemInfo.Length > 40)
            {
                itemInfo = itemInfo.Substring(0, 37) + "...";
            }
        }

        // Add hierarchy level information for nested collections
        string levelInfo = HierarchyLevel > 0 ? $", level: {HierarchyLevel}" : "";

        // Add parent context info if available
        string parentInfo = "";
        if (ParentContext != null)
        {
            string parentCollection = ParentContext.CollectionName;
            int parentOffset = ParentContext.Offset;
            parentInfo = $", parent: [{parentCollection}:{parentOffset}]";
        }

        // Add group info if applicable
        string groupInfo = ItemsInGroup > 1 ? $", items: {ItemsInGroup}" : "";

        // Add range info for grouped contexts
        string rangeInfo = "";
        if (IsGroupContext && TotalItems > 0)
        {
            rangeInfo = $", range: {ItemNumber}-{LastItemNumber} of {TotalItems}";
        }

        return $"SlideContext - {CollectionName}, offset: {Offset}{levelInfo}{parentInfo}{groupInfo}{rangeInfo}, item: {itemInfo}";
    }

    /// <summary>
    /// Creates a new slide context with the specified parameters
    /// </summary>
    public static SlideContext Create(string collectionName, int offset, object item, object rootData, int totalItems)
    {
        // Default parameters for non-nested items
        int hierarchyLevel = 0;
        object[] ancestors = null;
        var levelOffsets = new Dictionary<int, int>();
        levelOffsets[0] = offset;

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
            RootData = rootData,
            TotalItems = totalItems,
            HierarchyLevel = hierarchyLevel,
            Ancestors = ancestors,
            LevelOffsets = levelOffsets
        };
    }

    /// <summary>
    /// Gets a property value from the context
    /// </summary>
    public string GetContextValue(string propertyPath)
    {
        // Handle null or empty path
        if (string.IsNullOrEmpty(propertyPath))
            return string.Empty;

        Logger.Debug($"Getting context value for path: {propertyPath}");

        // Special property evaluation based on context
        if (propertyPath.Equals("IsFirst", StringComparison.OrdinalIgnoreCase))
            return IsFirst.ToString();
        if (propertyPath.Equals("IsLast", StringComparison.OrdinalIgnoreCase))
            return IsLast.ToString();
        if (propertyPath.Equals("IsEven", StringComparison.OrdinalIgnoreCase))
            return IsEven.ToString();
        if (propertyPath.Equals("IsOdd", StringComparison.OrdinalIgnoreCase))
            return IsOdd.ToString();
        if (propertyPath.Equals("ItemNumber", StringComparison.OrdinalIgnoreCase))
            return ItemNumber.ToString();
        if (propertyPath.Equals("LastItemNumber", StringComparison.OrdinalIgnoreCase))
            return LastItemNumber.ToString();
        if (propertyPath.Equals("Offset", StringComparison.OrdinalIgnoreCase))
            return Offset.ToString();
        if (propertyPath.Equals("CollectionName", StringComparison.OrdinalIgnoreCase))
            return CollectionName;
        if (propertyPath.Equals("TotalItems", StringComparison.OrdinalIgnoreCase))
            return TotalItems.ToString();
        if (propertyPath.Equals("ItemsInGroup", StringComparison.OrdinalIgnoreCase))
            return ItemsInGroup.ToString();
        if (propertyPath.Equals("Level", StringComparison.OrdinalIgnoreCase) ||
            propertyPath.Equals("HierarchyLevel", StringComparison.OrdinalIgnoreCase))
            return HierarchyLevel.ToString();

        // Collection-specific array index notation: Categories[0], Categories_Products[1], etc.
        if (propertyPath.Contains("[") && propertyPath.Contains("]"))
        {
            string result = ProcessArrayNotation(propertyPath);
            if (result != null)
                return result;
        }

        // Group item access with syntax: Item[0], Item[1], etc.
        if (TryGetGroupItemValue(propertyPath, out string groupItemValue))
            return groupItemValue;

        // Multi-level ancestor access with syntax: Ancestor[0], Ancestor[1], etc.
        if (TryGetAncestorValue(propertyPath, out string ancestorValue))
            return ancestorValue;

        // Parent access
        if (TryGetParentValue(propertyPath, out string parentValue))
            return parentValue;

        // Look up property in current item
        if (CurrentItem != null && !IsGroupContext)
        {
            string itemValue = ObjectExtensions.GetPropertyValue(CurrentItem, propertyPath);
            if (!string.IsNullOrEmpty(itemValue))
            {
                Logger.Debug($"Found value in current item: {itemValue}");
                return itemValue;
            }
        }

        // Try parent context if available
        if (ParentContext != null)
        {
            string parentContextValue = ParentContext.GetContextValue(propertyPath);
            if (!string.IsNullOrEmpty(parentContextValue))
            {
                Logger.Debug($"Found value in parent context: {parentContextValue}");
                return parentContextValue;
            }
        }

        // Fall back to root data
        if (RootData != null)
        {
            string rootValue = ObjectExtensions.GetPropertyValue(RootData, propertyPath);
            if (!string.IsNullOrEmpty(rootValue))
            {
                Logger.Debug($"Found value in root data: {rootValue}");
                return rootValue;
            }
        }

        Logger.Debug($"No value found for path: {propertyPath}");
        return string.Empty;
    }

    /// <summary>
    /// Tries to get a value from a group item
    /// </summary>
    private bool TryGetGroupItemValue(string propertyPath, out string value)
    {
        value = string.Empty;

        // Group item access with syntax: Item[0], Item[1], etc.
        if (IsGroupContext && propertyPath.StartsWith("Item[", StringComparison.OrdinalIgnoreCase) && propertyPath.EndsWith("]"))
        {
            // Extract the index
            string indexStr = propertyPath.Substring(5, propertyPath.Length - 6);
            if (int.TryParse(indexStr, out int index))
            {
                // Get item from group
                var item = GetGroupItem(index);
                value = item?.ToString() ?? string.Empty;
                return true;
            }
        }

        // Group item property access with syntax: Item[0].PropertyName
        if (IsGroupContext && propertyPath.Contains(".", StringComparison.OrdinalIgnoreCase) &&
            propertyPath.StartsWith("Item[", StringComparison.OrdinalIgnoreCase))
        {
            int closeBracketIndex = propertyPath.IndexOf("]");
            if (closeBracketIndex > 0)
            {
                string indexStr = propertyPath.Substring(5, closeBracketIndex - 5);
                string itemPropertyPath = propertyPath.Substring(closeBracketIndex + 2); // +2 to skip ].

                if (int.TryParse(indexStr, out int index))
                {
                    // Get item from group
                    var item = GetGroupItem(index);
                    if (item != null)
                    {
                        value = ObjectExtensions.GetPropertyValue(item, itemPropertyPath);
                        return true;
                    }
                }
            }
        }

        return false;
    }

    /// <summary>
    /// Tries to get a value from an ancestor
    /// </summary>
    private bool TryGetAncestorValue(string propertyPath, out string value)
    {
        value = string.Empty;

        // Multi-level ancestor access with syntax: Ancestor[0], Ancestor[1], etc.
        if (propertyPath.StartsWith("Ancestor[", StringComparison.OrdinalIgnoreCase) && propertyPath.EndsWith("]"))
        {
            // Extract the index
            string indexStr = propertyPath.Substring(9, propertyPath.Length - 10);
            if (int.TryParse(indexStr, out int index) && Ancestors != null && index >= 0 && index < Ancestors.Length)
            {
                value = Ancestors[index]?.ToString() ?? string.Empty;
                return true;
            }
        }

        return false;
    }

    /// <summary>
    /// Tries to get a value from parent
    /// </summary>
    private bool TryGetParentValue(string propertyPath, out string value)
    {
        value = string.Empty;

        // Parent access for direct parent
        if (propertyPath.Equals("Parent", StringComparison.OrdinalIgnoreCase))
        {
            value = Parent?.ToString() ?? string.Empty;
            return true;
        }

        // Handle Parent. prefix for nested collections
        if (propertyPath.StartsWith("Parent.", StringComparison.OrdinalIgnoreCase) && Parent != null)
        {
            string parentPropertyPath = propertyPath.Substring(7); // Remove "Parent."
            value = ObjectExtensions.GetPropertyValue(Parent, parentPropertyPath);
            return !string.IsNullOrEmpty(value);
        }

        return false;
    }

    /// <summary>
    /// Processes array notation like Categories[0], Categories_Products[1], etc.
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

    internal object? GetData()
    {
        var splits = CollectionName.Split("_");
        if (this.Parent != null && splits.Length > 1)
        {
            var last = splits.Last();
            if (this.Parent.GetType().GetProperty(last) is PropertyInfo pInfo)
            {
                var value = pInfo.GetValue(this.Parent);
                return new Dictionary<string, object?>()
                {
                    { CollectionName, value }
                };
            }
        }

        return null;
    }
}