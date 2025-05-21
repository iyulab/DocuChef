using System.Dynamic;
using System.Reflection;

namespace DocuChef.Presentation.Models;

public partial class SlideContext
{
    /// <summary>
    /// Gets data object for binding with this context
    /// </summary>
    /// <returns>Data object with properties from the context</returns>
    internal object? GetData()
    {
        try
        {
            // Create a dynamic object to store context data
            var expandoObj = new ExpandoObject() as IDictionary<string, object>;

            // Add all basic context properties
            expandoObj["CollectionName"] = CollectionName;
            expandoObj["Offset"] = Offset;
            expandoObj["TotalItems"] = TotalItems;
            expandoObj["ItemNumber"] = ItemNumber;
            expandoObj["LastItemNumber"] = LastItemNumber;
            expandoObj["IsFirst"] = IsFirst;
            expandoObj["IsLast"] = IsLast;
            expandoObj["HierarchyLevel"] = HierarchyLevel;

            // Process hierarchical path if present
            string delimiter = PowerPointOptions.Current.HierarchyDelimiter;
            if (CollectionName.Contains(delimiter))
            {
                // Split hierarchical path into segments
                string[] segments = CollectionName.Split(
                    new[] { delimiter },
                    StringSplitOptions.None);

                // Add all hierarchy segments for direct access
                for (int i = 0; i < segments.Length; i++)
                {
                    // Process parent (top-level) collection if this is a nested collection
                    if (i == 0 && segments.Length > 1)
                    {
                        string parentName = segments[0];

                        // Try to get parent data from parent context
                        if (ParentContext != null)
                        {
                            // First check if we can get the parent item directly
                            if (ParentContext.CurrentItem != null)
                            {
                                expandoObj[parentName] = ParentContext.CurrentItem;

                                // Also add all parent's properties directly
                                AddObjectProperties(expandoObj, ParentContext.CurrentItem);
                            }
                        }
                        else if (Ancestors != null && Ancestors.Length > 0)
                        {
                            // Use the last ancestor as parent if no parent context
                            expandoObj[parentName] = Ancestors[Ancestors.Length - 1];
                        }
                    }

                    // Process current (last) segment - the collection or nested collection
                    if (i == segments.Length - 1)
                    {
                        string currentName = segments[i];

                        // Add current item with its segment name
                        expandoObj[currentName] = CurrentItem;

                        // For groups, add special Item[] array access
                        if (IsGroupContext && CurrentItem is IEnumerable<object> groupItems)
                        {
                            // Add special "Item[index]" access for group items
                            var itemArray = groupItems.ToArray();
                            for (int j = 0; j < itemArray.Length; j++)
                            {
                                expandoObj[$"Item[{j}]"] = itemArray[j];

                                // Also add properties from each item in the group
                                AddObjectProperties(expandoObj, itemArray[j], $"Item[{j}].");
                            }
                        }
                        else
                        {
                            // For single items, add all properties directly
                            AddObjectProperties(expandoObj, CurrentItem);
                        }
                    }
                }

                // Add special combined name for 2-level hierarchies
                if (segments.Length == 2)
                {
                    string combinedName = segments[0] + "__" + segments[1];
                    expandoObj[combinedName] = CurrentItem;
                }
            }
            else
            {
                // Non-hierarchical collection, just add current item properties
                expandoObj[CollectionName] = CurrentItem;

                // Add direct properties from current item
                AddObjectProperties(expandoObj, CurrentItem);
            }

            // Add root data for global access (if not already in context)
            if (RootData != null)
            {
                // Add properties from root data without overriding context properties
                AddObjectProperties(expandoObj, RootData, overwrite: false);
            }

            // Add parent context data if available
            if (ParentContext != null)
            {
                // Get parent context data
                var parentData = ParentContext.GetData();

                // If parent data is a dictionary, merge its properties without overriding
                if (parentData is IDictionary<string, object> parentDict)
                {
                    foreach (var key in parentDict.Keys)
                    {
                        if (!expandoObj.ContainsKey(key))
                        {
                            expandoObj[key] = parentDict[key];
                        }
                    }
                }
            }

            return expandoObj;
        }
        catch (Exception ex)
        {
            Logger.Warning($"Error creating context data: {ex.Message}");

            // Return empty object as fallback
            return new ExpandoObject();
        }
    }

    /// <summary>
    /// Adds properties from a source object to a dictionary
    /// </summary>
    private void AddObjectProperties(IDictionary<string, object> target, object source, string prefix = "", bool overwrite = true)
    {
        if (source == null)
            return;

        try
        {
            // If it's already a dictionary, add its entries directly
            if (source is IDictionary<string, object> dict)
            {
                foreach (var key in dict.Keys)
                {
                    string propName = prefix + key;
                    if (overwrite || !target.ContainsKey(propName))
                    {
                        target[propName] = dict[key];
                    }
                }
                return;
            }

            // For regular objects, use reflection to get properties
            var properties = source.GetType().GetProperties(
                BindingFlags.Public | BindingFlags.Instance);

            foreach (var prop in properties)
            {
                try
                {
                    if (prop.CanRead)
                    {
                        string propName = prefix + prop.Name;
                        if (overwrite || !target.ContainsKey(propName))
                        {
                            var value = prop.GetValue(source);
                            target[propName] = value;
                        }
                    }
                }
                catch
                {
                    // Skip properties that throw exceptions
                }
            }
        }
        catch (Exception ex)
        {
            Logger.Debug($"Error adding object properties: {ex.Message}");
        }
    }
}