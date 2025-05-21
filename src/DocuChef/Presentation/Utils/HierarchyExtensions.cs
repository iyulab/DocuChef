namespace DocuChef.Presentation.Utils;

/// <summary>
/// Extension methods for working with collection hierarchies
/// </summary>
public static class HierarchyExtensions
{
    /// <summary>
    /// Creates a collection path by joining parent and child collection names
    /// using the configured hierarchy delimiter
    /// </summary>
    public static string CreateHierarchicalPath(this string parentCollection, string childCollection)
    {
        if (string.IsNullOrEmpty(parentCollection))
            return childCollection;

        if (string.IsNullOrEmpty(childCollection))
            return parentCollection;

        // Get the configured delimiter
        string delimiter = PowerPointOptions.Current.HierarchyDelimiter;

        // Join with the configured delimiter
        return $"{parentCollection}{delimiter}{childCollection}";
    }

    /// <summary>
    /// Splits a hierarchical collection path into its constituent segments
    /// using the configured hierarchy delimiter
    /// </summary>
    public static string[] SplitHierarchicalPath(this string collectionPath)
    {
        if (string.IsNullOrEmpty(collectionPath))
            return new string[0];

        // Get the configured delimiter
        string delimiter = PowerPointOptions.Current.HierarchyDelimiter;

        // Split using the configured delimiter
        return collectionPath.Split(new[] { delimiter }, StringSplitOptions.None);
    }

    /// <summary>
    /// Gets the parent collection path from a hierarchical collection path
    /// using the configured hierarchy delimiter
    /// </summary>
    public static string GetParentCollectionPath(this string collectionPath)
    {
        if (string.IsNullOrEmpty(collectionPath))
            return string.Empty;

        // Get the configured delimiter
        string delimiter = PowerPointOptions.Current.HierarchyDelimiter;

        // Find the last delimiter
        int lastDelimiterIndex = collectionPath.LastIndexOf(delimiter);

        if (lastDelimiterIndex > 0)
            return collectionPath.Substring(0, lastDelimiterIndex);

        return string.Empty;
    }

    /// <summary>
    /// Gets the last segment from a hierarchical collection path
    /// using the configured hierarchy delimiter
    /// </summary>
    public static string GetLastSegment(this string collectionPath)
    {
        if (string.IsNullOrEmpty(collectionPath))
            return string.Empty;

        // Get the configured delimiter
        string delimiter = PowerPointOptions.Current.HierarchyDelimiter;

        // Find the last delimiter
        int lastDelimiterIndex = collectionPath.LastIndexOf(delimiter);

        if (lastDelimiterIndex >= 0 && lastDelimiterIndex < collectionPath.Length - delimiter.Length)
            return collectionPath.Substring(lastDelimiterIndex + delimiter.Length);

        return collectionPath;
    }

    /// <summary>
    /// Checks if a collection path is a hierarchical path
    /// using the configured hierarchy delimiter
    /// </summary>
    public static bool IsHierarchicalPath(this string collectionPath)
    {
        if (string.IsNullOrEmpty(collectionPath))
            return false;

        // Get the configured delimiter
        string delimiter = PowerPointOptions.Current.HierarchyDelimiter;

        // Check if the path contains the delimiter
        return collectionPath.Contains(delimiter);
    }
}