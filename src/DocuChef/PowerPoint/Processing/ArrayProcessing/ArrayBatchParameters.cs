namespace DocuChef.PowerPoint.Processing.ArrayProcessing;

/// <summary>
/// Parameters for array batch processing
/// </summary>
internal partial class ArrayBatchParameters
{
    /// <summary>
    /// Name of the collection/array to process
    /// </summary>
    public string CollectionName { get; set; }

    /// <summary>
    /// Maximum items per slide (defaults to auto-detect)
    /// </summary>
    public int MaxItemsPerSlide { get; set; } = -1;

    /// <summary>
    /// Starting offset in the collection (defaults to 0)
    /// </summary>
    public int Offset { get; set; } = 0;

    /// <summary>
    /// Whether parameters were explicitly specified via directive (vs auto-detected)
    /// </summary>
    public bool IsExplicitlySpecified { get; set; }

    /// <summary>
    /// Collection parts for hierarchical path (e.g., "Parent", "Sub", "Child")
    /// </summary>
    public string[] CollectionParts { get; set; }

    /// <summary>
    /// Collection indices for each level in the hierarchy except the deepest level
    /// </summary>
    public int[] CollectionIndices { get; set; }

    /// <summary>
    /// Whether this is a nested collection
    /// </summary>
    public bool IsNestedCollection => CollectionParts != null && CollectionParts.Length > 1;

    /// <summary>
    /// Creates a new instance with auto-detection (no explicit parameters)
    /// </summary>
    public static ArrayBatchParameters CreateAutoDetect(string collectionName)
    {
        // Parse collection parts for nested collections
        var parts = collectionName.Split('_');

        return new ArrayBatchParameters
        {
            CollectionName = collectionName,
            CollectionParts = parts.Length > 1 ? parts : new[] { collectionName },
            CollectionIndices = parts.Length > 1 ? new int[parts.Length - 1] : Array.Empty<int>(),
            MaxItemsPerSlide = -1, // Auto-detect
            Offset = 0,
            IsExplicitlySpecified = false
        };
    }

    /// <summary>
    /// Creates a new instance with explicitly specified parameters
    /// </summary>
    public static ArrayBatchParameters CreateExplicit(string collectionName, int maxItemsPerSlide = -1, int offset = 0)
    {
        // Parse collection parts for nested collections
        var parts = collectionName.Split('_');

        return new ArrayBatchParameters
        {
            CollectionName = collectionName,
            CollectionParts = parts.Length > 1 ? parts : new[] { collectionName },
            CollectionIndices = parts.Length > 1 ? new int[parts.Length - 1] : Array.Empty<int>(),
            MaxItemsPerSlide = maxItemsPerSlide,
            Offset = offset,
            IsExplicitlySpecified = true
        };
    }

    /// <summary>
    /// Creates parameters for a nested collection with multiple levels
    /// </summary>
    public static ArrayBatchParameters CreateNestedCollection(
        string[] collectionParts,
        int[] collectionIndices,
        int maxItemsPerSlide = -1,
        int offset = 0)
    {
        if (collectionParts == null || collectionParts.Length == 0)
            throw new ArgumentException("Collection parts cannot be null or empty", nameof(collectionParts));

        if (collectionIndices == null)
            throw new ArgumentNullException(nameof(collectionIndices));

        if (collectionParts.Length <= 1)
            throw new ArgumentException("Collection parts must contain at least two parts for nesting", nameof(collectionParts));

        if (collectionIndices.Length != collectionParts.Length - 1)
            throw new ArgumentException("Collection indices must have exactly one fewer element than collection parts", nameof(collectionIndices));

        return new ArrayBatchParameters
        {
            CollectionName = string.Join("_", collectionParts),
            CollectionParts = collectionParts,
            CollectionIndices = collectionIndices,
            MaxItemsPerSlide = maxItemsPerSlide,
            Offset = offset,
            IsExplicitlySpecified = true
        };
    }
}