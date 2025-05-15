namespace DocuChef.PowerPoint.Processing.ArrayProcessing;

/// <summary>
/// Extension to ArrayBatchProcessor for handling nested collection parameters
/// </summary>
internal partial class ArrayBatchParameters
{
    /// <summary>
    /// Gets the parent collection name (first part in a nested collection)
    /// </summary>
    public string GetParentCollectionName()
    {
        if (!IsNestedCollection)
            return CollectionName;

        return CollectionParts[0];
    }

    /// <summary>
    /// Gets the child collection name (second part in a nested collection)
    /// </summary>
    public string GetChildCollectionName()
    {
        if (!IsNestedCollection || CollectionParts.Length < 2)
            return null;

        return CollectionParts[1];
    }

    /// <summary>
    /// Gets the parent index (first index in a nested collection)
    /// </summary>
    public int GetParentIndex()
    {
        if (!IsNestedCollection || CollectionIndices.Length == 0)
            return 0;

        return CollectionIndices[0];
    }

    /// <summary>
    /// Create a combination of the parent and child collection names (Parent_Child)
    /// </summary>
    public string GetCombinedName()
    {
        if (!IsNestedCollection || CollectionParts.Length < 2)
            return CollectionName;

        return $"{CollectionParts[0]}_{CollectionParts[1]}";
    }
}