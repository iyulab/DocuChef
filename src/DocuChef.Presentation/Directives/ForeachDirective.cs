namespace DocuChef.Presentation.Directives;

/// <summary>
/// Represents a foreach directive for iterating over a collection
/// </summary>
internal class ForeachDirective : Directive
{
    /// <summary>
    /// The name of the collection to iterate over
    /// </summary>
    public string CollectionName { get; set; }

    /// <summary>
    /// The maximum number of items to process
    /// </summary>
    public int MaxItems { get; set; } = int.MaxValue;

    /// <summary>
    /// Gets the directive type
    /// </summary>
    public override DirectiveType Type => DirectiveType.Foreach;

    /// <summary>
    /// Evaluates if the collection exists and has items
    /// </summary>
    public override bool Evaluate(Models.SlideContext context)
    {
        // A foreach directive is considered valid if the collection exists and has items
        if (context == null || string.IsNullOrEmpty(CollectionName))
            return false;

        var collection = context.RootData.GetCollection(CollectionName);
        return collection != null && collection.Count() > 0;
    }

    /// <summary>
    /// Returns a string representation of this directive
    /// </summary>
    public override string ToString()
    {
        return $"Foreach: {CollectionName}, Max: {(MaxItems == int.MaxValue ? "All" : MaxItems.ToString())}";
    }
}