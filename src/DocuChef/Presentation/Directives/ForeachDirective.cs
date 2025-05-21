using DocuChef.Presentation.Core;

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
    /// The maximum number of items to display per slide
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

    /// <summary>
    /// Determines if this directive should use grouped items mode
    /// </summary>
    public bool ShouldUseGroupedMode()
    {
        // If MaxItems is specified and greater than 1, use grouped mode
        return MaxItems > 1 && MaxItems < int.MaxValue;
    }

    /// <summary>
    /// Processes the slides that should be included in the foreach section
    /// (this slide and potential following slides)
    /// </summary>
    public List<int> GetSectionSlides(List<Models.SlideInfo> slides, int currentSlideIndex)
    {
        var result = new List<int> { currentSlideIndex };

        // Find subsequent slides to include until the next source slide with foreach directive
        for (int i = currentSlideIndex + 1; i < slides.Count; i++)
        {
            // Stop if we hit another slide with foreach directive
            if (slides[i].DirectiveType == DirectiveType.Foreach)
            {
                break;
            }

            // Add this slide to the section
            result.Add(i);
        }

        return result;
    }
}