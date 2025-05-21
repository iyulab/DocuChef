using DocuChef.Presentation.Models;

namespace DocuChef.Presentation.Processing;

/// <summary>
/// Contains information about collection hierarchy
/// </summary>
internal class HierarchyInfo
{
    /// <summary>
    /// Collections grouped by their nesting level (0 for top-level)
    /// </summary>
    public Dictionary<int, List<string>> CollectionsByLevel { get; set; } = new Dictionary<int, List<string>>();

    /// <summary>
    /// Maps collection name to its hierarchy level
    /// </summary>
    public Dictionary<string, int> CollectionLevel { get; set; } = new Dictionary<string, int>();

    /// <summary>
    /// Maps each collection to its slides in the template
    /// </summary>
    public Dictionary<string, List<SlideInfo>> SlidesByCollection { get; set; } = new Dictionary<string, List<SlideInfo>>();

    /// <summary>
    /// Maps each nested collection to its parent collection
    /// </summary>
    public Dictionary<string, string> ParentCollections { get; set; } = new Dictionary<string, string>();

    /// <summary>
    /// Maps each collection to its child collections
    /// </summary>
    public Dictionary<string, List<string>> ChildCollections { get; set; } = new Dictionary<string, List<string>>();

    /// <summary>
    /// Maps each collection to its max items per slide
    /// </summary>
    public Dictionary<string, int> MaxItemsPerSlide { get; set; } = new Dictionary<string, int>();
}