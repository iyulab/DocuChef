namespace DocuChef.Presentation.Models;

/// <summary>
/// Plan containing all slide instances that need to be generated
/// </summary>
public class SlidePlan
{
    /// <summary>
    /// List of slide instances to be created
    /// </summary>
    public List<SlideInstance> SlideInstances { get; set; } = new List<SlideInstance>();
    
    /// <summary>
    /// Alias mappings from alias names to full paths
    /// </summary>
    public Dictionary<string, string> Aliases { get; set; } = new Dictionary<string, string>();
    
    /// <summary>
    /// Context chains for nested collections
    /// </summary>
    public Dictionary<string, List<string>> ContextChains { get; set; } = new Dictionary<string, List<string>>();
    
    /// <summary>
    /// Total number of slides in the final presentation
    /// </summary>
    public int TotalSlideCount => SlideInstances.Count;
    
    /// <summary>
    /// Add a slide instance to the plan
    /// </summary>
    public void AddSlideInstance(SlideInstance instance)
    {
        SlideInstances.Add(instance);
    }
    
    /// <summary>
    /// Get slide instances by source slide ID
    /// </summary>
    public List<SlideInstance> GetInstancesBySourceSlideId(int sourceSlideId)
    {
        return SlideInstances.Where(s => s.SourceSlideId == sourceSlideId).ToList();
    }
}

/// <summary>
/// Instance of a slide to be generated, with context and positioning information
/// </summary>
public class SlideInstance
{
    /// <summary>
    /// ID of the source slide to clone from
    /// </summary>
    public int SourceSlideId { get; set; }
    
    /// <summary>
    /// Type of this slide instance
    /// </summary>
    public SlideInstanceType Type { get; set; }
    
    /// <summary>
    /// Position where this slide should be placed in final presentation
    /// </summary>
    public int Position { get; set; }
    
    /// <summary>
    /// Context information for data binding
    /// </summary>
    public List<string> ContextPath { get; set; } = new List<string>();
      /// <summary>
    /// Index offset for array references in this slide
    /// </summary>
    public int IndexOffset { get; set; }
    
    /// <summary>
    /// Collection name this instance processes
    /// </summary>
    public string? CollectionName { get; set; }
    
    /// <summary>
    /// Starting index for this batch
    /// </summary>
    public int StartIndex { get; set; }
    
    /// <summary>
    /// Number of items this instance should display
    /// </summary>
    public int ItemsPerSlide { get; set; }
    
    /// <summary>
    /// Whether this instance represents empty data
    /// </summary>
    public bool IsEmpty { get; set; }
    
    /// <summary>
    /// Parent index for nested collections (used for Products>Items scenarios)
    /// </summary>
    public int? ParentIndex { get; set; }
    
    /// <summary>
    /// Context path as a joined string
    /// </summary>
    public string ContextPathString => string.Join(">", ContextPath);
}

/// <summary>
/// Types of slide instances
/// </summary>
public enum SlideInstanceType
{
    /// <summary>
    /// Static slide with no data binding
    /// </summary>
    Static,
    
    /// <summary>
    /// Generated slide from template processing
    /// </summary>
    Generated
}