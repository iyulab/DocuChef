namespace DocuChef.PowerPoint;

/// <summary>
/// Context for PowerPoint template processing
/// </summary>
public partial class PowerPointContext
{
    /// <summary>
    /// Current slide
    /// </summary>
    public SlideContext Slide { get; set; }

    /// <summary>
    /// Current shape
    /// </summary>
    public ShapeContext Shape { get; set; }

    /// <summary>
    /// Current directive
    /// </summary>
    public DirectiveContext Directive { get; set; }

    /// <summary>
    /// Variables
    /// </summary>
    public Dictionary<string, object> Variables { get; set; }

    /// <summary>
    /// Global variables
    /// </summary>
    public Dictionary<string, Func<object>> GlobalVariables { get; set; }

    /// <summary>
    /// Functions
    /// </summary>
    public Dictionary<string, PowerPointFunction> Functions { get; set; }

    /// <summary>
    /// PowerPoint processing options
    /// </summary>
    public PowerPointOptions Options { get; set; }

    /// <summary>
    /// Current SlidePart being processed
    /// </summary>
    public SlidePart SlidePart { get; set; }

    /// <summary>
    /// Tracks slides that have been processed with array batch data
    /// </summary>
    public HashSet<string> ProcessedArraySlides { get; } = new HashSet<string>();

    /// <summary>
    /// Creates a new PowerPointContext with default values
    /// </summary>
    public PowerPointContext()
    {
        Slide = new SlideContext();
        Shape = new ShapeContext();
        Directive = new DirectiveContext();
        Variables = new Dictionary<string, object>();
        GlobalVariables = new Dictionary<string, Func<object>>();
        Functions = new Dictionary<string, PowerPointFunction>();
    }

    /// <summary>
    /// Resolves a variable by name, supporting property paths
    /// </summary>
    public object ResolveVariable(string name)
    {
        // Check for simple variable
        if (Variables.TryGetValue(name, out var value))
            return value;

        // Check for global variable
        if (GlobalVariables.TryGetValue(name, out var factory))
            return factory();

        // Check for property path (e.g., "data.Items[0].Name")
        if (name.Contains('.'))
        {
            var parts = name.Split('.');
            if (Variables.TryGetValue(parts[0], out var obj) && obj != null)
            {
                return ResolvePropertyPath(obj, parts, 1);
            }
        }

        return null;
    }

    /// <summary>
    /// Resolves a property path on an object
    /// </summary>
    private object ResolvePropertyPath(object obj, string[] parts, int startIndex)
    {
        for (int i = startIndex; i < parts.Length && obj != null; i++)
        {
            string part = parts[i];

            // Handle indexers like Items[0]
            var indexerMatch = System.Text.RegularExpressions.Regex.Match(part, @"^(.+)\[(\d+)\]$");
            if (indexerMatch.Success)
            {
                string propName = indexerMatch.Groups[1].Value;
                int index = int.Parse(indexerMatch.Groups[2].Value);

                // Get property value
                var property = obj.GetType().GetProperty(propName);
                if (property == null)
                    return null;

                var collection = property.GetValue(obj);
                if (collection == null)
                    return null;

                // If it's an array
                if (collection is Array array && index < array.Length)
                {
                    obj = array.GetValue(index);
                }
                // If it's a list or other indexable collection
                else if (collection.GetType().GetProperty("Item") != null &&
                         collection.GetType().GetProperty("Count") != null)
                {
                    var countProp = collection.GetType().GetProperty("Count");
                    int count = (int)countProp.GetValue(collection);

                    if (index < count)
                    {
                        var indexerProp = collection.GetType().GetProperty("Item");
                        obj = indexerProp.GetValue(collection, new object[] { index });
                    }
                    else
                    {
                        return null;
                    }
                }
                else
                {
                    return null;
                }
            }
            else
            {
                // Regular property
                var property = obj.GetType().GetProperty(part);
                if (property == null)
                    return null;

                obj = property.GetValue(obj);
            }
        }

        return obj;
    }
}

/// <summary>
/// Context for slide processing
/// </summary>
public class SlideContext
{
    /// <summary>
    /// Slide index
    /// </summary>
    public int Index { get; set; }

    /// <summary>
    /// Slide ID
    /// </summary>
    public string Id { get; set; }

    /// <summary>
    /// Slide notes
    /// </summary>
    public string Notes { get; set; }
}

/// <summary>
/// Context for shape processing
/// </summary>
public class ShapeContext
{
    /// <summary>
    /// Shape name
    /// </summary>
    public string Name { get; set; }

    /// <summary>
    /// Shape type
    /// </summary>
    public string Type { get; set; }

    /// <summary>
    /// Shape ID
    /// </summary>
    public string Id { get; set; }

    /// <summary>
    /// Shape text
    /// </summary>
    public string Text { get; set; }

    /// <summary>
    /// The actual shape object
    /// </summary>
    public Shape ShapeObject { get; set; }
}

/// <summary>
/// Context for directive processing
/// </summary>
public class DirectiveContext
{
    /// <summary>
    /// Directive name
    /// </summary>
    public string Name { get; set; }

    /// <summary>
    /// Directive value
    /// </summary>
    public string Value { get; set; }

    /// <summary>
    /// Directive parameters
    /// </summary>
    public Dictionary<string, string> Parameters { get; set; } = new();
}