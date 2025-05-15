namespace DocuChef.PowerPoint;

/// <summary>
/// Represents a hierarchical path to navigate nested collections and properties
/// </summary>
public class HierarchicalPath
{
    private static readonly Regex IndexerPattern = new(@"^(.+?)\[(\d+)\]$", RegexOptions.Compiled);
    private static readonly Regex ArrayIndexPattern = new(@"(.+?)\[(\d+)\]", RegexOptions.Compiled);
    private static readonly Regex NestedPathPattern = new(@"^(\w+)\.(\w+(?:\.\w+)*)$", RegexOptions.Compiled);

    /// <summary>
    /// Individual segments in the path (e.g., ["Departments", "Teams", "Members"])
    /// </summary>
    public List<PathSegment> Segments { get; } = new();

    /// <summary>
    /// Full path as a dot-separated string (e.g., "Departments.Teams.Members")
    /// </summary>
    public string FullPath => string.Join(".", Segments.Select(s => s.ToString()));

    /// <summary>
    /// Creates a new empty hierarchical path
    /// </summary>
    public HierarchicalPath() { }

    /// <summary>
    /// Creates a hierarchical path from a dot-separated string
    /// </summary>
    public HierarchicalPath(string path)
    {
        if (string.IsNullOrEmpty(path))
            return;

        Parse(path);
    }

    /// <summary>
    /// Creates a hierarchical path from another path
    /// </summary>
    public HierarchicalPath(HierarchicalPath other)
    {
        if (other == null)
            return;

        foreach (var segment in other.Segments)
        {
            Segments.Add(new PathSegment(segment));
        }
    }

    /// <summary>
    /// Parses a dot-separated path string into segments
    /// </summary>
    public void Parse(string path)
    {
        Segments.Clear();

        if (string.IsNullOrEmpty(path))
            return;

        // Special handling for ppt.Function("param") pattern
        if (path.StartsWith("ppt.") && path.Contains("(") && path.EndsWith(")"))
        {
            Segments.Add(new PathSegment(path));
            return;
        }

        // Special handling for underscore notation (Categories_Products)
        if (path.Contains('_') && !path.Contains('.') && !path.Contains('['))
        {
            ParseUnderscoreNotation(path);
            return;
        }

        // Special handling for paths with array indices (Categories[0].Products[1])
        if (path.Contains('[') && path.Contains(']'))
        {
            ParseArrayIndexNotation(path);
            return;
        }

        // Standard parsing for dot notation (Categories.Products)
        ParseDotNotation(path);
    }

    /// <summary>
    /// Parse underscore notation (Categories_Products)
    /// </summary>
    private void ParseUnderscoreNotation(string path)
    {
        string[] parts = path.Split('_');
        foreach (string part in parts)
        {
            if (string.IsNullOrEmpty(part))
                continue;

            // Check if part has an indexer: Products0
            var digitMatch = Regex.Match(part, @"^(\w+)(\d+)$");
            if (digitMatch.Success && digitMatch.Groups.Count >= 3)
            {
                string name = digitMatch.Groups[1].Value;
                int index = int.Parse(digitMatch.Groups[2].Value);
                Segments.Add(new PathSegment(name, index));
            }
            else
            {
                Segments.Add(new PathSegment(part));
            }
        }
    }

    /// <summary>
    /// Parse array index notation (Categories[0].Products[1])
    /// </summary>
    private void ParseArrayIndexNotation(string path)
    {
        // First check for a simple case like "Items[0]"
        var simpleMatch = Regex.Match(path, @"^(\w+)\[(\d+)\]$");
        if (simpleMatch.Success)
        {
            string name = simpleMatch.Groups[1].Value;
            int index = int.Parse(simpleMatch.Groups[2].Value);
            Segments.Add(new PathSegment(name, index));
            return;
        }

        // Handle complex case like "Categories[0].Products[1]"
        if (path.Contains('.'))
        {
            ParseDotNotation(path);
            return;
        }

        // Complex case like "Items[0]Property" or multiple indices
        var parts = new List<string>();
        var currentPath = path;

        // Extract all array references
        var arrayMatches = ArrayIndexPattern.Matches(path);
        foreach (Match match in arrayMatches.Cast<Match>())
        {
            string fullMatch = match.Value;
            string name = match.Groups[1].Value;
            int index = int.Parse(match.Groups[2].Value);

            int position = currentPath.IndexOf(fullMatch);
            if (position > 0)
            {
                // Handle prefix before this match
                string prefix = currentPath.Substring(0, position);
                parts.Add(prefix);
                currentPath = currentPath.Substring(position + fullMatch.Length);
            }
            else
            {
                currentPath = currentPath.Substring(fullMatch.Length);
            }

            parts.Add($"{name}[{index}]");
        }

        // Add remaining text
        if (!string.IsNullOrEmpty(currentPath))
        {
            parts.Add(currentPath);
        }

        // Process each part
        foreach (var part in parts)
        {
            var indexMatch = IndexerPattern.Match(part);
            if (indexMatch.Success)
            {
                string name = indexMatch.Groups[1].Value;
                int index = int.Parse(indexMatch.Groups[2].Value);
                Segments.Add(new PathSegment(name, index));
            }
            else
            {
                Segments.Add(new PathSegment(part));
            }
        }
    }

    /// <summary>
    /// Parse dot notation (Categories.Products)
    /// </summary>
    private void ParseDotNotation(string path)
    {
        string[] dotParts = path.Split('.');
        foreach (string part in dotParts)
        {
            if (string.IsNullOrEmpty(part))
                continue;

            // Check if segment has an indexer: Property[0]
            var match = IndexerPattern.Match(part);
            if (match.Success)
            {
                string name = match.Groups[1].Value;
                int index = int.Parse(match.Groups[2].Value);
                Segments.Add(new PathSegment(name, index));
            }
            else
            {
                Segments.Add(new PathSegment(part));
            }
        }
    }

    /// <summary>
    /// Gets the root segment of the path
    /// </summary>
    public PathSegment GetRoot()
    {
        return Segments.Count > 0 ? Segments[0] : null;
    }

    /// <summary>
    /// Gets a subpath up to the specified index (exclusive)
    /// </summary>
    public HierarchicalPath GetPathUpTo(int endIndex)
    {
        if (endIndex <= 0)
            return new HierarchicalPath();

        int effectiveEnd = Math.Min(endIndex, Segments.Count);

        var result = new HierarchicalPath();
        for (int i = 0; i < effectiveEnd; i++)
        {
            result.Segments.Add(new PathSegment(Segments[i]));
        }

        return result;
    }

    /// <summary>
    /// Gets the path converted to underscore format (e.g., "Departments_Teams_Members")
    /// </summary>
    public string ToUnderscoreFormat()
    {
        return string.Join("_", Segments.Select(s => s.Name));
    }

    /// <summary>
    /// Gets the path converted to dot format (e.g., "Departments.Teams.Members")
    /// </summary>
    public string ToDotFormat()
    {
        return string.Join(".", Segments.Select(s => s.Name));
    }

    /// <summary>
    /// Adds a segment to the path
    /// </summary>
    public void AddSegment(string name, int? index = null)
    {
        Segments.Add(new PathSegment(name, index));
    }

    /// <summary>
    /// Gets a combined path with another hierarchical path
    /// </summary>
    public HierarchicalPath Combine(HierarchicalPath other)
    {
        if (other == null)
            return new HierarchicalPath(this);

        var result = new HierarchicalPath(this);
        foreach (var segment in other.Segments)
        {
            result.Segments.Add(new PathSegment(segment));
        }

        return result;
    }

    /// <summary>
    /// Determines if this path starts with another path
    /// </summary>
    public bool StartsWith(HierarchicalPath other)
    {
        if (other == null || other.Segments.Count == 0 || other.Segments.Count > Segments.Count)
            return false;

        for (int i = 0; i < other.Segments.Count; i++)
        {
            if (!Segments[i].Name.Equals(other.Segments[i].Name, StringComparison.OrdinalIgnoreCase))
                return false;
        }

        return true;
    }

    /// <summary>
    /// Returns the depth (number of segments) in the path
    /// </summary>
    public int Depth => Segments.Count;

    /// <summary>
    /// Returns the path as a string
    /// </summary>
    public override string ToString()
    {
        return FullPath;
    }

    /// <summary>
    /// Creates a hierarchical path from underscore format (e.g., "Departments_Teams_Members")
    /// </summary>
    public static HierarchicalPath FromUnderscoreFormat(string path)
    {
        if (string.IsNullOrEmpty(path))
            return new HierarchicalPath();

        return new HierarchicalPath(path.Replace("_", "."));
    }

    /// <summary>
    /// Creates a hierarchical path with index adjustments for context
    /// </summary>
    public HierarchicalPath WithIndexAdjustments(Dictionary<string, int> indices)
    {
        var result = new HierarchicalPath();

        foreach (var segment in Segments)
        {
            // Copy segment without index first
            var newSegment = new PathSegment(segment.Name);

            // Handle index adjustments
            if (segment.Index.HasValue)
            {
                // Keep the explicit index
                newSegment.Index = segment.Index.Value;
            }
            else if (indices.TryGetValue(segment.Name, out int contextIndex))
            {
                // Use context index if available
                newSegment.Index = contextIndex;
            }

            result.Segments.Add(newSegment);
        }

        return result;
    }
}

/// <summary>
/// Represents a segment in a hierarchical path
/// </summary>
public class PathSegment
{
    /// <summary>
    /// Name of the segment (e.g., "Teams")
    /// </summary>
    public string Name { get; set; }

    /// <summary>
    /// Optional index for array access (e.g., Teams[0])
    /// </summary>
    public int? Index { get; set; }

    /// <summary>
    /// Creates a new path segment with the specified name and optional index
    /// </summary>
    public PathSegment(string name, int? index = null)
    {
        Name = name;
        Index = index;
    }

    /// <summary>
    /// Creates a copy of another path segment
    /// </summary>
    public PathSegment(PathSegment other)
    {
        if (other == null)
            throw new ArgumentNullException(nameof(other));

        Name = other.Name;
        Index = other.Index;
    }

    /// <summary>
    /// Returns the segment as a string
    /// </summary>
    public override string ToString()
    {
        return Index.HasValue ? $"{Name}[{Index.Value}]" : Name;
    }
}