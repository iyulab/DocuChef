using System.Text.RegularExpressions;

namespace DocuChef.Presentation.Processing;

/// <summary>
/// Provides utility methods for parsing and analyzing expressions in slides
/// </summary>
internal static class ExpressionParser
{
    // Pattern for finding array index expressions like ${Collection[0].Property}
    private static readonly Regex IndexedExpressionPattern = new Regex(
        @"\$\{([A-Za-z0-9_]+(?:_[A-Za-z0-9_]+)*)\[(\d+)\]((?:\.[A-Za-z0-9_]+)*)\}",
        RegexOptions.Compiled);

    // Pattern for finding DollarSignEngine expressions like $Collection[0].Property$
    private static readonly Regex DollarSignExpressionPattern = new Regex(
        @"\$([A-Za-z0-9_]+(?:_[A-Za-z0-9_]+)*)\[(\d+)\]((?:\.[A-Za-z0-9_]+)*)\$",
        RegexOptions.Compiled);

    /// <summary>
    /// Extracts collection path and index information from expression text
    /// </summary>
    /// <param name="text">The text containing expressions</param>
    /// <returns>A collection of detected path and max index pairs</returns>
    public static List<CollectionIndexInfo> ExtractCollectionIndexInfo(string text)
    {
        if (string.IsNullOrEmpty(text))
            return new List<CollectionIndexInfo>();

        var result = new Dictionary<string, int>();

        // Check for ${Collection[index].Property} pattern
        foreach (Match match in IndexedExpressionPattern.Matches(text))
        {
            ProcessExpressionMatch(match, result);
        }

        // Check for $Collection[index].Property$ pattern
        foreach (Match match in DollarSignExpressionPattern.Matches(text))
        {
            ProcessExpressionMatch(match, result);
        }

        return result
            .Select(kvp => new CollectionIndexInfo { CollectionPath = kvp.Key, MaxIndex = kvp.Value })
            .ToList();
    }

    /// <summary>
    /// Processes a regex match to extract collection path and index
    /// </summary>
    private static void ProcessExpressionMatch(Match match, Dictionary<string, int> result)
    {
        if (match.Success && match.Groups.Count >= 3)
        {
            string collectionPath = match.Groups[1].Value;

            if (int.TryParse(match.Groups[2].Value, out int index))
            {
                // Store the highest index found for each collection path
                if (!result.ContainsKey(collectionPath) || result[collectionPath] < index)
                {
                    result[collectionPath] = index;
                }
            }
        }
    }

    /// <summary>
    /// Creates implicit foreach directives based on expression indexing
    /// </summary>
    /// <param name="text">The text containing expressions</param>
    /// <returns>A list of implicit directives detected</returns>
    public static List<ImplicitDirectiveInfo> CreateImplicitDirectives(string text)
    {
        var indexInfo = ExtractCollectionIndexInfo(text);
        var result = new List<ImplicitDirectiveInfo>();

        foreach (var info in indexInfo)
        {
            // For each collection, create a foreach directive with max items = max index + 1
            // (since indices are 0-based but we need at least max index + 1 items)
            result.Add(new ImplicitDirectiveInfo
            {
                DirectiveType = "foreach",
                CollectionPath = info.CollectionPath,
                MaxItems = info.MaxIndex + 1
            });
        }

        return result;
    }
}

/// <summary>
/// Contains collection path and index information extracted from expressions
/// </summary>
public class CollectionIndexInfo
{
    /// <summary>
    /// The collection path (e.g. "Categories" or "Categories_Products")
    /// </summary>
    public string CollectionPath { get; set; }

    /// <summary>
    /// The maximum index used in expressions for this collection
    /// </summary>
    public int MaxIndex { get; set; }

    /// <summary>
    /// Returns a string representation of the collection index info
    /// </summary>
    public override string ToString()
    {
        return $"{CollectionPath}[{MaxIndex}]";
    }
}

/// <summary>
/// Contains information about an implicit directive detected from expressions
/// </summary>
public class ImplicitDirectiveInfo
{
    /// <summary>
    /// The type of directive (e.g. "foreach")
    /// </summary>
    public string DirectiveType { get; set; }

    /// <summary>
    /// The collection path for the directive
    /// </summary>
    public string CollectionPath { get; set; }

    /// <summary>
    /// The maximum number of items to process
    /// </summary>
    public int MaxItems { get; set; }

    /// <summary>
    /// Returns a string representation of the implicit directive
    /// </summary>
    public override string ToString()
    {
        return $"#{DirectiveType}: {CollectionPath}, max: {MaxItems}";
    }
}