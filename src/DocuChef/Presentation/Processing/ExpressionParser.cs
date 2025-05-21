namespace DocuChef.Presentation.Processing;

/// <summary>
/// Provides utility methods for parsing and analyzing expressions in slides
/// </summary>
internal static class ExpressionParser
{
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

        // Find all ${...} expressions
        var curlyBraceMatches = Regex.Matches(text, @"\$\{([^}]+)\}");
        foreach (Match match in curlyBraceMatches)
        {
            if (match.Success)
            {
                ProcessExpressionContent(match.Groups[1].Value, result);
            }
        }

        // Find all $...$ expressions
        var dollarSignMatches = Regex.Matches(text, @"\$([^$]+)\$");
        foreach (Match match in dollarSignMatches)
        {
            if (match.Success)
            {
                ProcessExpressionContent(match.Groups[1].Value, result);
            }
        }

        return result
            .Select(kvp => new CollectionIndexInfo { CollectionPath = kvp.Key, MaxIndex = kvp.Value })
            .ToList();
    }

    private static void ProcessExpressionContent(string expressionContent, Dictionary<string, int> result)
    {
        // Extract the collection path and index using string operations
        int openBracketIndex = expressionContent.IndexOf('[');
        if (openBracketIndex > 0)
        {
            string collectionPath = expressionContent.Substring(0, openBracketIndex);
            int closeBracketIndex = expressionContent.IndexOf(']', openBracketIndex);

            if (closeBracketIndex > openBracketIndex)
            {
                string indexStr = expressionContent.Substring(openBracketIndex + 1, closeBracketIndex - openBracketIndex - 1);
                if (int.TryParse(indexStr, out int index))
                {
                    // Store the highest index found for each collection path
                    if (!result.ContainsKey(collectionPath) || result[collectionPath] < index)
                    {
                        result[collectionPath] = index;
                    }
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
    /// The collection path (e.g. "Categories" or "Categories>Products" depending on delimiter)
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