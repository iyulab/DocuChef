namespace DocuChef.PowerPoint.Helpers;

/// <summary>
/// Helper class for array-related operations in document templates
/// </summary>
public static class ArrayReferenceHelper
{
    /// <summary>
    /// Pattern for array references in expressions: ${ArrayName[index].Property}
    /// </summary>
    public static readonly Regex DollarSignArrayPattern = new Regex(@"\${(\w+)\[(\d+)\](\.[\w\.]+)?(:.*)?}", RegexOptions.Compiled);

    /// <summary>
    /// Pattern for direct array references: ArrayName[index].Property
    /// </summary>
    public static readonly Regex DirectArrayPattern = new Regex(@"(?<!\$\{)(\w+)\[(\d+)\](\.[\w\.]+)?", RegexOptions.Compiled);

    /// <summary>
    /// Pattern for function arguments with array references
    /// </summary>
    public static readonly Regex FunctionArgArrayPattern = new Regex(@"(\w+)\[(\d+)\](\.[\w\.]+)?([,\)])", RegexOptions.Compiled);

    /// <summary>
    /// Extract array references from text
    /// </summary>
    public static List<ArrayReference> ExtractArrayReferences(string text)
    {
        var result = new List<ArrayReference>();

        if (string.IsNullOrEmpty(text))
            return result;

        // Check dollar sign pattern
        var dollarMatches = DollarSignArrayPattern.Matches(text);
        foreach (Match match in dollarMatches)
        {
            result.Add(ParseArrayReference(match));
        }

        // Check direct pattern
        var directMatches = DirectArrayPattern.Matches(text);
        foreach (Match match in directMatches)
        {
            // Skip if already captured by dollar sign pattern
            if (!dollarMatches.Cast<Match>().Any(m => m.Value.Contains(match.Value)))
            {
                result.Add(ParseArrayReference(match));
            }
        }

        // Check function argument pattern
        var funcArgMatches = FunctionArgArrayPattern.Matches(text);
        foreach (Match match in funcArgMatches)
        {
            // Skip if already captured by other patterns
            if (!dollarMatches.Cast<Match>().Any(m => m.Value.Contains(match.Value)) &&
                !directMatches.Cast<Match>().Any(m => m.Value.Contains(match.Value)))
            {
                result.Add(ParseArrayReference(match));
            }
        }

        return result;
    }

    /// <summary>
    /// Update array references in text with offset
    /// </summary>
    public static string UpdateArrayReferences(string text, string arrayName, int offset)
    {
        if (string.IsNullOrEmpty(text) || offset == 0 || !text.Contains(arrayName))
            return text;

        Logger.Debug($"[ARRAY-REF] Updating array references: arrayName={arrayName}, offset={offset}");

        // 1. Update ${ArrayName[index]} pattern
        text = DollarSignArrayPattern.Replace(text, match =>
        {
            if (match.Groups[1].Value != arrayName)
                return match.Value;

            // Check if index in range
            if (match.Groups.Count > 2 && int.TryParse(match.Groups[2].Value, out int localIndex))
            {
                if (!IsValidIndex(localIndex, offset))
                    return match.Value;

                int newIndex = localIndex + offset;
                string propPath = match.Groups.Count > 3 ? match.Groups[3].Value : "";
                string format = match.Groups.Count > 4 ? match.Groups[4].Value : "";

                return $"${{{arrayName}[{newIndex}]{propPath}{format}}}";
            }
            return match.Value;
        });

        // 2. Update direct ArrayName[index] pattern
        text = DirectArrayPattern.Replace(text, match =>
        {
            if (match.Groups[1].Value != arrayName)
                return match.Value;

            // Check if index in range
            if (match.Groups.Count > 2 && int.TryParse(match.Groups[2].Value, out int localIndex))
            {
                if (!IsValidIndex(localIndex, offset))
                    return match.Value;

                int newIndex = localIndex + offset;
                string propPath = match.Groups.Count > 3 ? match.Groups[3].Value : "";

                // Direct pattern (without ${})
                return $"{arrayName}[{newIndex}]{propPath}";
            }
            return match.Value;
        });

        return text;
    }

    private static ArrayReference ParseArrayReference(Match match)
    {
        string arrayName = match.Groups[1].Value;
        int index = int.Parse(match.Groups[2].Value);
        string propPath = match.Groups.Count > 3 ? match.Groups[3].Value : "";

        return new ArrayReference
        {
            ArrayName = arrayName,
            Index = index,
            PropertyPath = propPath,
            Pattern = match.Value
        };
    }

    private static bool IsValidIndex(int index, int offset)
    {
        if (index < 0 || index > 1000)
        {
            Logger.Warning($"[ARRAY-REF] Invalid index: {index} - out of reasonable range");
            return false;
        }

        int newIndex = index + offset;
        if (newIndex < 0 || newIndex > 1000)
        {
            Logger.Warning($"[ARRAY-REF] Invalid calculated index: {newIndex} - out of reasonable range");
            return false;
        }

        return true;
    }
}

/// <summary>
/// Represents an array reference in document template
/// </summary>
public class ArrayReference
{
    /// <summary>
    /// The name of the array
    /// </summary>
    public string ArrayName { get; set; }

    /// <summary>
    /// The index referenced in the array
    /// </summary>
    public int Index { get; set; }

    /// <summary>
    /// The property path after the array index (if any)
    /// </summary>
    public string PropertyPath { get; set; }

    /// <summary>
    /// The full pattern matched in the text
    /// </summary>
    public string Pattern { get; set; }
}