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
    /// Improved pattern for array references with indirect syntax (for shape names and alt text)
    /// </summary>
    public static readonly Regex IndirectArrayPattern = new Regex(@"(\w+)_(\d+)(?:_|$)", RegexOptions.Compiled);

    /// <summary>
    /// Extract array references from text with support for multiple patterns
    /// </summary>
    public static List<ArrayReference> ExtractArrayReferences(string text)
    {
        var result = new List<ArrayReference>();

        if (string.IsNullOrEmpty(text))
            return result;

        // 1. Check dollar sign pattern - ${Items[0].Property}
        var dollarMatches = DollarSignArrayPattern.Matches(text);
        foreach (Match match in dollarMatches)
        {
            result.Add(ParseArrayReference(match));
        }

        // 2. Check direct pattern - Items[0].Property
        var directMatches = DirectArrayPattern.Matches(text);
        foreach (Match match in directMatches)
        {
            // Skip if already captured by dollar sign pattern
            if (!dollarMatches.Cast<Match>().Any(m => m.Value.Contains(match.Value)))
            {
                result.Add(ParseArrayReference(match));
            }
        }

        // 3. Check function argument pattern
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

        // 4. Check indirect pattern (ArrayName_Index format used in shape names)
        var indirectMatches = IndirectArrayPattern.Matches(text);
        foreach (Match match in indirectMatches)
        {
            if (match.Groups.Count >= 3)
            {
                string arrayName = match.Groups[1].Value;
                if (int.TryParse(match.Groups[2].Value, out int index))
                {
                    result.Add(new ArrayReference
                    {
                        ArrayName = arrayName,
                        Index = index,
                        PropertyPath = "",
                        Pattern = match.Value
                    });
                }
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

        // 3. Update indirect ArrayName_Index pattern (for shape names)
        text = IndirectArrayPattern.Replace(text, match =>
        {
            if (match.Groups[1].Value != arrayName)
                return match.Value;

            // Check if index in range
            if (match.Groups.Count > 2 && int.TryParse(match.Groups[2].Value, out int localIndex))
            {
                if (!IsValidIndex(localIndex, offset))
                    return match.Value;

                int newIndex = localIndex + offset;

                // Replace only the number part, preserve the rest of the string
                return $"{arrayName}_{newIndex}{match.Groups[0].Value.Substring(match.Groups[2].Index + match.Groups[2].Length - match.Groups[0].Index)}";
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
