using System.Text.RegularExpressions;

namespace DocuChef.PowerPoint;

/// <summary>
/// Parser for slide directives in PowerPoint templates
/// </summary>
internal static class DirectiveParser
{
    // Improved regex pattern for directives that supports foreach-items pattern
    private static readonly Regex DirectivePattern = new(@"#(\w+(?:-\w+)?):([^,]+)(?:,\s*(.+))?", RegexOptions.Compiled);

    /// <summary>
    /// Parse directives from slide notes
    /// </summary>
    public static List<SlideDirective> ParseDirectives(string notes)
    {
        var directives = new List<SlideDirective>();
        if (string.IsNullOrEmpty(notes))
            return directives;

        try
        {
            // Find all directive patterns in the notes
            var matches = DirectivePattern.Matches(notes);
            foreach (Match match in matches)
            {
                if (match.Groups.Count >= 3)
                {
                    string directiveName = match.Groups[1].Value.Trim();
                    string directiveValue = match.Groups[2].Value.Trim();

                    // Special handling for foreach-items directive
                    if (directiveName.StartsWith("foreach-"))
                    {
                        // Extract the collection path from foreach-items directive
                        string collectionType = directiveName.Substring(8); // after "foreach-"

                        // Parse path like "Groups.Items" into parent and child components
                        var pathParts = directiveValue.Split('.');
                        if (pathParts.Length >= 2)
                        {
                            var directive = new SlideDirective
                            {
                                Name = "foreach-nested",
                                Value = directiveValue, // Keep the original path (e.g., "Groups.Items")
                                Parameters = new Dictionary<string, string>
                                {
                                    ["parent"] = pathParts[0],
                                    ["child"] = pathParts[1],
                                    ["type"] = collectionType
                                }
                            };

                            // Parse additional parameters if present
                            if (match.Groups.Count > 3 && !string.IsNullOrEmpty(match.Groups[3].Value))
                            {
                                ParseParameters(match.Groups[3].Value, directive.Parameters);
                            }

                            directives.Add(directive);
                            Logger.Debug($"Parsed nested directive: {directive.Name} with parent: {pathParts[0]}, child: {pathParts[1]}");
                        }
                    }
                    else
                    {
                        // Standard directive handling
                        var directive = new SlideDirective
                        {
                            Name = directiveName,
                            Value = directiveValue,
                            Parameters = new Dictionary<string, string>()
                        };

                        // Parse additional parameters if present
                        if (match.Groups.Count > 3 && !string.IsNullOrEmpty(match.Groups[3].Value))
                        {
                            ParseParameters(match.Groups[3].Value, directive.Parameters);
                        }

                        directives.Add(directive);
                        Logger.Debug($"Parsed directive: {directive.Name} with value: {directive.Value}");
                    }
                }
            }
        }
        catch (Exception ex)
        {
            Logger.Error($"Error parsing directives from notes: {ex.Message}", ex);
        }

        return directives;
    }

    /// <summary>
    /// Parse directive parameters (param1: value1, param2: value2)
    /// </summary>
    private static void ParseParameters(string paramString, Dictionary<string, string> parameters)
    {
        foreach (var paramPair in SplitParameters(paramString))
        {
            var pair = paramPair.Trim();
            var separatorIndex = pair.IndexOf(':');

            if (separatorIndex > 0)
            {
                var name = pair.Substring(0, separatorIndex).Trim();
                var value = pair.Substring(separatorIndex + 1).Trim();

                // Remove surrounding quotes if present
                if (value.StartsWith("\"") && value.EndsWith("\"") && value.Length > 1)
                {
                    value = value.Substring(1, value.Length - 2);
                }

                parameters[name] = value;
                Logger.Debug($"Parsed parameter: {name} = {value}");
            }
        }
    }

    /// <summary>
    /// Split parameters string respecting quotes
    /// </summary>
    private static IEnumerable<string> SplitParameters(string paramString)
    {
        var result = new List<string>();
        bool inQuotes = false;
        int startIndex = 0;

        for (int i = 0; i < paramString.Length; i++)
        {
            char c = paramString[i];

            // Handle quotes
            if (c == '"' && (i == 0 || paramString[i - 1] != '\\'))
            {
                inQuotes = !inQuotes;
            }
            // Handle parameter separator
            else if (c == ',' && !inQuotes)
            {
                result.Add(paramString.Substring(startIndex, i - startIndex));
                startIndex = i + 1;
            }
        }

        // Add the last parameter
        if (startIndex < paramString.Length)
        {
            result.Add(paramString.Substring(startIndex));
        }

        return result;
    }
}

/// <summary>
/// Represents a directive parsed from slide notes
/// </summary>
internal class SlideDirective
{
    /// <summary>
    /// Directive name (e.g., "foreach")
    /// </summary>
    public string Name { get; set; }

    /// <summary>
    /// Primary directive value
    /// </summary>
    public string Value { get; set; }

    /// <summary>
    /// Additional directive parameters
    /// </summary>
    public Dictionary<string, string> Parameters { get; set; } = new Dictionary<string, string>();

    /// <summary>
    /// Checks if a parameter exists
    /// </summary>
    public bool HasParameter(string name)
    {
        return Parameters.ContainsKey(name);
    }

    /// <summary>
    /// Gets a parameter value, or default if not found
    /// </summary>
    public string GetParameter(string name, string defaultValue = null)
    {
        return Parameters.TryGetValue(name, out var value) ? value : defaultValue;
    }

    /// <summary>
    /// Gets a parameter value as int, or default if not found or not parsable
    /// </summary>
    public int GetParameterAsInt(string name, int defaultValue = 0)
    {
        if (Parameters.TryGetValue(name, out var value) && int.TryParse(value, out var intValue))
        {
            return intValue;
        }
        return defaultValue;
    }

    /// <summary>
    /// Gets a parameter value as bool, or default if not found or not parsable
    /// </summary>
    public bool GetParameterAsBool(string name, bool defaultValue = false)
    {
        if (Parameters.TryGetValue(name, out var value))
        {
            value = value.ToLowerInvariant();
            if (value == "true" || value == "yes" || value == "1")
                return true;
            if (value == "false" || value == "no" || value == "0")
                return false;
        }
        return defaultValue;
    }
}