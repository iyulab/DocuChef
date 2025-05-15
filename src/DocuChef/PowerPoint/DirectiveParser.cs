using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace DocuChef.PowerPoint;

/// <summary>
/// Parser for slide directives in PowerPoint templates with hierarchical path support
/// </summary>
internal static class DirectiveParser
{
    // Enhanced directive pattern supporting hierarchical paths
    private static readonly Regex DirectivePattern = new(@"#(\w+(?:-\w+)?):([^,]+)(?:,\s*(.+))?", RegexOptions.Compiled);

    /// <summary>
    /// Parse directives from slide notes with hierarchical path support
    /// </summary>
    public static List<Directive> ParseDirectives(string notes)
    {
        var directives = new List<Directive>();
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

                    // Create the directive with hierarchical path support
                    var directive = new Directive
                    {
                        Name = directiveName,
                        Value = directiveValue,
                        Parameters = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                    };

                    // Parse path
                    if (directiveName.StartsWith("foreach", StringComparison.OrdinalIgnoreCase))
                    {
                        directive.Path = new HierarchicalPath(directiveValue);
                    }

                    // Parse additional parameters if present
                    if (match.Groups.Count > 3 && !string.IsNullOrEmpty(match.Groups[3].Value))
                    {
                        ParseParameters(match.Groups[3].Value, directive.Parameters);
                    }

                    directives.Add(directive);
                    Logger.Debug($"Parsed hierarchical directive: {directive.Name} with path: {directive.Path}");
                }
            }
        }
        catch (Exception ex)
        {
            Logger.Error($"Error parsing hierarchical directives from notes: {ex.Message}", ex);
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
/// Represents a directive with hierarchical path support
/// </summary>
internal class Directive
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
    /// Hierarchical path for the directive
    /// </summary>
    public HierarchicalPath Path { get; set; }

    /// <summary>
    /// Additional directive parameters
    /// </summary>
    public Dictionary<string, string> Parameters { get; set; } = new(StringComparer.OrdinalIgnoreCase);

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