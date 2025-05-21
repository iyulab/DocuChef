using System.Text.RegularExpressions;

namespace DocuChef.Presentation.Directives;

/// <summary>
/// Interface for directive parsers
/// </summary>
public interface IDirectiveParser
{
    /// <summary>
    /// Gets the directive type this parser handles
    /// </summary>
    DirectiveType DirectiveType { get; }

    /// <summary>
    /// Tries to parse a directive from the given text
    /// </summary>
    bool TryParse(string text, out Directive directive);
}

/// <summary>
/// Processes directives in slide notes to control presentation generation flow
/// </summary>
internal class DirectiveProcessor
{
    private readonly IEnumerable<IDirectiveParser> _parsers;

    /// <summary>
    /// Initializes a new instance of DirectiveProcessor with default parsers
    /// </summary>
    public DirectiveProcessor()
    {
        _parsers = new IDirectiveParser[]
        {
            new ForeachDirectiveParser(),
            new IfDirectiveParser()
        };
    }

    /// <summary>
    /// Initializes a new instance of DirectiveProcessor with custom parsers
    /// </summary>
    public DirectiveProcessor(IEnumerable<IDirectiveParser> parsers)
    {
        _parsers = parsers ?? throw new ArgumentNullException(nameof(parsers));
    }

    /// <summary>
    /// Parses a directive from slide note text
    /// </summary>
    public Directive Parse(string noteText)
    {
        if (string.IsNullOrEmpty(noteText))
        {
            Logger.Debug("DirectiveProcessor.Parse: noteText is null or empty");
            return null;
        }

        // Clean up text
        noteText = noteText.Trim();
        Logger.Debug($"DirectiveProcessor.Parse: Processing note text: '{noteText}'");

        // Try each parser in sequence
        foreach (var parser in _parsers)
        {
            if (parser.TryParse(noteText, out var directive))
            {
                Logger.Debug($"DirectiveProcessor.Parse: Successfully parsed {directive.Type} directive");
                return directive;
            }
        }

        Logger.Debug("DirectiveProcessor.Parse: No directive pattern matched");
        return null;
    }

    /// <summary>
    /// Evaluates a directive condition using the provided context
    /// </summary>
    public static bool EvaluateCondition(string condition, Models.SlideContext context)
    {
        if (string.IsNullOrEmpty(condition) || context == null)
            return false;

        // Get value from context
        string value = context.GetContextValue(condition);
        Logger.Debug($"Evaluating condition: '{condition}' = '{value}'");

        // Simplified condition evaluation logic
        bool result = !string.IsNullOrEmpty(value) &&
                    (value.Equals("true", StringComparison.OrdinalIgnoreCase) ||
                    (!value.Equals("false", StringComparison.OrdinalIgnoreCase) &&
                    !string.IsNullOrWhiteSpace(value)));

        return result;
    }
}

/// <summary>
/// Parser for foreach directives
/// </summary>
internal class ForeachDirectiveParser : IDirectiveParser
{
    // Regex pattern for foreach directive
    private static readonly Regex ForeachPattern = new Regex(
        @"#foreach\s*:\s*([\w\._]+)(?:\s*,\s*max\s*:\s*(\d+))?",
        RegexOptions.Compiled | RegexOptions.IgnoreCase);

    /// <summary>
    /// Gets the directive type this parser handles
    /// </summary>
    public DirectiveType DirectiveType => DirectiveType.Foreach;

    /// <summary>
    /// Tries to parse a foreach directive from the given text
    /// </summary>
    public bool TryParse(string text, out Directive directive)
    {
        directive = null;

        if (string.IsNullOrEmpty(text))
            return false;

        // Try standard pattern
        var match = ForeachPattern.Match(text);
        if (match.Success)
        {
            string collectionName = match.Groups[1].Value;
            int maxItems = int.MaxValue;

            // Extract max items if present
            if (match.Groups[2].Success && int.TryParse(match.Groups[2].Value, out int parsedMax))
            {
                maxItems = parsedMax > 0 ? parsedMax : int.MaxValue;
            }

            directive = new ForeachDirective
            {
                CollectionName = collectionName,
                MaxItems = maxItems
            };

            Logger.Debug($"Parsed foreach directive for collection '{collectionName}' with max items {maxItems}");
            return true;
        }

        // Try fallback method for non-standard formats
        if (text.Contains("#foreach", StringComparison.OrdinalIgnoreCase))
        {
            return TryParseFallback(text, out directive);
        }

        return false;
    }

    /// <summary>
    /// Fallback parsing for foreach directives with unusual formatting
    /// </summary>
    private bool TryParseFallback(string text, out Directive directive)
    {
        directive = null;
        Logger.Debug($"ForeachDirectiveParser.TryParseFallback: Attempting fallback parse for: '{text}'");

        // Extract collection name
        int foreachIndex = text.IndexOf("#foreach", StringComparison.OrdinalIgnoreCase);
        int colonIndex = text.IndexOf(':', foreachIndex);

        if (colonIndex < 0)
        {
            Logger.Debug("ForeachDirectiveParser.TryParseFallback: No colon found after #foreach");
            return false;
        }

        string afterColon = text.Substring(colonIndex + 1).Trim();
        string collectionName;
        int maxItems = int.MaxValue;

        // Try to separate collection name and max value
        int commaIndex = afterColon.IndexOf(',');

        if (commaIndex > 0)
        {
            // Extract collection name before comma
            collectionName = afterColon.Substring(0, commaIndex).Trim();

            // Look for max value after comma
            string afterComma = afterColon.Substring(commaIndex + 1).Trim();

            if (afterComma.StartsWith("max:", StringComparison.OrdinalIgnoreCase))
            {
                string valueText = afterComma.Substring(4).Trim();
                if (valueText.StartsWith(":"))
                    valueText = valueText.Substring(1).Trim();

                // Extract digits
                string numericPart = ExtractNumericPart(valueText);
                if (!string.IsNullOrEmpty(numericPart) && int.TryParse(numericPart, out int parsedMax))
                {
                    maxItems = parsedMax > 0 ? parsedMax : int.MaxValue;
                }
            }
        }
        else
        {
            // No comma, assume entire text after colon is collection name
            collectionName = afterColon.Trim();
        }

        if (string.IsNullOrEmpty(collectionName))
        {
            Logger.Debug("ForeachDirectiveParser.TryParseFallback: Could not extract collection name");
            return false;
        }

        Logger.Debug($"ForeachDirectiveParser.TryParseFallback: Created foreach directive - Collection: {collectionName}, MaxItems: {maxItems}");

        directive = new ForeachDirective
        {
            CollectionName = collectionName,
            MaxItems = maxItems
        };

        return true;
    }

    /// <summary>
    /// Extracts the numeric part from a string
    /// </summary>
    private string ExtractNumericPart(string text)
    {
        if (string.IsNullOrEmpty(text))
            return string.Empty;

        string numericPart = "";
        bool foundDigit = false;

        foreach (char c in text)
        {
            if (char.IsDigit(c))
            {
                numericPart += c;
                foundDigit = true;
            }
            else if (foundDigit)
            {
                // Stop at first non-digit after finding a digit
                break;
            }
        }

        return numericPart;
    }
}

/// <summary>
/// Parser for if directives
/// </summary>
internal class IfDirectiveParser : IDirectiveParser
{
    // Regex pattern for if directive
    private static readonly Regex IfPattern = new Regex(
        @"#if\s*:\s*([\w\._]+)",
        RegexOptions.Compiled | RegexOptions.IgnoreCase);

    /// <summary>
    /// Gets the directive type this parser handles
    /// </summary>
    public DirectiveType DirectiveType => DirectiveType.If;

    /// <summary>
    /// Tries to parse an if directive from the given text
    /// </summary>
    public bool TryParse(string text, out Directive directive)
    {
        directive = null;

        if (string.IsNullOrEmpty(text))
            return false;

        // Try standard pattern
        var match = IfPattern.Match(text);
        if (match.Success)
        {
            string condition = match.Groups[1].Value;

            directive = new IfDirective
            {
                Condition = condition
            };

            Logger.Debug($"Parsed if directive with condition '{condition}'");
            return true;
        }

        // Try fallback method for non-standard formats
        if (text.Contains("#if", StringComparison.OrdinalIgnoreCase))
        {
            return TryParseFallback(text, out directive);
        }

        return false;
    }

    /// <summary>
    /// Fallback parsing for if directives with unusual formatting
    /// </summary>
    private bool TryParseFallback(string text, out Directive directive)
    {
        directive = null;
        Logger.Debug($"IfDirectiveParser.TryParseFallback: Attempting fallback parse for: '{text}'");

        int ifIndex = text.IndexOf("#if", StringComparison.OrdinalIgnoreCase);
        int colonIndex = text.IndexOf(':', ifIndex);

        if (colonIndex < 0)
        {
            Logger.Debug("IfDirectiveParser.TryParseFallback: No colon found after #if");
            return false;
        }

        string condition = text.Substring(colonIndex + 1).Trim();

        if (string.IsNullOrEmpty(condition))
        {
            Logger.Debug("IfDirectiveParser.TryParseFallback: Could not extract condition");
            return false;
        }

        Logger.Debug($"IfDirectiveParser.TryParseFallback: Created if directive - Condition: {condition}");

        directive = new IfDirective
        {
            Condition = condition
        };

        return true;
    }
}