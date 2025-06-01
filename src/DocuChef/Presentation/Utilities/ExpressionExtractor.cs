using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;

namespace DocuChef.Presentation.Utilities;

/// <summary>
/// Centralized expression extraction utility that prioritizes format preservation
/// </summary>
public static class ExpressionExtractor
{
    private static readonly Regex BindingExpressionRegex = new(@"\$\{([^}]+)\}", RegexOptions.Compiled);    /// <summary>
                                                                                                            /// Extracts expressions from a slide using a format-preserving approach
                                                                                                            /// 1. First attempts span-level extraction to preserve formatting within expressions
                                                                                                            /// 2. Falls back to paragraph-level processing for incomplete expressions
                                                                                                            /// </summary>
                                                                                                            /// <param name="slide">The slide object (Slide or SlidePart) to extract expressions from</param>
                                                                                                            /// <returns>Collection of unique binding expressions</returns>
    public static IEnumerable<string> ExtractExpressionsFromSlide(object slide)
    {
        var expressions = new HashSet<string>();

        // Handle both Slide and SlidePart objects
        Slide? presentationSlide = slide switch
        {
            Slide s => s,
            SlidePart sp => sp.Slide,
            _ => null
        };

        if (presentationSlide == null)
        {
            DocuChef.Logging.Logger.Debug($"ExpressionExtractor: Could not get Slide object, input type: {slide?.GetType().Name ?? "null"}");
            return expressions;
        }

        DocuChef.Logging.Logger.Debug($"ExpressionExtractor: Processing slide for expression extraction");

        // Priority 1: Extract complete expressions at span level (preserves formatting)
        ExtractFromSpanLevel(presentationSlide, expressions);
        DocuChef.Logging.Logger.Debug($"ExpressionExtractor: Span-level extraction found {expressions.Count} expressions");

        // Priority 2: Extract from paragraph level for any missed expressions
        var beforeParagraphCount = expressions.Count;
        ExtractFromParagraphLevel(presentationSlide, expressions);
        DocuChef.Logging.Logger.Debug($"ExpressionExtractor: Paragraph-level extraction added {expressions.Count - beforeParagraphCount} additional expressions");

        DocuChef.Logging.Logger.Debug($"ExpressionExtractor: Total expressions found: {expressions.Count}");
        foreach (var expr in expressions)
        {
            DocuChef.Logging.Logger.Debug($"ExpressionExtractor: Found expression: '{expr}'");
        }

        return expressions;
    }    /// <summary>
         /// Extracts expressions from individual text spans to preserve formatting
         /// This is the preferred method as it maintains format boundaries
         /// </summary>
    private static void ExtractFromSpanLevel(Slide slide, HashSet<string> expressions)
    {
        var textElements = slide.Descendants<DocumentFormat.OpenXml.Drawing.Text>().ToList();
        DocuChef.Logging.Logger.Debug($"ExpressionExtractor: Span-level processing {textElements.Count} text elements");

        foreach (var textElement in textElements)
        {
            if (string.IsNullOrEmpty(textElement.Text))
                continue;

            DocuChef.Logging.Logger.Debug($"ExpressionExtractor: Processing text span: '{textElement.Text}'");

            // Extract complete expressions from this span
            var matches = BindingExpressionRegex.Matches(textElement.Text);
            foreach (Match match in matches)
            {
                expressions.Add(match.Value);
                DocuChef.Logging.Logger.Debug($"ExpressionExtractor: Found span-level expression: '{match.Value}'");
            }
        }
    }

    /// <summary>
    /// Extracts expressions from paragraph level as fallback
    /// Used when expressions span multiple formatting runs
    /// </summary>
    private static void ExtractFromParagraphLevel(Slide slide, HashSet<string> expressions)
    {
        var paragraphs = slide.Descendants<DocumentFormat.OpenXml.Drawing.Paragraph>().ToList();
        DocuChef.Logging.Logger.Debug($"ExpressionExtractor: Paragraph-level processing {paragraphs.Count} paragraphs");

        foreach (var paragraph in paragraphs)
        {
            var paragraphText = GetParagraphText(paragraph);
            if (string.IsNullOrEmpty(paragraphText))
                continue;

            DocuChef.Logging.Logger.Debug($"ExpressionExtractor: Processing paragraph text: '{paragraphText}'");

            // Extract expressions that might span multiple runs
            var matches = BindingExpressionRegex.Matches(paragraphText);
            foreach (Match match in matches)
            {
                if (expressions.Add(match.Value)) // Only log if it's a new expression
                {
                    DocuChef.Logging.Logger.Debug($"ExpressionExtractor: Found paragraph-level expression: '{match.Value}'");
                }
            }
        }
    }    /// <summary>
         /// Concatenates all text from a paragraph's runs
         /// </summary>
    private static string GetParagraphText(DocumentFormat.OpenXml.Drawing.Paragraph paragraph)
    {
        var textParts = new List<string>();

        foreach (var run in paragraph.Descendants<DocumentFormat.OpenXml.Drawing.Run>())
        {
            foreach (var text in run.Descendants<DocumentFormat.OpenXml.Drawing.Text>())
            {
                if (!string.IsNullOrEmpty(text.Text))
                {
                    textParts.Add(text.Text);
                }
            }
        }

        return string.Join("", textParts);
    }

    /// <summary>
    /// Extracts expressions from slide using direct text concatenation (legacy fallback)
    /// </summary>
    public static IEnumerable<string> ExtractExpressionsFromSlideText(string slideText)
    {
        if (string.IsNullOrEmpty(slideText))
            return Enumerable.Empty<string>();

        var matches = BindingExpressionRegex.Matches(slideText);
        return matches.Cast<Match>()
                     .Select(m => m.Value)
                     .Distinct();
    }

    /// <summary>
    /// Extracts variable names from expression text (without ${} wrapper)
    /// </summary>
    /// <param name="expressions">Collection of expression strings</param>
    /// <returns>Collection of unique variable names</returns>
    public static IEnumerable<string> ExtractVariableNames(IEnumerable<string> expressions)
    {
        var variableNames = new HashSet<string>();

        foreach (var expression in expressions)
        {
            // Extract inner expression from ${...}
            var innerExpression = expression;
            if (expression.StartsWith("${") && expression.EndsWith("}"))
            {
                innerExpression = expression.Substring(2, expression.Length - 3);
            }

            // Handle different expression types
            if (innerExpression.StartsWith("ppt."))
            {
                // PowerPoint functions - extract variable references from parameters
                var paramMatch = Regex.Match(innerExpression, @"ppt\.\w+\(([^)]*)\)");
                if (paramMatch.Success)
                {
                    var parameters = paramMatch.Groups[1].Value.Split(',')
                        .Select(p => p.Trim(' ', '"', '\''))
                        .Where(p => !string.IsNullOrEmpty(p) && !p.All(char.IsDigit));

                    foreach (var param in parameters)
                    {
                        variableNames.Add(param);
                    }
                }
            }
            else
            {
                // Regular variable or property access
                // Remove array indices and format specifiers for variable name extraction
                var cleanName = RemoveArrayIndicesAndFormatters(innerExpression);
                if (!string.IsNullOrEmpty(cleanName))
                {
                    variableNames.Add(cleanName);
                }
            }
        }

        return variableNames;
    }

    /// <summary>
    /// Removes array indices [n] and format specifiers :format from expression
    /// </summary>
    private static string RemoveArrayIndicesAndFormatters(string expression)
    {
        // Remove format specifiers (everything after :)
        var colonIndex = expression.IndexOf(':');
        if (colonIndex >= 0)
        {
            expression = expression.Substring(0, colonIndex);
        }

        // Remove array indices [n]
        expression = Regex.Replace(expression, @"\[\d+\]", "");

        // For context operators, extract the base variable name
        if (expression.Contains('>'))
        {
            var parts = expression.Split('>');
            return parts[0].Trim();
        }

        return expression.Trim();
    }
}
