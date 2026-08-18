using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using DollarSignEngine;
using DocuChef.Presentation.Exceptions;
using DocuChef.Word;
using WText = DocumentFormat.OpenXml.Wordprocessing.Text;

namespace DocuChef.Word.Processors;

/// <summary>
/// Evaluates ${expression} placeholders in Word document text elements
/// using DollarSignEngine for expression evaluation.
/// </summary>
public static class TextBinder
{
    /// <summary>
    /// Binds data to all ${expression} placeholders in the given container.
    /// </summary>
    public static void Bind(OpenXmlElement container, Dictionary<string, object> data, WordOptions? options = null)
    {
        var throwOnMissingVariable = options?.ThrowOnMissingVariable ?? false;
        var evalOptions = new DollarSignOptions
        {
            SupportDollarSignSyntax = true,
            ThrowOnError = throwOnMissingVariable
        };
        // See DataBinder.EvaluateTemplate for why ErrorHandler is wired only in the non-throwing
        // case: DollarSignEngine lets a configured ErrorHandler suppress ThrowOnError entirely.
        if (!throwOnMissingVariable)
        {
            evalOptions.ErrorHandler = (expression, exception) =>
            {
                Logger.Warning($"TextBinder: expression '{expression}' failed to evaluate, degrading to empty text", exception);
                return string.Empty;
            };
        }

        var runs = container.Descendants<Run>().ToList();
        foreach (var run in runs)
        {
            var textElement = run.Elements<WText>().FirstOrDefault();
            if (textElement == null || string.IsNullOrEmpty(textElement.Text))
                continue;
            if (!textElement.Text.Contains("${"))
                continue;
            // Skip image expressions — handled separately
            if (textElement.Text.Contains("word.Image"))
                continue;

            var result = EvaluateText(textElement.Text, data, evalOptions, throwOnMissingVariable);
            if (result != textElement.Text)
            {
                textElement.Text = result;
                textElement.Space = SpaceProcessingModeValues.Preserve;
            }
        }
    }

    private static string EvaluateText(string text, Dictionary<string, object> data, DollarSignOptions evalOptions, bool throwOnMissingVariable)
    {
        try
        {
            var result = DollarSign.Eval(text, data, evalOptions);
            if (result != null && result.Contains("[ERROR:"))
            {
                Logger.Debug($"TextBinder: Error in evaluation, preserving original: {text}");
                return text;
            }
            return result ?? text;
        }
        catch (Exception ex) when (throwOnMissingVariable)
        {
            throw new BindingExpressionException($"Failed to evaluate expression '{text}'", ex);
        }
        catch (Exception ex)
        {
            Logger.Debug($"TextBinder: Failed to evaluate: {text} — {ex.Message}");
            return text;
        }
    }
}
