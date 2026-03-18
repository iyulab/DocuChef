using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using DollarSignEngine;
using WText = DocumentFormat.OpenXml.Wordprocessing.Text;

namespace DocuChef.Word.Processors;

/// <summary>
/// Evaluates ${expression} placeholders in Word document text elements
/// using DollarSignEngine for expression evaluation.
/// </summary>
public static class TextBinder
{
    private static readonly DollarSignOptions EvalOptions = new()
    {
        SupportDollarSignSyntax = true,
        ThrowOnError = false
    };

    /// <summary>
    /// Binds data to all ${expression} placeholders in the given container.
    /// </summary>
    public static void Bind(OpenXmlElement container, Dictionary<string, object> data)
    {
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

            var result = EvaluateText(textElement.Text, data);
            if (result != textElement.Text)
            {
                textElement.Text = result;
                textElement.Space = SpaceProcessingModeValues.Preserve;
            }
        }
    }

    private static string EvaluateText(string text, Dictionary<string, object> data)
    {
        try
        {
            var result = DollarSign.Eval(text, data, EvalOptions);
            if (result != null && result.Contains("[ERROR:"))
            {
                Logger.Debug($"TextBinder: Error in evaluation, preserving original: {text}");
                return text;
            }
            return result ?? text;
        }
        catch (Exception ex)
        {
            Logger.Debug($"TextBinder: Failed to evaluate: {text} — {ex.Message}");
            return text;
        }
    }
}
