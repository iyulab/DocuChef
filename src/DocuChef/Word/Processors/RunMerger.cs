using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using WText = DocumentFormat.OpenXml.Wordprocessing.Text;

namespace DocuChef.Word.Processors;

/// <summary>
/// Merges expression runs that Word has split across multiple Run elements.
/// Word editors often split text like ${Variable} into separate Runs due to
/// spell-check, formatting changes, or editing history. This class consolidates
/// them back so template processing can find complete expressions.
/// </summary>
public static class RunMerger
{
    private static readonly Regex ExpressionPattern = new(@"\$\{[^}]+\}", RegexOptions.Compiled);

    /// <summary>
    /// Scans all paragraphs in the given container and merges runs where
    /// a ${...} expression is split across multiple Run elements.
    /// </summary>
    public static void MergeExpressionRuns(OpenXmlElement container)
    {
        foreach (var paragraph in container.Descendants<Paragraph>())
        {
            MergeParagraphRuns(paragraph);
        }
    }

    private static void MergeParagraphRuns(Paragraph paragraph)
    {
        var runs = paragraph.Elements<Run>().ToList();
        if (runs.Count <= 1)
            return;

        // Build the full paragraph text from runs
        var fullText = new StringBuilder();
        foreach (var run in runs)
        {
            fullText.Append(run.InnerText);
        }

        var paragraphText = fullText.ToString();

        // Check if there are any expressions in the full text
        if (!ExpressionPattern.IsMatch(paragraphText))
            return;

        // Check if each expression is already contained within a single run.
        // If so, no merging is needed.
        bool needsMerge = false;
        foreach (Match match in ExpressionPattern.Matches(paragraphText))
        {
            if (!IsExpressionInSingleRun(runs, match.Value))
            {
                needsMerge = true;
                break;
            }
        }

        if (!needsMerge)
            return;

        // Put all text into the first run, clear the rest
        var firstRun = runs[0];
        var textElement = firstRun.GetFirstChild<WText>();
        if (textElement == null)
        {
            textElement = new WText();
            firstRun.Append(textElement);
        }

        textElement.Text = paragraphText;
        textElement.Space = SpaceProcessingModeValues.Preserve;

        // Clear text from all subsequent runs
        for (int i = 1; i < runs.Count; i++)
        {
            var runText = runs[i].GetFirstChild<WText>();
            if (runText != null)
            {
                runText.Text = string.Empty;
            }
            else
            {
                // If there are multiple Text elements, clear them all
                foreach (var t in runs[i].Elements<WText>().ToList())
                {
                    t.Text = string.Empty;
                }
            }
        }

        // Remove ProofError elements from the paragraph
        foreach (var proofErr in paragraph.Elements<ProofError>().ToList())
        {
            proofErr.Remove();
        }
    }

    private static bool IsExpressionInSingleRun(List<Run> runs, string expression)
    {
        foreach (var run in runs)
        {
            if (run.InnerText.Contains(expression))
                return true;
        }
        return false;
    }
}
