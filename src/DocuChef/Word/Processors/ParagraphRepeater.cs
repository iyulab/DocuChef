using System.Collections;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using WText = DocumentFormat.OpenXml.Wordprocessing.Text;

namespace DocuChef.Word.Processors;

/// <summary>
/// Expands #foreach: CollectionName / #end paragraph blocks by cloning
/// the template paragraphs for each item in the collection and rewriting
/// variable references from ${Property} to ${CollectionName[index].Property}.
/// </summary>
public static class ParagraphRepeater
{
    private static readonly Regex ForeachPattern =
        new(@"^#foreach\s*:\s*(\w+)$", RegexOptions.Compiled);

    private static readonly Regex VariablePattern =
        new(@"\$\{(\w+)\}", RegexOptions.Compiled);

    /// <summary>
    /// Processes all #foreach/#end blocks in the given container.
    /// </summary>
    public static void ProcessParagraphs(OpenXmlElement container, Dictionary<string, object> data)
    {
        // Process from bottom to top so index positions remain stable
        while (true)
        {
            var paragraphs = container.Elements<Paragraph>().ToList();
            var block = FindForeachBlock(paragraphs);
            if (block == null) break;

            var (foreachIndex, endIndex, collectionName) = block.Value;

            // Get collection from data
            IList<object>? items = null;
            if (data.TryGetValue(collectionName, out var collectionObj) &&
                collectionObj is IEnumerable collection)
            {
                items = collection.Cast<object>().ToList();
            }

            // Extract template paragraphs (between #foreach and #end, exclusive)
            var templateParagraphs = new List<Paragraph>();
            for (int i = foreachIndex + 1; i < endIndex; i++)
                templateParagraphs.Add(paragraphs[i]);

            // Determine insertion point: insert before #foreach paragraph
            var foreachParagraph = paragraphs[foreachIndex];
            var endParagraph = paragraphs[endIndex];

            // Clone and rewrite for each item
            if (items != null && items.Count > 0)
            {
                OpenXmlElement insertBefore = foreachParagraph;
                for (int idx = 0; idx < items.Count; idx++)
                {
                    foreach (var templatePara in templateParagraphs)
                    {
                        var cloned = (Paragraph)templatePara.CloneNode(true);
                        RewriteVariables(cloned, collectionName, idx, data);
                        container.InsertBefore(cloned, foreachParagraph);
                    }
                }
            }

            // Remove #foreach, template paragraphs, and #end
            foreach (var templatePara in templateParagraphs)
                templatePara.Remove();
            foreachParagraph.Remove();
            endParagraph.Remove();
        }
    }

    private static (int foreachIndex, int endIndex, string collectionName)?
        FindForeachBlock(List<Paragraph> paragraphs)
    {
        for (int i = 0; i < paragraphs.Count; i++)
        {
            var text = paragraphs[i].InnerText.Trim();
            var match = ForeachPattern.Match(text);
            if (!match.Success) continue;

            var collectionName = match.Groups[1].Value;

            // Find matching #end
            for (int j = i + 1; j < paragraphs.Count; j++)
            {
                if (paragraphs[j].InnerText.Trim() == "#end")
                    return (i, j, collectionName);
            }
        }

        return null;
    }

    /// <summary>
    /// Rewrites ${Property} to ${CollectionName[index].Property} for variables
    /// that are NOT already top-level keys in the data dictionary.
    /// </summary>
    private static void RewriteVariables(Paragraph paragraph, string collectionName, int index,
        Dictionary<string, object> data)
    {
        foreach (var textElement in paragraph.Descendants<WText>())
        {
            if (textElement.Text == null || !textElement.Text.Contains("${"))
                continue;

            textElement.Text = VariablePattern.Replace(textElement.Text, match =>
            {
                var variableName = match.Groups[1].Value;
                // If the variable is a top-level data key, leave it untouched
                if (data.ContainsKey(variableName))
                    return match.Value;
                // Otherwise, rewrite to indexed collection access
                return $"${{{collectionName}[{index}].{variableName}}}";
            });
        }
    }
}
