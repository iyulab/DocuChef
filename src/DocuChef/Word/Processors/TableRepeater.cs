using System.Collections;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using WText = DocumentFormat.OpenXml.Wordprocessing.Text;

namespace DocuChef.Word.Processors;

public static class TableRepeater
{
    // Matches ${CollectionName[].Property} — empty brackets = array iteration
    private static readonly Regex ArrayExpressionPattern =
        new(@"\$\{(\w+)\[\]\.(\w+)\}", RegexOptions.Compiled);

    public static void ProcessTables(OpenXmlElement container, Dictionary<string, object> data)
    {
        var tables = container.Descendants<Table>().ToList();
        foreach (var table in tables)
            ProcessTable(table, data);
    }

    private static void ProcessTable(Table table, Dictionary<string, object> data)
    {
        var rows = table.Elements<TableRow>().ToList();
        for (int i = rows.Count - 1; i >= 0; i--)
        {
            var row = rows[i];
            var match = ArrayExpressionPattern.Match(row.InnerText);
            if (!match.Success) continue;

            var collectionName = match.Groups[1].Value;
            if (!data.TryGetValue(collectionName, out var collectionObj) ||
                collectionObj is not IEnumerable collection)
            {
                row.Remove();
                continue;
            }

            var items = collection.Cast<object>().ToList();
            if (items.Count == 0) { row.Remove(); continue; }

            var insertAfter = row;
            for (int idx = 0; idx < items.Count; idx++)
            {
                var newRow = (TableRow)row.CloneNode(true);
                // Replace ${Collection[].Prop} with ${Collection[idx].Prop}
                foreach (var text in newRow.Descendants<WText>())
                {
                    if (text.Text != null && text.Text.Contains($"${{{collectionName}[]."))
                        text.Text = text.Text.Replace($"{collectionName}[].", $"{collectionName}[{idx}].");
                }
                table.InsertAfter(newRow, insertAfter);
                insertAfter = newRow;
            }
            row.Remove(); // Remove template row
        }
    }
}
