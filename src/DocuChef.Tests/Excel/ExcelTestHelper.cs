using ClosedXML.Excel;

namespace DocuChef.Tests.Excel;

/// <summary>
/// Creates and reads .xlsx documents programmatically for testing.
/// </summary>
public static class ExcelTestHelper
{
    /// <summary>
    /// Creates an xlsx with one sheet named "Sheet1".
    /// cells: (row, col, text) — 1-based indexing.
    /// </summary>
    public static MemoryStream CreateXlsx(params (int row, int col, string text)[] cells)
    {
        var stream = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("Sheet1");
            foreach (var (row, col, text) in cells)
                ws.Cell(row, col).Value = text;
            wb.SaveAs(stream);
        }
        stream.Position = 0;
        return stream;
    }

    /// <summary>
    /// Creates an xlsx with a ClosedXML.Report "flat table" list template.
    /// The named range is defined as <paramref name="rangeName"/> (must match the bound
    /// variable name); cells reference the fixed <c>item</c> keyword, e.g. {{item.Name}} —
    /// NOT the range/variable name itself (verified against closedxml.io Flat-tables docs;
    /// {{rangeName.Property}} resolves as "Unknown identifier").
    /// A flat-table range needs a leftmost service column and a bottom service row in
    /// addition to the data row, so <paramref name="cells"/> must use columns ≥ 2 — the
    /// range then reserves column 1 and the row below <paramref name="templateRow"/> as
    /// the service column/row.
    /// </summary>
    public static MemoryStream CreateXlsxWithNamedRange(
        string rangeName, int templateRow, params (int col, string text)[] cells)
    {
        var stream = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("Sheet1");
            int maxCol = cells.Max(c => c.col);

            foreach (var (col, text) in cells)
                ws.Cell(templateRow, col).Value = text;

            // Column 1 = service column, templateRow + 1 = service row.
            var range = ws.Range(templateRow, 1, templateRow + 1, maxCol);
            wb.DefinedNames.Add(rangeName, range);

            wb.SaveAs(stream);
        }
        stream.Position = 0;
        return stream;
    }

    /// <summary>
    /// Reads the string value of a cell in Sheet1 from an xlsx stream.
    /// </summary>
    public static string ReadCellValue(Stream stream, int row, int col)
    {
        stream.Position = 0;
        using var wb = new XLWorkbook(stream);
        var ws = wb.Worksheets.First();
        return ws.Cell(row, col).GetString();
    }

    /// <summary>
    /// Returns the number of rows that have any non-empty cell anywhere in the row
    /// (ClosedXML's "used rows").
    /// </summary>
    public static int CountNonEmptyRows(Stream stream)
    {
        stream.Position = 0;
        using var wb = new XLWorkbook(stream);
        var ws = wb.Worksheets.First();
        return ws.RowsUsed().Count();
    }
}
