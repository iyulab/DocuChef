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
    /// Creates an xlsx with a list template row.
    /// The row at <paramref name="templateRow"/> is given a named range matching <paramref name="rangeName"/>.
    /// Cells in that row use {{Items.PropertyName}} notation.
    /// </summary>
    public static MemoryStream CreateXlsxWithNamedRange(
        string rangeName, int templateRow, params (int col, string text)[] cells)
    {
        var stream = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("Sheet1");
            int minCol = cells.Min(c => c.col);
            int maxCol = cells.Max(c => c.col);

            foreach (var (col, text) in cells)
                ws.Cell(templateRow, col).Value = text;

            // Define named range covering the template row
            var range = ws.Range(templateRow, minCol, templateRow, maxCol);
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
    /// Returns the number of rows that have any non-empty cell in the first column.
    /// </summary>
    public static int CountNonEmptyRows(Stream stream)
    {
        stream.Position = 0;
        using var wb = new XLWorkbook(stream);
        var ws = wb.Worksheets.First();
        return ws.RowsUsed().Count();
    }
}
