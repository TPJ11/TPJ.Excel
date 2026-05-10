using OfficeOpenXml;

namespace TPJ.Excel;

/// <summary>
/// Helper class to keep track of the current row and column when working with an excel worksheet. 
/// This is useful when you want to write to a cell and then move to the next cell without 
/// having to keep track of the row and column values yourself.
/// </summary>
public class ExcelWorksheetExtension
{
    /// <summary>
    /// Gets or sets the worksheet associated with the current operation.
    /// </summary>
    /// <remarks>Use this property to access or modify the underlying Excel worksheet. Changes to this
    /// property affect which worksheet is targeted by subsequent operations.</remarks>
    public required ExcelWorksheet Worksheet { get; set; }

    /// <summary>
    /// Gets or sets the row index associated with this instance.
    /// </summary>
    public int Row { get; set; } = 1;

    /// <summary>
    /// Gets or sets the zero-based column index associated with this element.
    /// </summary>
    public int Column { get; set; }

    /// <summary>
    /// Set the row and column back to default (A)
    /// </summary>
    public void Reset()
    {
        Row = 1;
        Column = 0;
    }

    /// <summary>
    /// Adds 1 to the current row value
    /// </summary>
    /// <param name="resetColumn">If [True] the column value will be reset back to 0 else [False] the column value remains unchanged</param>
    public void NextRow(bool resetColumn = true)
    {
        Row++;
        if (resetColumn) Column = 0;
    }

    /// <summary>
    /// Adds 1 to the current column value
    /// </summary>
    public void NextColumn() => Column++;

    /// <summary>
    /// Removes 1 from the current row value
    /// </summary>
    /// <param name="resetColumn">If [True] the column value will be reset back to 0 else [False] the column value remains unchanged</param>
    public void PreviousRow(bool resetColumn = true)
    {
        if (Row == 1)
            return;

        Row--;
        if (resetColumn) Column = 0;
    }

    /// <summary>
    /// Removes 1 from the current column value
    /// </summary>
    public void PreviousColumn()
    {
        if (Column == 0) return;
        Column--;
    }

    /// <summary>
    /// Gets the current cell based on the value of row and column within this helper
    /// </summary>
    /// <returns>Currently 'selected' cell</returns>
    public ExcelRange CurrentCell() => Worksheet.Cells[$"{EPPlusHelper.GetColumnName(Column)}{Row}"];

    /// <summary>
    /// Writes a value to the current cell with optional formatting.
    /// </summary>
    /// <param name="value">Value to write</param>
    /// <param name="dateFormat">Optional Excel date format string</param>
    /// <param name="numberFormat">Optional Excel number format string</param>
    /// <param name="bold">True to make the cell bold</param>
    /// <returns>The current worksheet helper</returns>
    public ExcelWorksheetExtension Write(object? value, string? dateFormat = null, string? numberFormat = null, bool bold = false)
    {
        var cell = CurrentCell();
        cell.Value = value;

        if (bold)
            cell.Style.Font.Bold = true;

        if (!string.IsNullOrWhiteSpace(dateFormat))
            cell.Style.Numberformat.Format = dateFormat;
        else if (!string.IsNullOrWhiteSpace(numberFormat))
            cell.Style.Numberformat.Format = numberFormat;

        return this;
    }

    /// <summary>
    /// Writes a value to the current cell and then moves to the next column.
    /// </summary>
    /// <param name="value">Value to write</param>
    /// <param name="dateFormat">Optional Excel date format string</param>
    /// <param name="numberFormat">Optional Excel number format string</param>
    /// <param name="bold">True to make the cell bold</param>
    /// <returns>The current worksheet helper</returns>
    public ExcelWorksheetExtension WriteNext(object? value, string? dateFormat = null, string? numberFormat = null, bool bold = false)
    {
        Write(value, dateFormat, numberFormat, bold);
        NextColumn();
        return this;
    }

    /// <summary>
    /// Writes a header row in bold and then moves to the next row.
    /// </summary>
    /// <param name="headers">Header values</param>
    /// <returns>The current worksheet helper</returns>
    public ExcelWorksheetExtension WriteHeader(params string[] headers)
    {
        ArgumentNullException.ThrowIfNull(headers);

        foreach (var header in headers)
            WriteNext(header, bold: true);

        NextRow();
        return this;
    }

    /// <summary>
    /// Writes a row of values and then moves to the next row.
    /// </summary>
    /// <param name="values">Values to write</param>
    /// <returns>The current worksheet helper</returns>
    public ExcelWorksheetExtension WriteRow(params object?[] values)
    {
        ArgumentNullException.ThrowIfNull(values);

        foreach (var value in values)
            WriteNext(value);

        NextRow();
        return this;
    }
}