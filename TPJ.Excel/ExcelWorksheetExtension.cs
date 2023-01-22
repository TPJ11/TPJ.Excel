using OfficeOpenXml;

namespace TPJ.Excel;

public class ExcelWorksheetExtension
{
    public ExcelWorksheet Worksheet { get; set; }
    public int Row { get; set; } = 1;
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
        Row--;
        if (resetColumn) Column = 0;
    }

    /// <summary>
    /// Removes 1 from the current column value
    /// </summary>
    public void PreviousColumn() => Column--;

    /// <summary>
    /// Gets the current cell based on the value of row and column within this helper
    /// </summary>
    /// <param name="ws">Worksheet to get cell from</param>
    /// <returns>Currently 'selected' cell</returns>
    public ExcelRange Cell() => Worksheet.Cells[$"{EPPlusHelper.Columns[Column]}{Row}"];
}