namespace TPJ.Excel.Models;

/// <summary>
/// Helper class to represent a row in an excel file when using EPPlus.
/// </summary>
public record EPPlusRow
{
    /// <summary>
    /// Data to be written to the cells in the row. Each item in the collection represents a cell in the row, and the order of the items corresponds to the column order in the Excel worksheet. The Value property of each EPPlusData instance will be written to the corresponding cell, 
    /// and any specified formatting (DateFormat or NumberFormat) will be applied as needed.
    /// </summary>
    public required IEnumerable<EPPlusData> Data { get; set; }
}
