namespace TPJ.Excel.Models;

/// <summary>
/// Typed column mapping used by the EPPlus helper overloads.
/// </summary>
/// <typeparam name="T">Row type</typeparam>
public record EPPlusColumn<T>
{
    /// <summary>
    /// Gets or sets the header text for the column.
    /// </summary>
    public required string Header { get; set; }

    /// <summary>
    /// Gets or sets the function used to get the cell value from the row.
    /// </summary>
    public required Func<T, object?> ValueSelector { get; set; }

    /// <summary>
    /// Gets or sets the Excel date format string used for the column.
    /// </summary>
    public string? DateFormat { get; set; }

    /// <summary>
    /// Gets or sets the Excel number format string used for the column.
    /// </summary>
    public string? NumberFormat { get; set; }
}
