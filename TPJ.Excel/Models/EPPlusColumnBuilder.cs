namespace TPJ.Excel.Models;

/// <summary>
/// Builder used to create typed EPPlus column mappings without repeating the row type.
/// </summary>
/// <typeparam name="T">Row type</typeparam>
public class EPPlusColumnBuilder<T>
{
    private readonly List<EPPlusColumn<T>> columns = [];

    /// <summary>
    /// Gets the configured columns.
    /// </summary>
    public IReadOnlyList<EPPlusColumn<T>> Columns => columns;

    /// <summary>
    /// Adds a column mapping.
    /// </summary>
    /// <param name="header">Column header</param>
    /// <param name="valueSelector">Function used to get the cell value from the row</param>
    /// <param name="dateFormat">Optional Excel date format string</param>
    /// <param name="numberFormat">Optional Excel number format string</param>
    /// <returns>The builder</returns>
    public EPPlusColumnBuilder<T> Add(string header, Func<T, object?> valueSelector, string? dateFormat = null, string? numberFormat = null)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(header);
        ArgumentNullException.ThrowIfNull(valueSelector);

        columns.Add(new EPPlusColumn<T>
        {
            Header = header,
            ValueSelector = valueSelector,
            DateFormat = dateFormat,
            NumberFormat = numberFormat
        });

        return this;
    }
}
