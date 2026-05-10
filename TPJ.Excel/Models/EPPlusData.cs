namespace TPJ.Excel.Models;

/// <summary>
/// Model to represent the data to be written to an excel cell using EPPlus. 
/// This class allows you to specify the value of the cell as well as optional formatting for dates and numbers. 
/// By using this model, you can easily control how your data is displayed in the resulting Excel document when using the EPPlusHelper methods.
/// </summary>
public record EPPlusData
{
    /// <summary>
    /// Gets or sets the value associated with this instance.
    /// </summary>
    public object? Value { get; set; }

    /// <summary>
    /// Gets or sets the date format string used for parsing or formatting date values.
    /// </summary>
    /// <remarks>The format string should follow standard .NET date and time format patterns. If not set, a
    /// default format may be used depending on the context.</remarks>
    public string? DateFormat { get; set; }

    /// <summary>
    /// Gets or sets the format string used to display numeric values.
    /// </summary>
    /// <remarks>The format string should follow standard .NET numeric format specifiers. If null or empty,
    /// the default formatting is applied.</remarks>
    public string? NumberFormat { get; set; }
}
