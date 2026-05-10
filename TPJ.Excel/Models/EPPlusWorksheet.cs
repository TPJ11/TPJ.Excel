namespace TPJ.Excel.Models;

/// <summary>
/// Helper class to represent a worksheet in an excel file when using EPPlus. 
/// This class allows you to specify the name of the worksheet, the headers for the columns, 
/// and the rows of data to be written to the worksheet. 
/// By using this model, you can easily control the structure and content of your Excel 
/// document when using the EPPlusHelper methods.
/// </summary>
public record EPPlusWorksheet
{
    /// <summary>
    /// Gets or sets the name associated with this instance.
    /// </summary>
    public required string Name { get; set; }

    /// <summary>
    /// Gets or sets the collection of header values associated with the request or response.
    /// </summary>
    public IEnumerable<string>? Headers { get; set; }

    /// <summary>
    /// Gets or sets the collection of worksheet rows.
    /// </summary>
    public required IEnumerable<EPPlusRow> Rows { get; set; }
}
