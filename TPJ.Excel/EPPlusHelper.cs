using OfficeOpenXml;
using TPJ.Excel.Models;

namespace TPJ.Excel;

/// <summary>
/// Helper class for working with EPPlus to create excel files. 
/// This class provides methods to create excel files in memory or on disk, add worksheets,
/// and write data to cells with optional formatting. It abstracts away the complexities of 
/// working directly with the EPPlus library, allowing you to easily generate Excel documents from your data models.
/// </summary>
public static class EPPlusHelper
{
    /// <summary>
    /// Creates a typed column mapping for use with the typed EPPlus helper overloads.
    /// </summary>
    /// <typeparam name="T">Row type</typeparam>
    /// <param name="header">Column header</param>
    /// <param name="valueSelector">Function used to get the cell value from the row</param>
    /// <param name="dateFormat">Optional Excel date format string</param>
    /// <param name="numberFormat">Optional Excel number format string</param>
    /// <returns>Typed column mapping</returns>
    public static EPPlusColumn<T> Column<T>(string header, Func<T, object?> valueSelector, string? dateFormat = null, string? numberFormat = null)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(header);
        ArgumentNullException.ThrowIfNull(valueSelector);

        return new EPPlusColumn<T>
        {
            Header = header,
            ValueSelector = valueSelector,
            DateFormat = dateFormat,
            NumberFormat = numberFormat
        };
    }

    /// <summary>
    /// Gets the Excel column name (e.g., "A", "B", ..., "Z", "AA", etc.) for a given zero-based column index.
    /// </summary>
    /// <param name="columnIndex">Column index (zero-based).</param>
    /// <returns>The Excel column name.</returns>
    public static string GetColumnName(int columnIndex)
    {
        if(columnIndex < 0) 
            throw new ArgumentOutOfRangeException(nameof(columnIndex), "Column index must be a non-negative integer.");

        var dividend = columnIndex + 1;
        var columnName = string.Empty;

        while (dividend > 0)
        {
            var modulo = (dividend - 1) % 26;
            columnName = Convert.ToChar(65 + modulo) + columnName;
            dividend = (dividend - modulo) / 26;
        }

        return columnName;
    }

    /// <summary>
    /// Gets a cell based on the passed in column and row value e.g. 0, 1 will return cell A and and 5, 3 will return cell BF
    /// </summary>
    /// <param name="ws">Worksheet to get cell from</param>
    /// <param name="column">Column to get (0 base)</param>
    /// <param name="row">Row to get (1 base)</param>
    /// <returns>Cell</returns>
    public static ExcelRange Cell(this ExcelWorksheet ws, int column, int row) => ws.Cells[$"{GetColumnName(column)}{row}"];

    /// <summary>
    /// Adds a worksheet to the workbook with the given name and returns an ExcelWorksheetExtension which contains the created worksheet
    /// </summary>
    /// <param name="workbook">Workbook to add the worksheet to</param>
    /// <param name="name">Name of the worksheet</param>
    /// <returns>ExcelWorksheetExtension containing the created worksheet</returns>
    public static ExcelWorksheetExtension AddWorksheet(this ExcelWorkbook workbook, string name)
    {
        ArgumentNullException.ThrowIfNull(workbook);

        var worksheetName = ValidateWorksheetName(name);
        if (workbook.Worksheets.Any(x => string.Equals(x.Name, worksheetName, StringComparison.OrdinalIgnoreCase)))
            throw new ArgumentException($"A worksheet named '{worksheetName}' has already been added.", nameof(name));

        var ws = workbook.Worksheets.Add(worksheetName);

        return new ExcelWorksheetExtension()
        {
            Worksheet = ws
        };
    }

    /// <summary>
    /// Creates a worksheet model from typed column mappings.
    /// </summary>
    /// <typeparam name="T">Row type</typeparam>
    /// <param name="rows">Rows to write</param>
    /// <param name="worksheetName">Worksheet name</param>
    /// <param name="columns">Typed column mappings</param>
    /// <returns>Worksheet model</returns>
    public static EPPlusWorksheet Worksheet<T>(IEnumerable<T> rows, string worksheetName, params EPPlusColumn<T>[] columns) =>
        Worksheet(rows, worksheetName, (IEnumerable<EPPlusColumn<T>>)columns);

    /// <summary>
    /// Creates a worksheet model from a typed column builder.
    /// </summary>
    /// <typeparam name="T">Row type</typeparam>
    /// <param name="rows">Rows to write</param>
    /// <param name="worksheetName">Worksheet name</param>
    /// <param name="configureColumns">Column builder configuration</param>
    /// <returns>Worksheet model</returns>
    public static EPPlusWorksheet Worksheet<T>(IEnumerable<T> rows, string worksheetName, Action<EPPlusColumnBuilder<T>> configureColumns) =>
        Worksheet(rows, worksheetName, CreateColumns(configureColumns));

    /// <summary>
    /// Creates a worksheet model from typed column mappings.
    /// </summary>
    /// <typeparam name="T">Row type</typeparam>
    /// <param name="rows">Rows to write</param>
    /// <param name="worksheetName">Worksheet name</param>
    /// <param name="columns">Typed column mappings</param>
    /// <returns>Worksheet model</returns>
    public static EPPlusWorksheet Worksheet<T>(IEnumerable<T> rows, string worksheetName, IEnumerable<EPPlusColumn<T>> columns) =>
        CreateWorksheet(rows, worksheetName, columns);

    /// <summary>
    /// Creates an excel file in memory using typed column mappings.
    /// </summary>
    /// <typeparam name="T">Row type</typeparam>
    /// <param name="rows">Rows to write</param>
    /// <param name="worksheetName">Worksheet name</param>
    /// <param name="columns">Typed column mappings</param>
    /// <returns>Bytes of the created excel worksheet</returns>
    public static byte[] Create<T>(IEnumerable<T> rows, string worksheetName, params EPPlusColumn<T>[] columns) =>
        Create(rows, worksheetName, (IEnumerable<EPPlusColumn<T>>)columns);

    /// <summary>
    /// Creates an excel file in memory using a typed column builder.
    /// </summary>
    /// <typeparam name="T">Row type</typeparam>
    /// <param name="rows">Rows to write</param>
    /// <param name="worksheetName">Worksheet name</param>
    /// <param name="configureColumns">Column builder configuration</param>
    /// <returns>Bytes of the created excel worksheet</returns>
    public static byte[] Create<T>(IEnumerable<T> rows, string worksheetName, Action<EPPlusColumnBuilder<T>> configureColumns) =>
        Create(rows, worksheetName, CreateColumns(configureColumns));

    /// <summary>
    /// Creates an excel file in memory using typed column mappings.
    /// </summary>
    /// <typeparam name="T">Row type</typeparam>
    /// <param name="rows">Rows to write</param>
    /// <param name="worksheetName">Worksheet name</param>
    /// <param name="columns">Typed column mappings</param>
    /// <returns>Bytes of the created excel worksheet</returns>
    public static byte[] Create<T>(IEnumerable<T> rows, string worksheetName, IEnumerable<EPPlusColumn<T>> columns) =>
        Create(CreateWorksheet(rows, worksheetName, columns));

    /// <summary>
    /// Creates an excel file to disk using typed column mappings.
    /// </summary>
    /// <typeparam name="T">Row type</typeparam>
    /// <param name="rows">Rows to write</param>
    /// <param name="worksheetName">Worksheet name</param>
    /// <param name="filePath">Save excel file path e.g. C:\example.xlsx</param>
    /// <param name="columns">Typed column mappings</param>
    public static void Create<T>(IEnumerable<T> rows, string worksheetName, string filePath, params EPPlusColumn<T>[] columns) =>
        Create(rows, worksheetName, filePath, (IEnumerable<EPPlusColumn<T>>)columns);

    /// <summary>
    /// Creates an excel file to disk using a typed column builder.
    /// </summary>
    /// <typeparam name="T">Row type</typeparam>
    /// <param name="rows">Rows to write</param>
    /// <param name="worksheetName">Worksheet name</param>
    /// <param name="filePath">Save excel file path e.g. C:\example.xlsx</param>
    /// <param name="configureColumns">Column builder configuration</param>
    public static void Create<T>(IEnumerable<T> rows, string worksheetName, string filePath, Action<EPPlusColumnBuilder<T>> configureColumns) =>
        Create(rows, worksheetName, filePath, CreateColumns(configureColumns));

    /// <summary>
    /// Creates an excel file to disk using typed column mappings.
    /// </summary>
    /// <typeparam name="T">Row type</typeparam>
    /// <param name="rows">Rows to write</param>
    /// <param name="worksheetName">Worksheet name</param>
    /// <param name="filePath">Save excel file path e.g. C:\example.xlsx</param>
    /// <param name="columns">Typed column mappings</param>
    public static void Create<T>(IEnumerable<T> rows, string worksheetName, string filePath, IEnumerable<EPPlusColumn<T>> columns) =>
        Create(CreateWorksheet(rows, worksheetName, columns), filePath);

    /// <summary>
    /// Creates an excel file in memory containing the given worksheet
    /// </summary>
    /// <param name="worksheet">Worksheet to create</param>
    /// <returns>Bytes of the created excel worksheet</returns>
    public static byte[] Create(EPPlusWorksheet worksheet) => Create([worksheet]);

    /// <summary>
    /// Creates an excel file in memory containing the given worksheets
    /// </summary>
    /// <param name="worksheets">Worksheets to create</param>
    /// <returns>Bytes of the created excel worksheet</returns>
    public static byte[] Create(IEnumerable<EPPlusWorksheet> worksheets)
    {
        using var ms = new MemoryStream();
        using var excelPackage = CreateExcel(worksheets);
        excelPackage.SaveAs(ms);
        return ms.ToArray();
    }

    /// <summary>
    /// Creates an excel file to disk containing the given worksheets
    /// </summary>
    /// <param name="worksheet">Worksheet to create</param>
    /// <param name="filePath">Save excel file path e.g. C:\example.xlsx</param>
    public static void Create(EPPlusWorksheet worksheet, string filePath) => Create([worksheet], filePath);
    
    /// <summary>
    /// Creates an excel file to disk containing the given worksheets
    /// </summary>
    /// <param name="worksheets">Worksheets to create</param>
    /// <param name="filePath">Save excel file path e.g. C:\example.xlsx</param>
    public static void Create(IEnumerable<EPPlusWorksheet> worksheets, string filePath)
    {
        using var excelPackage = CreateExcel(worksheets);
        excelPackage.SaveAs(new FileInfo(filePath));
    }

    /// <summary>
    /// Creates the excel package with the given worksheets
    /// </summary>
    /// <param name="worksheets">Worksheets to add to excel package</param>
    /// <returns>Excel package</returns>
    private static ExcelPackage CreateExcel(IEnumerable<EPPlusWorksheet> worksheets)
    {
        ArgumentNullException.ThrowIfNull(worksheets);

        var worksheetList = worksheets.ToArray();
        var excelPackage = new ExcelPackage();
        var worksheetNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        foreach (var worksheet in worksheetList)
        {
            ArgumentNullException.ThrowIfNull(worksheet);

            var worksheetName = ValidateWorksheetName(worksheet.Name);
            if (!worksheetNames.Add(worksheetName))
                throw new ArgumentException($"A worksheet named '{worksheetName}' has already been added.", nameof(worksheets));

            var ws = excelPackage.Workbook.Worksheets.Add(worksheetName);

            var startingRow = 1;
            if (worksheet.Headers is not null && worksheet.Headers.Any())
            {
                ws.AddHeaders(worksheet.Headers);
                startingRow = 2;
            }

            ws.AddData(worksheet.Rows, startingRow);
        }

        return excelPackage;
    }

    /// <summary>
    /// Creates a worksheet model from typed column mappings.
    /// </summary>
    /// <typeparam name="T">Row type</typeparam>
    /// <param name="rows">Rows to write</param>
    /// <param name="worksheetName">Worksheet name</param>
    /// <param name="columns">Typed column mappings</param>
    /// <returns>Worksheet model</returns>
    private static EPPlusWorksheet CreateWorksheet<T>(IEnumerable<T> rows, string worksheetName, IEnumerable<EPPlusColumn<T>> columns)
    {
        ArgumentNullException.ThrowIfNull(rows);
        ArgumentNullException.ThrowIfNull(columns);

        worksheetName = ValidateWorksheetName(worksheetName);

        var mappedColumns = columns.ToArray();

        if (mappedColumns.Length == 0)
            throw new ArgumentException("At least one column must be supplied.", nameof(columns));

        return new EPPlusWorksheet
        {
            Name = worksheetName,
            Headers = mappedColumns.Select(x => x.Header),
            Rows = rows.Select(row => new EPPlusRow
            {
                Data = mappedColumns.Select(column => new EPPlusData
                {
                    Value = row is null ? null : column.ValueSelector(row),
                    DateFormat = column.DateFormat,
                    NumberFormat = column.NumberFormat
                })
            })
        };
    }

    /// <summary>
    /// Creates typed column mappings from a builder callback.
    /// </summary>
    /// <typeparam name="T">Row type</typeparam>
    /// <param name="configureColumns">Column builder configuration</param>
    /// <returns>Typed column mappings</returns>
    private static IEnumerable<EPPlusColumn<T>> CreateColumns<T>(Action<EPPlusColumnBuilder<T>> configureColumns)
    {
        ArgumentNullException.ThrowIfNull(configureColumns);

        var builder = new EPPlusColumnBuilder<T>();
        configureColumns(builder);
        return builder.Columns;
    }

    /// <summary>
    /// Validates and normalizes a worksheet name.
    /// </summary>
    /// <param name="worksheetName">Worksheet name</param>
    /// <returns>Validated worksheet name</returns>
    private static string ValidateWorksheetName(string worksheetName)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(worksheetName);

        var value = worksheetName.Trim();

        if (value.Length > 31)
            throw new ArgumentException("Worksheet name must be 31 characters or fewer.", nameof(worksheetName));

        if (value.IndexOfAny(['[', ']', ':', '*', '?', '/', '\\']) >= 0)
            throw new ArgumentException("Worksheet name contains invalid characters.", nameof(worksheetName));

        return value;
    }

    /// <summary>
    /// Adds the passed in headers to the worksheet on the 
    /// first row starting at 'A' in bold
    /// </summary>
    /// <param name="worksheet">Worksheet to add headers to</param>
    /// <param name="headers">Headers to add</param>
    private static void AddHeaders(this ExcelWorksheet worksheet, IEnumerable<string> headers)
    {
        var column = 0;
        foreach (var header in headers)
        {
            worksheet.Cell(column, 1).Value = header;
            worksheet.Cell(column, 1).Style.Font.Bold = true;
            column++;
        }
    }

    /// <summary>
    /// Adds the passed in data to the worksheet starting 
    /// at the passed in row and on column 0 
    /// </summary>
    /// <param name="worksheet">Worksheet to add data to</param>
    /// <param name="rows">Rows of data to add</param>
    /// <param name="rowNumber">Starting row to add the data onto</param>
    private static void AddData(this ExcelWorksheet worksheet, IEnumerable<EPPlusRow> rows, int rowNumber)
    {
        foreach (var row in rows)
        {
            var columnNumber = 0;

            foreach (var data in row.Data)
            {
                var cell = worksheet.Cell(columnNumber, rowNumber);
                cell.Value = data.Value;

                if (!string.IsNullOrWhiteSpace(data.DateFormat))
                    cell.Style.Numberformat.Format = data.DateFormat;
                else if (!string.IsNullOrWhiteSpace(data.NumberFormat))
                    cell.Style.Numberformat.Format = data.NumberFormat;

                columnNumber++;
            }

            rowNumber++;
        }
    }

}
