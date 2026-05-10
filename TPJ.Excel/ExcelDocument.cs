using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Data;
using System.Globalization;

namespace TPJ.Excel;

/// <summary>
/// Helper class to create excel documents using OpenXML. 
/// This class can be used to create excel documents either by writing to a file, 
/// or by writing to a MemoryStream and returning the bytes of the created excel document.
/// </summary>
public static class ExcelDocument
{
    /// <summary>
    /// Create an excel document using the 
    /// given list of data and the excel document path
    /// </summary>
    /// <typeparam name="T">Object Type</typeparam>
    /// <param name="list">Data</param>
    /// <param name="filePath">Save excel file path e.g. C:\example.xlsx</param>
    public static void Create<T>(List<T> list, string filePath) => Create((IEnumerable<T>)list, filePath);

    /// <summary>
    /// Create an excel document using the
    /// given data enumerable and the excel document path
    /// </summary>
    /// <typeparam name="T">Object Type</typeparam>
    /// <param name="list">Data</param>
    /// <param name="filePath">Save excel file path e.g. C:\example.xlsx</param>
    /// <param name="worksheetName">Optional worksheet name</param>
    public static void Create<T>(IEnumerable<T> list, string filePath, string? worksheetName = null)
    {
        ArgumentNullException.ThrowIfNull(list);
        ArgumentException.ThrowIfNullOrWhiteSpace(filePath);

        Create(ListToDataTable(list, worksheetName), filePath);
    }

    /// <summary>
    /// Create an excel document using the 
    /// given data table and the excel document path
    /// </summary>
    /// <param name="dt">Data Table</param>
    /// <param name="filePath">Save excel file path e.g. C:\example.xlsx</param>
    /// <param name="worksheetName">Optional worksheet name</param>
    public static void Create(DataTable dt, string filePath, string? worksheetName = null)
    {
        ArgumentNullException.ThrowIfNull(dt);
        ArgumentException.ThrowIfNullOrWhiteSpace(filePath);

        var ds = new DataSet();
        ds.Tables.Add(CloneTable(dt, worksheetName));
        Create(ds, filePath);
    }

    /// <summary>
    /// Create an excel document using the 
    /// given data set and the excel document path
    /// </summary>
    /// <param name="ds">Data Set</param>
    /// <param name="filePath">Save excel file path e.g. C:\example.xlsx</param>
    public static void Create(DataSet ds, string filePath)
    {
        ArgumentNullException.ThrowIfNull(ds);
        ArgumentException.ThrowIfNullOrWhiteSpace(filePath);

        using var document = SpreadsheetDocument.Create(filePath, SpreadsheetDocumentType.Workbook);
        WriteExcelFile(ds, document);
    }

    /// <summary>
    /// Create an excel document with the given data enumerable
    /// </summary>
    /// <typeparam name="T">Object Type</typeparam>
    /// <param name="list">Data</param>
    /// <returns>Excel file in bytes</returns>
    public static byte[] Create<T>(IEnumerable<T> list)
    {
        ArgumentNullException.ThrowIfNull(list);
        return Create(ListToDataTable(list));
    }

    /// <summary>
    /// Create an excel document using a data table
    /// </summary>
    /// <param name="dt">Data Table</param>
    /// <returns>Excel file in bytes</returns>
    public static byte[] Create(DataTable dt)
    {
        ArgumentNullException.ThrowIfNull(dt);

        var ds = new DataSet();
        ds.Tables.Add(CloneTable(dt));
        return Create(ds);
    }

    /// <summary>
    /// Create an excel document using the 
    /// given data set and the memory stream
    /// </summary>
    /// <param name="ds">Data Set</param>
    public static byte[] Create(DataSet ds)
    {
        ArgumentNullException.ThrowIfNull(ds);

        using var ms = new MemoryStream();
        using (var document = SpreadsheetDocument.Create(ms, SpreadsheetDocumentType.Workbook))
        {
            WriteExcelFile(ds, document);
        }

        return ms.ToArray();
    }

    /// <summary>
    /// Convert a list of data into a data table
    /// </summary>
    /// <typeparam name="T">Object Type</typeparam>
    /// <param name="list">Data List</param>
    /// <returns>Data table with all data within data list</returns>
    private static DataTable ListToDataTable<T>(IEnumerable<T> list, string? worksheetName = null)
    {
        var dt = new DataTable(GetWorksheetName(worksheetName));
        var itemType = typeof(T);
        var valueType = GetNullableType(itemType);

        if (IsSimpleValueType(valueType))
        {
            dt.Columns.Add(new DataColumn("Value", valueType));

            foreach (var item in list)
            {
                var row = dt.NewRow();
                row[0] = item is null ? DBNull.Value : item;
                dt.Rows.Add(row);
            }

            return dt;
        }

        var properties = itemType
            .GetProperties()
            .Where(x => x.CanRead && x.GetIndexParameters().Length == 0)
            .ToArray();

        foreach (var info in properties)
            dt.Columns.Add(new DataColumn(info.Name, GetNullableType(info.PropertyType)));

        foreach (var t in list)
        {
            var row = dt.NewRow();
            foreach (var info in properties)
                row[info.Name] = t is null ? DBNull.Value : (info.GetValue(t, null) ?? DBNull.Value);

            dt.Rows.Add(row);
        }
        return dt;
    }

    /// <summary>
    /// Gets the nullable version of the given type if exists else the type given is returned
    /// e.g. string -> string? 
    /// </summary>
    /// <param name="t">Type to find nullable version for</param>
    /// <returns>Nullable version or given type</returns>
    private static Type GetNullableType(Type t)
    {
        var returnType = t;
        if (t.IsGenericType && t.GetGenericTypeDefinition().Equals(typeof(Nullable<>)))            
            returnType = Nullable.GetUnderlyingType(t);
        
        return returnType!;
    }

    /// <summary>
    /// Checks to see if the given type is a simple scalar value that should be exported as a single column.
    /// </summary>
    /// <param name="type">Type to check</param>
    /// <returns>True if the type is a simple scalar value</returns>
    private static bool IsSimpleValueType(Type type)
    {
        type = GetNullableType(type);

        return type.IsPrimitive
            || type.IsEnum
            || type == typeof(string)
            || type == typeof(decimal)
            || type == typeof(DateTime)
            || type == typeof(DateTimeOffset)
            || type == typeof(DateOnly)
            || type == typeof(TimeOnly)
            || type == typeof(TimeSpan)
            || type == typeof(Guid);
    }

    /// <summary>
    /// Checks to see if the given type is numeric.
    /// </summary>
    /// <param name="type">Type to check</param>
    /// <returns>True if the type is numeric</returns>
    private static bool IsNumericType(Type type)
    {
        type = GetNullableType(type);

        return Type.GetTypeCode(type) switch
        {
            TypeCode.Byte => true,
            TypeCode.SByte => true,
            TypeCode.Int16 => true,
            TypeCode.UInt16 => true,
            TypeCode.Int32 => true,
            TypeCode.UInt32 => true,
            TypeCode.Int64 => true,
            TypeCode.UInt64 => true,
            TypeCode.Single => true,
            TypeCode.Double => true,
            TypeCode.Decimal => true,
            _ => false
        };
    }

    /// <summary>
    /// Gets a valid worksheet name.
    /// </summary>
    /// <param name="worksheetName">Worksheet name to validate</param>
    /// <param name="worksheetNumber">Worksheet number to use when a name is not supplied</param>
    /// <returns>Worksheet name</returns>
    private static string GetWorksheetName(string? worksheetName, uint worksheetNumber = 1)
    {
        var value = string.IsNullOrWhiteSpace(worksheetName)
            ? $"Sheet{worksheetNumber}"
            : worksheetName.Trim();

        if (value.Length > 31)
            throw new ArgumentException("Worksheet name must be 31 characters or fewer.", nameof(worksheetName));

        if (value.IndexOfAny(['[', ']', ':', '*', '?', '/', '\\']) >= 0)
            throw new ArgumentException("Worksheet name contains invalid characters.", nameof(worksheetName));

        return value;
    }

    /// <summary>
    /// Creates a copy of the given data table and applies a valid worksheet name.
    /// </summary>
    /// <param name="dt">Data table</param>
    /// <param name="worksheetName">Optional worksheet name</param>
    /// <returns>Cloned data table</returns>
    private static DataTable CloneTable(DataTable dt, string? worksheetName = null)
    {
        var clonedTable = dt.Copy();
        clonedTable.TableName = GetWorksheetName(worksheetName ?? dt.TableName);
        return clonedTable;
    }

    /// <summary>
    /// Adds the given data set into the given spreadsheet document object
    /// This function is used when creating an Excel file either writing 
    /// to a file, or writing to a MemoryStream
    /// </summary>
    /// <param name="ds">Data set to add to excel</param>
    /// <param name="spreadsheet">Spreadsheet to add data into</param>
    private static void WriteExcelFile(DataSet ds, SpreadsheetDocument spreadsheet)
    {
        spreadsheet.AddWorkbookPart();
        spreadsheet.WorkbookPart!.Workbook = new Workbook();
        spreadsheet.WorkbookPart.Workbook.Append(new BookViews(new WorkbookView()));

        // If we don't add a "WorkbookStylesPart", OLEDB will refuse to connect to this .xlsx file
        var workbookStylesPart = spreadsheet.WorkbookPart.AddNewPart<WorkbookStylesPart>("rIdStyles");
        var stylesheet = new Stylesheet();
        workbookStylesPart.Stylesheet = stylesheet;

        // Loop through each of the DataTables in our DataSet, and create a new Excel Worksheet for each.
        uint worksheetNumber = 1;
        foreach (DataTable dt in ds.Tables)
        {
            var worksheetName = GetWorksheetName(dt.TableName, worksheetNumber);

            var newWorksheetPart = spreadsheet.WorkbookPart.AddNewPart<WorksheetPart>();
            newWorksheetPart.Worksheet = new Worksheet();

            // Create sheet data
            newWorksheetPart.Worksheet.AppendChild(new SheetData());

            // Save worksheet
            WriteDataTableToExcelWorksheet(dt, newWorksheetPart);
            newWorksheetPart.Worksheet.Save();

            // Create the worksheet to workbook relation
            if (worksheetNumber == 1)
                spreadsheet.WorkbookPart.Workbook.AppendChild(new Sheets());

            spreadsheet.WorkbookPart.Workbook.GetFirstChild<Sheets>()!.AppendChild(new Sheet()
            {
                Id = spreadsheet.WorkbookPart.GetIdOfPart(newWorksheetPart),
                SheetId = (uint)worksheetNumber,
                Name = worksheetName
            });

            worksheetNumber++;
        }
                   
        spreadsheet.WorkbookPart.Workbook.Save();
    }

    /// <summary>
    /// Writes the given data table into the given worksheet part
    /// </summary>
    /// <param name="dt">Data to add to worksheet</param>
    /// <param name="worksheetPart">Worksheet</param>
    private static void WriteDataTableToExcelWorksheet(DataTable dt, WorksheetPart worksheetPart)
    {
        var worksheet = worksheetPart.Worksheet;
        var sheetData = worksheet!.GetFirstChild<SheetData>();

        //  Create a Header Row in our Excel file, containing one header for each Column of data in our DataTable.
        //  We'll also create an array, showing which type each column of data is (Text or Numeric), so when we come to write the actual
        //  cells of data, we'll know if to write Text values or Numeric cell values.
        var numberOfColumns = dt.Columns.Count;
        var columnTypes = new Type[numberOfColumns];

        var excelColumnNames = new string[numberOfColumns];
        for (var n = 0; n < numberOfColumns; n++)
        {
            excelColumnNames[n] = EPPlusHelper.GetColumnName(n);
            columnTypes[n] = GetNullableType(dt.Columns[n].DataType);
        }

        //  Create the Header row in our Excel Worksheet
        uint rowIndex = 1;

        // Add a row at the top of spreadsheet
        var headerRow = new Row { RowIndex = rowIndex };
        sheetData!.Append(headerRow);

        for (var colInx = 0; colInx < numberOfColumns; colInx++)
        {
            var col = dt.Columns[colInx];
            AppendTextCell(excelColumnNames[colInx] + "1", col.ColumnName, headerRow);
        }

        foreach (DataRow dr in dt.Rows)
        {
            // Create a new row, and append a set of this row's data to it.
            ++rowIndex;

            // Add a row at the top of spreadsheet
            var newExcelRow = new Row { RowIndex = rowIndex };           
            sheetData.Append(newExcelRow);

            for (var colInx = 0; colInx < numberOfColumns; colInx++)
            {
                var cellReference = excelColumnNames[colInx] + rowIndex.ToString();
                var cellValue = dr.ItemArray[colInx];

                // Create cell with data
                if (cellValue is null || cellValue == DBNull.Value)
                {
                    if (ShouldWriteEmptyTextCell(columnTypes[colInx]))
                        AppendTextCell(cellReference, string.Empty, newExcelRow);

                    continue;
                }

                AppendCell(cellReference, cellValue, newExcelRow);
            }
        }
    }

    /// <summary>
    /// Adds a cell using the most appropriate Excel cell type for the given value.
    /// </summary>
    /// <param name="cellReference">Cell to add e.g. [A1]</param>
    /// <param name="value">Value to set in cell</param>
    /// <param name="excelRow">Row to add cell onto</param>
    private static void AppendCell(string cellReference, object value, Row excelRow)
    {
        var valueType = GetNullableType(value.GetType());

        if (IsNumericType(valueType))
        {
            AppendNumericCell(cellReference, Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty, excelRow);
            return;
        }

        if (valueType == typeof(bool))
        {
            AppendBooleanCell(cellReference, (bool)value, excelRow);
            return;
        }

        if (TryGetExcelSerialValue(value, out var serialValue))
        {
            AppendNumericCell(cellReference, serialValue.ToString(CultureInfo.InvariantCulture), excelRow);
            return;
        }

        if (valueType.IsEnum)
        {
            AppendNumericCell(cellReference, Convert.ToInt64(value, CultureInfo.InvariantCulture).ToString(CultureInfo.InvariantCulture), excelRow);
            return;
        }

        AppendTextCell(cellReference, Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty, excelRow);
    }

    /// <summary>
    /// Checks to see if a column should write an explicit empty text cell for null values.
    /// </summary>
    /// <param name="type">Column type</param>
    /// <returns>True if an empty text cell should be written</returns>
    private static bool ShouldWriteEmptyTextCell(Type type)
    {
        type = GetNullableType(type);

        return !IsNumericType(type)
            && type != typeof(bool)
            && !type.IsEnum
            && type != typeof(DateTime)
            && type != typeof(DateTimeOffset)
            && type != typeof(DateOnly)
            && type != typeof(TimeOnly)
            && type != typeof(TimeSpan);
    }

    /// <summary>
    /// Converts supported date and time values into Excel serial numbers.
    /// </summary>
    /// <param name="value">Value to convert</param>
    /// <param name="serialValue">Excel serial number</param>
    /// <returns>True if the value was converted</returns>
    private static bool TryGetExcelSerialValue(object value, out double serialValue)
    {
        switch (value)
        {
            case DateTime dateTime:
                serialValue = dateTime.ToOADate();
                return true;
            case DateTimeOffset dateTimeOffset:
                serialValue = dateTimeOffset.DateTime.ToOADate();
                return true;
            case DateOnly dateOnly:
                serialValue = dateOnly.ToDateTime(TimeOnly.MinValue).ToOADate();
                return true;
            case TimeOnly timeOnly:
                serialValue = timeOnly.ToTimeSpan().TotalDays;
                return true;
            case TimeSpan timeSpan:
                serialValue = timeSpan.TotalDays;
                return true;
            default:
                serialValue = default;
                return false;
        }
    }

    /// <summary>
    /// Add a cell to the row with the given string value
    /// </summary>
    /// <param name="cellReference">Cell to add e.g. [A1]</param>
    /// <param name="cellStringValue">Value to set in cell</param>
    /// <param name="excelRow">Row to add cell onto</param>
    private static void AppendTextCell(string cellReference, string cellStringValue, Row excelRow)
    {
        // Add a new Excel Cell to our Row 
        var cell = new Cell()
        {
            CellReference = cellReference,
            DataType = CellValues.String
        };

        var cellValue = new CellValue { Text = cellStringValue };

        cell.Append(cellValue);
        excelRow.Append(cell);
    }

    /// <summary>
    /// Add a cell to the row with the given numeric value
    /// </summary>
    /// <param name="cellReference">Cell to add e.g. [A1]</param>
    /// <param name="cellStringValue">Value to set in cell</param>
    /// <param name="excelRow">Row to add cell onto</param>
    private static void AppendNumericCell(string cellReference, string cellStringValue, Row excelRow)
    {
        //  Add a new Excel Cell to our Row 
        var cell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = cellReference };
        var cellValue = new CellValue { Text = cellStringValue };

        cell.Append(cellValue);
        excelRow.Append(cell);
    }

    /// <summary>
    /// Add a cell to the row with the given boolean value.
    /// </summary>
    /// <param name="cellReference">Cell to add e.g. [A1]</param>
    /// <param name="value">Value to set in cell</param>
    /// <param name="excelRow">Row to add cell onto</param>
    private static void AppendBooleanCell(string cellReference, bool value, Row excelRow)
    {
        var cell = new Cell()
        {
            CellReference = cellReference,
            DataType = CellValues.Boolean
        };

        var cellValue = new CellValue { Text = value ? "1" : "0" };

        cell.Append(cellValue);
        excelRow.Append(cell);
    }
}
