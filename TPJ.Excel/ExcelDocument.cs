using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Data;

namespace TPJ.Excel;

public static class ExcelDocument
{
    /// <summary>
    /// Create an excel document using the 
    /// given list of data and the excel document path
    /// </summary>
    /// <typeparam name="T">Object Type</typeparam>
    /// <param name="list">Data</param>
    /// <param name="filePath">Save excel file path e.g. C:\example.xlsx</param>
    public static void Create<T>(List<T> list, string filePath) => Create(ListToDataTable(list), filePath);

    /// <summary>
    /// Create an excel document using the 
    /// given data table and the excel document path
    /// </summary>
    /// <param name="dt">Data Table</param>
    /// <param name="filePath">Save excel file path e.g. C:\example.xlsx</param>
    public static void Create(DataTable dt, string filePath)
    {
        var ds = new DataSet();
        ds.Tables.Add(dt);
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
        using var document = SpreadsheetDocument.Create(filePath, SpreadsheetDocumentType.Workbook);
        WriteExcelFile(ds, document);
    }

    /// <summary>
    /// Create an excel document with the given data enumerable
    /// </summary>
    /// <typeparam name="T">Object Type</typeparam>
    /// <param name="list">Data</param>
    /// <returns>Excel file in bytes</returns>
    public static byte[] Create<T>(IEnumerable<T> list) => Create(ListToDataTable(list));

    /// <summary>
    /// Create an excel document using a data table
    /// </summary>
    /// <param name="dt">Data Table</param>
    /// <returns>Excel file in bytes</returns>
    public static byte[] Create(DataTable dt)
    {
        var ds = new DataSet();
        ds.Tables.Add(dt);
        return Create(ds);
    }

    /// <summary>
    /// Create an excel document using the 
    /// given data set and the memory stream
    /// </summary>
    /// <param name="ds">Data Set</param>
    /// <param name="stream">Memory Stream</param>
    public static byte[] Create(DataSet ds)
    {
        using var ms = new MemoryStream();
        using var document = SpreadsheetDocument.Create(ms, SpreadsheetDocumentType.Workbook);
        WriteExcelFile(ds, document);

        return ms.ToArray();
    }

    /// <summary>
    /// Convert a list of data into a data table
    /// </summary>
    /// <typeparam name="T">Object Type</typeparam>
    /// <param name="list">Data List</param>
    /// <returns>Data table with all data within data list</returns>
    private static DataTable ListToDataTable<T>(IEnumerable<T> list)
    {
        var dt = new DataTable();

        foreach (var info in typeof(T).GetProperties())            
            dt.Columns.Add(new DataColumn(info.Name, GetNullableType(info.PropertyType)));
        
        foreach (var t in list)
        {
            var row = dt.NewRow();
            foreach (var info in typeof(T).GetProperties())
            {
                if (!IsNullableType(info.PropertyType))
                    row[info.Name] = info.GetValue(t, null);
                else
                    row[info.Name] = (info.GetValue(t, null) ?? DBNull.Value);
            }
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
        
        return returnType;
    }

    /// <summary>
    /// Checks to see if the given type is nullable
    /// </summary>
    /// <param name="type"></param>
    /// <returns></returns>
    private static bool IsNullableType(Type type) => (type == typeof(string) ||
                type.IsArray ||
                (type.IsGenericType &&
                 type.GetGenericTypeDefinition().Equals(typeof(Nullable<>))));

    // Note - Code below is not mine but can't remember where it is from.

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
        spreadsheet.WorkbookPart.Workbook = new Workbook();
        spreadsheet.WorkbookPart.Workbook.Append(new BookViews(new WorkbookView()));

        // If we don't add a "WorkbookStylesPart", OLEDB will refuse to connect to this .xlsx file
        var workbookStylesPart = spreadsheet.WorkbookPart.AddNewPart<WorkbookStylesPart>("rIdStyles");
        var stylesheet = new Stylesheet();
        workbookStylesPart.Stylesheet = stylesheet;

        // Loop through each of the DataTables in our DataSet, and create a new Excel Worksheet for each.
        uint worksheetNumber = 1;
        foreach (DataTable dt in ds.Tables)
        {
            // For each worksheet you want to create
            var workSheetID = "rId" + worksheetNumber.ToString();
            var worksheetName = dt.TableName;

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

            spreadsheet.WorkbookPart.Workbook.GetFirstChild<Sheets>().AppendChild(new Sheet()
            {
                Id = spreadsheet.WorkbookPart.GetIdOfPart(newWorksheetPart),
                SheetId = (uint)worksheetNumber,
                Name = dt.TableName
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
        var sheetData = worksheet.GetFirstChild<SheetData>();

        //  Create a Header Row in our Excel file, containing one header for each Column of data in our DataTable.
        //  We'll also create an array, showing which type each column of data is (Text or Numeric), so when we come to write the actual
        //  cells of data, we'll know if to write Text values or Numeric cell values.
        var numberOfColumns = dt.Columns.Count;
        var IsNumericColumn = new bool[numberOfColumns];

        var excelColumnNames = new string[numberOfColumns];
        for (var n = 0; n < numberOfColumns; n++)
            excelColumnNames[n] = GetExcelColumnName(n);

        //  Create the Header row in our Excel Worksheet
        uint rowIndex = 1;

        // Add a row at the top of spreadsheet
        var headerRow = new Row { RowIndex = rowIndex };
        sheetData.Append(headerRow);

        for (var colInx = 0; colInx < numberOfColumns; colInx++)
        {
            var col = dt.Columns[colInx];
            AppendTextCell(excelColumnNames[colInx] + "1", col.ColumnName, headerRow);
            IsNumericColumn[colInx] = (col.DataType.FullName == "System.Decimal") 
                || (col.DataType.FullName == "System.Int32") 
                || (col.DataType.FullName == "System.Double") 
                || (col.DataType.FullName == "System.Single");               
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
                var cellValue = dr.ItemArray[colInx].ToString();

                // Create cell with data
                if (IsNumericColumn[colInx])
                {
                    // Now, step through each row of data in our DataTable...
                    // For numeric cells, make sure our input data IS a number, then write it out to the Excel file.
                    // If this numeric value is NULL, then don't write anything to the Excel file.
                    if (double.TryParse(cellValue, out var cellNumericValue))
                    {
                        cellValue = cellNumericValue.ToString();
                        AppendNumericCell(excelColumnNames[colInx] + rowIndex.ToString(), cellValue, newExcelRow);
                    }
                }
                else
                {
                    // For text cells, just write the input data straight out to the Excel file.
                    AppendTextCell(excelColumnNames[colInx] + rowIndex.ToString(), cellValue, newExcelRow);
                }
            }
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
    /// Convert a zero-based column index into an Excel column reference  (A, B, C.. Y, Y, AA, AB, AC... AY, AZ, B1, B2..) e.g.
    /// GetExcelColumnName(0) should return "A"
    /// GetExcelColumnName(1) should return "B"
    /// GetExcelColumnName(25) should return "Z"
    /// GetExcelColumnName(26) should return "AA"
    /// GetExcelColumnName(27) should return "AB"
    /// </summary>
    /// <param name="columnIndex">Coulmn Index</param>
    /// <returns>Excel column reference</returns>
    private static string GetExcelColumnName(int columnIndex)
    {
        if (columnIndex < 26)
            return ((char)('A' + columnIndex)).ToString();

        var firstChar = (char)('A' + (columnIndex / 26) - 1);
        var secondChar = (char)('A' + (columnIndex % 26));

        return string.Format("{0}{1}", firstChar, secondChar);
    }
}
