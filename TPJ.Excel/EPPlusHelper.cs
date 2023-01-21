using OfficeOpenXml;
using TPJ.Excel.Models;

namespace TPJ.Excel
{
    public static class EPPlusHelper
    {
        /// <summary>
        /// Creates an excel file in memory containing the given worksheet
        /// </summary>
        /// <param name="worksheet">Worksheet to create</param>
        /// <returns>Bytes of the created excel worksheet</returns>
        public static byte[] Create(EPPlusWorksheet worksheet) => Create(new EPPlusWorksheet[1] { worksheet });

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
        public static void Create(EPPlusWorksheet worksheet, string filePath) => Create(new EPPlusWorksheet[1] { worksheet }, filePath);

        /// <summary>
        /// Creates an excel file to disk containing the given worksheets
        /// </summary>
        /// <param name="worksheets">Worksheets to create</param>
        /// <param name="filePath">Save excel file path e.g. C:\example.xlsx</param>
        public static void Create(IEnumerable<EPPlusWorksheet> worksheets, string filePath)
        {
            using var ms = new MemoryStream();
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
            var excelPackage = new ExcelPackage();

            foreach (var worksheet in worksheets)
            {
                var ws = excelPackage.Workbook.Worksheets.Add(worksheet.Name);

                if (worksheet.Headers != null && worksheet.Headers.Any())
                    ws.AddHeaders(worksheet.Headers);

                ws.AddData(worksheet.Rows, worksheet.Headers is null || worksheet.Headers.Any() ? 2 : 1);
            }

            return excelPackage;
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
}
