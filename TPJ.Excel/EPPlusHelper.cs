using OfficeOpenXml;
using TPJ.Excel.Models;

namespace TPJ.Excel
{
    public static class EPPlusHelper
    {
        public static readonly string[] Columns = new string[702]
       {
         "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z",
        "AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH", "AI", "AJ", "AK", "AL", "AM", "AN", "AO", "AP", "AQ", "AR", "AS", "AT", "AU", "AV", "AW", "AX", "AY", "AZ",
        "BA", "BB", "BC", "BD", "BE", "BF", "BG", "BH", "BI", "BJ", "BK", "BL", "BM", "BN", "BO", "BP", "BQ", "BR", "BS", "BT", "BU", "BV", "BW", "BX", "BY", "BZ",
        "CA", "CB", "CC", "CD", "CE", "CF", "CG", "CH", "CI", "CJ", "CK", "CL", "CM", "CN", "CO", "CP", "CQ", "CR", "CS", "CT", "CU", "CV", "CW", "CX", "CY", "CZ",
        "DA", "DB", "DC", "DD", "DE", "DF", "DG", "DH", "DI", "DJ", "DK", "DL", "DM", "DN", "DO", "DP", "DQ", "DR", "DS", "DT", "DU", "DV", "DW", "DX", "DY", "DZ",
        "EA", "EB", "EC", "ED", "EE", "EF", "EG", "EH", "EI", "EJ", "EK", "EL", "EM", "EN", "EO", "EP", "EQ", "ER", "ES", "ET", "EU", "EV", "EW", "EX", "EY", "EZ",
        "FA", "FB", "FC", "FD", "FE", "FF", "FG", "FH", "FI", "FJ", "FK", "FL", "FM", "FN", "FO", "FP", "FQ", "FR", "FS", "FT", "FU", "FV", "FW", "FX", "FY", "FZ",
        "GA", "GB", "GC", "GD", "GE", "GF", "GG", "GH", "GI", "GJ", "GK", "GL", "GM", "GN", "GO", "GP", "GQ", "GR", "GS", "GT", "GU", "GV", "GW", "GX", "GY", "GZ",
        "HA", "HB", "HC", "HD", "HE", "HF", "HG", "HH", "HI", "HJ", "HK", "HL", "HM", "HN", "HO", "HP", "HQ", "HR", "HS", "HT", "HU", "HV", "HW", "HX", "HY", "HZ",
        "IA", "IB", "IC", "ID", "IE", "IF", "IG", "IH", "II", "IJ", "IK", "IL", "IM", "IN", "IO", "IP", "IQ", "IR", "IS", "IT", "IU", "IV", "IW", "IX", "IY", "IZ",
        "JA", "JB", "JC", "JD", "JE", "JF", "JG", "JH", "JI", "JJ", "JK", "JL", "JM", "JN", "JO", "JP", "JQ", "JR", "JS", "JT", "JU", "JV", "JW", "JX", "JY", "JZ",
        "KA", "KB", "KC", "KD", "KE", "KF", "KG", "KH", "KI", "KJ", "KK", "KL", "KM", "KN", "KO", "KP", "KQ", "KR", "KS", "KT", "KU", "KV", "KW", "KX", "KY", "KZ",
        "LA", "LB", "LC", "LD", "LE", "LF", "LG", "LH", "LI", "LJ", "LK", "LL", "LM", "LN", "LO", "LP", "LQ", "LR", "LS", "LT", "LU", "LV", "LW", "LX", "LY", "LZ",
        "MA", "MB", "MC", "MD", "ME", "MF", "MG", "MH", "MI", "MJ", "MK", "ML", "MM", "MN", "MO", "MP", "MQ", "MR", "MS", "MT", "MU", "MV", "MW", "MX", "MY", "MZ",
        "NA", "NB", "NC", "ND", "NE", "NF", "NG", "NH", "NI", "NJ", "NK", "NL", "NM", "NN", "NO", "NP", "NQ", "NR", "NS", "NT", "NU", "NV", "NW", "NX", "NY", "NZ",
        "PA", "PB", "PC", "PD", "PE", "PF", "PG", "PH", "PI", "PJ", "PK", "PL", "PM", "PN", "PO", "PP", "PQ", "PR", "PS", "PT", "PU", "PV", "PW", "PX", "PY", "PZ",
        "PA", "PB", "PC", "PD", "PE", "PF", "PG", "PH", "PI", "PJ", "PK", "PL", "PM", "PN", "PO", "PP", "PQ", "PR", "PS", "PT", "PU", "PV", "PW", "PX", "PY", "PZ",
        "QA", "QB", "QC", "QD", "QE", "QF", "QG", "QH", "QI", "QJ", "QK", "QL", "QM", "QN", "QO", "QP", "QQ", "QR", "QS", "QT", "QU", "QV", "QW", "QX", "QY", "QZ",
        "RA", "RB", "RC", "RD", "RE", "RF", "RG", "RH", "RI", "RJ", "RK", "RL", "RM", "RN", "RO", "RP", "RQ", "RR", "RS", "RT", "RU", "RV", "RW", "RX", "RY", "RZ",
        "SA", "SB", "SC", "SD", "SE", "SF", "SG", "SH", "SI", "SJ", "SK", "SL", "SM", "SN", "SO", "SP", "SQ", "SR", "SS", "ST", "SU", "SV", "SW", "SX", "SY", "SZ",
        "TA", "TB", "TC", "TD", "TE", "TF", "TG", "TH", "TI", "TJ", "TK", "TL", "TM", "TN", "TO", "TP", "TQ", "TR", "TS", "TT", "TU", "TV", "TW", "TX", "TY", "TZ",
        "UA", "UB", "UC", "UD", "UE", "UF", "UG", "UH", "UI", "UJ", "UK", "UL", "UM", "UN", "UO", "UP", "UQ", "UR", "US", "UT", "UU", "UV", "UW", "UX", "UY", "UZ",
        "VA", "VB", "VC", "VD", "VE", "VF", "VG", "VH", "VI", "VJ", "VK", "VL", "VM", "VN", "VO", "VP", "VQ", "VR", "VS", "VT", "VU", "VV", "VW", "VX", "VY", "VZ",
        "WA", "WB", "WC", "WD", "WE", "WF", "WG", "WH", "WI", "WJ", "WK", "WL", "WM", "WN", "WO", "WP", "WQ", "WR", "WS", "WT", "WU", "WV", "WW", "WX", "WY", "WZ",
        "XA", "XB", "XC", "XD", "XE", "XF", "XG", "XH", "XI", "XJ", "XK", "XL", "XM", "XN", "XO", "XP", "XQ", "XR", "XS", "XT", "XU", "XV", "XW", "XX", "XY", "XZ",
        "YA", "YB", "YC", "YD", "YE", "YF", "YG", "YH", "YI", "YJ", "YK", "YL", "YM", "YN", "YO", "YP", "YQ", "YR", "YS", "YT", "YU", "YV", "YW", "YX", "YY", "YZ",
        "ZA", "ZB", "ZC", "ZD", "ZE", "ZF", "ZG", "ZH", "ZI", "ZJ", "ZK", "ZL", "ZM", "ZN", "ZO", "ZP", "ZQ", "ZR", "ZS", "ZT", "ZU", "ZV", "ZW", "ZX", "ZY", "ZZ"
       };

        /// <summary>
        /// Gets a cell based on the passed in column and row value e.g. 0, 1 will return cell A and and 5, 3 will return cell BF
        /// </summary>
        /// <param name="ws">Worksheet to get cell from</param>
        /// <param name="column">Column to get (0 base)</param>
        /// <param name="row">Row to get (1 base)</param>
        /// <returns>Cell</returns>
        public static ExcelRange Cell(this ExcelWorksheet ws, int column, int row) => ws.Cells[$"{Columns[column]}{row}"];

        public static ExcelWorksheetExtension AddWorksheet(this ExcelWorkbook workbook, string name)
        {
            var ws = workbook.Worksheets.Add(name);

            return new ExcelWorksheetExtension()
            {
                Worksheet = ws
            };
        }

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
