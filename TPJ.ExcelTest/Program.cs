using OfficeOpenXml;
using TPJ.Excel;
using TPJ.Excel.Models;

namespace TPJ.ExcelTest
{
    internal class Program
    {
        static void Main(string[] args)
        {
            var baseSaveLocation = "C:\\Test";
            Directory.CreateDirectory(baseSaveLocation);

            var staff = new List<Staff>()
            {
                new Staff()
                {
                    Id = 1,
                    Name = "Jeff Bob",
                    StartDate = DateTime.Now.AddDays(-5),
                },
                new Staff()
                {
                    Id = 2,
                    Name = "Thomas James",
                    StartDate = DateTime.Now.AddYears(-2),
                },
                new Staff()
                {
                    Id = 3,
                    Name = "Sophie Hall",
                    StartDate = DateTime.Now,
                }
            };

            // As simple as you can get, pass in the data object you wish to create an excel file for
            // Note - the object must be a class e.g. you can't pass in a list of string but you can
            // pass in a list of a class that has a single property of a string.
            // This is due to the fact the property name is used as the column heading
            ExcelDocument.Create(staff, $@"{baseSaveLocation}\SimpleExcelDocument.xlsx");

            // The EPPlus helper gives you more control over how the simple table will look
            // You set the column heading name and can pass in how you wish the data to be formatted
            EPPlusHelper.Create(new EPPlusWorksheet("Staff")
            {
                Headers = new List<string>()
                {
                    "Staff #",
                    "Name",
                    "Start Date"
                },
                Rows = staff.Select(x => new EPPlusRow(new List<EPPlusData>()
                {
                    new EPPlusData(x.Id),
                    new EPPlusData(x.Name),
                    new EPPlusData(x.StartDate)
                    {
                        DateFormat = "dd/mm/yyyy"
                    },
                }))
            }, $@"{baseSaveLocation}\SimpleEPPlus.xlsx");

            // If you need full control over the excel document you can use full EPPlus.
            // Calling the extension method 'AddWorksheet' on the workbook returns an object containing the worksheet that
            // has the ability to track the currently 'selected' cell to make moving though the worksheet simple and clean
            // Note - the below will produce the same as the EPPlusHelper.Create above
            using var p = new ExcelPackage();
            var ws = p.Workbook.AddWorksheet("Staff");

            ws.Cell().Value = "Staff #";
            ws.Cell().Style.Font.Bold = true;
            ws.NextColumn();

            ws.Cell().Value = "Name";
            ws.Cell().Style.Font.Bold = true;
            ws.NextColumn();

            ws.Cell().Value = "Start Date";
            ws.Cell().Style.Font.Bold = true;

            ws.NextRow();

            foreach (var item in staff)
            {
                ws.Cell().Value = item.Id;
                ws.NextColumn();

                ws.Cell().Value = item.Name;
                ws.NextColumn();

                ws.Cell().Value = item.StartDate;
                ws.Cell().Style.Numberformat.Format = "dd/mm/yyyy";

                ws.NextRow();
            }

            p.SaveAs(new FileInfo($@"{baseSaveLocation}\ComplexEPPlus.xlsx"));
        }
    }

    class Staff
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public DateTime StartDate { get; set; }
    }
}