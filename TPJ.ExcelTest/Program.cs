using OfficeOpenXml;
using TPJ.Excel;

namespace TPJ.ExcelTest
{
    internal class Program
    {
        static void Main()
        {
            var baseSaveLocation = "C:\\Test";
            Directory.CreateDirectory(baseSaveLocation);

            var staff = new List<Staff>()
            {
                new ()
                {
                    Id = 1,
                    Name = "Jeff Bob",
                    StartDate = DateTime.Now.AddDays(-5),
                },
                new ()
                {
                    Id = 2,
                    Name = "Thomas James",
                    StartDate = DateTime.Now.AddYears(-2),
                },
                new ()
                {
                    Id = 3,
                    Name = "Sophie Hall",
                    StartDate = DateTime.Now,
                }
            };

            // As simple as you can get, pass in the data object you wish to create an excel file for
            ExcelDocument.Create(staff, $@"{baseSaveLocation}\SimpleExcelDocument.xlsx");

            // You can also write any IEnumerable<T> directly to disk and choose the worksheet name.
            ExcelDocument.Create(staff.Where(x => x.Id > 0), $@"{baseSaveLocation}\SimpleExcelDocumentNamedSheet.xlsx", "Staff");

            // Simple scalar values are exported as a single Value column.
            ExcelDocument.Create(staff.Select(x => x.Name), $@"{baseSaveLocation}\SimpleExcelDocumentScalarValues.xlsx", "Names");

            // The typed EPPlus helper keeps custom headers and formatting without the nested row/cell model.
            EPPlusHelper.Create(
                staff,
                "Staff",
                $@"{baseSaveLocation}\SimpleEPPlus.xlsx",
                columns => columns
                    .Add("Staff #", x => x.Id)
                    .Add("Name", x => x.Name)
                    .Add("Start Date", x => x.StartDate, dateFormat: "dd/mm/yyyy"));

            // If you need full control over the excel document you can use full EPPlus.
            // Calling the extension method 'AddWorksheet' on the workbook returns an object containing the worksheet that
            // has the ability to track the currently 'selected' cell to make moving though the worksheet simple and clean
            // Note - the below will produce the same as the EPPlusHelper.Create above
            using var p = new ExcelPackage();
            var ws = p.Workbook.AddWorksheet("Staff");

            ws.WriteHeader("Staff #", "Name", "Start Date");

            foreach (var item in staff)
            {
                ws.WriteNext(item.Id)
                  .WriteNext(item.Name)
                  .Write(item.StartDate, dateFormat: "dd/mm/yyyy");

                ws.NextRow();
            }

            p.SaveAs(new FileInfo($@"{baseSaveLocation}\ComplexEPPlus.xlsx"));
        }
    }

    class Staff
    {
        public int Id { get; set; }
        public required string Name { get; set; }
        public DateTime StartDate { get; set; }
    }
}