using OfficeOpenXml;
using TPJ.Excel;

namespace TPJ.Tests.Excel;

[TestFixture]
public class ExcelWorksheetExtensionTests
{
    [Test]
    public void WriteHelpers_WriteExpectedCellsAndMoveToNextRow()
    {
        using var package = new ExcelPackage();
        var worksheet = package.Workbook.AddWorksheet("Staff");

        worksheet.WriteHeader("Staff #", "Name", "Start Date");
        worksheet.WriteNext(1)
                 .WriteNext("Alice")
                 .Write(new DateTime(2024, 1, 15), dateFormat: "dd/mm/yyyy");
        worksheet.NextRow();
        worksheet.WriteRow(2, "Bob", new DateTime(2024, 6, 1));

        Assert.Multiple(() =>
        {
            Assert.That(worksheet.Row, Is.EqualTo(4));
            Assert.That(worksheet.Column, Is.EqualTo(0));
            Assert.That(worksheet.Worksheet.Cells["A1"].Value, Is.EqualTo("Staff #"));
            Assert.That(worksheet.Worksheet.Cells["A1"].Style.Font.Bold, Is.True);
            Assert.That(worksheet.Worksheet.Cells["B1"].Value, Is.EqualTo("Name"));
            Assert.That(worksheet.Worksheet.Cells["C1"].Value, Is.EqualTo("Start Date"));
            Assert.That(worksheet.Worksheet.Cells["A2"].Value, Is.EqualTo(1));
            Assert.That(worksheet.Worksheet.Cells["B2"].Value, Is.EqualTo("Alice"));
            Assert.That(worksheet.Worksheet.Cells["C2"].Value, Is.EqualTo(new DateTime(2024, 1, 15)));
            Assert.That(worksheet.Worksheet.Cells["C2"].Style.Numberformat.Format, Is.EqualTo("dd/mm/yyyy"));
            Assert.That(worksheet.Worksheet.Cells["A3"].Value, Is.EqualTo(2));
            Assert.That(worksheet.Worksheet.Cells["B3"].Value, Is.EqualTo("Bob"));
            Assert.That(worksheet.Worksheet.Cells["C3"].Value, Is.EqualTo(new DateTime(2024, 6, 1)));
        });
    }

    [Test]
    public void Write_WithNumberFormat_AppliesFormatWithoutMoving()
    {
        using var package = new ExcelPackage();
        var worksheet = package.Workbook.AddWorksheet("Numbers");

        worksheet.Write(12.34m, numberFormat: "#,##0.00", bold: true);

        Assert.Multiple(() =>
        {
            Assert.That(worksheet.Row, Is.EqualTo(1));
            Assert.That(worksheet.Column, Is.EqualTo(0));
            Assert.That(worksheet.Worksheet.Cells["A1"].Value, Is.EqualTo(12.34d));
            Assert.That(worksheet.Worksheet.Cells["A1"].Style.Numberformat.Format, Is.EqualTo("#,##0.00"));
            Assert.That(worksheet.Worksheet.Cells["A1"].Style.Font.Bold, Is.True);
        });
    }
}
