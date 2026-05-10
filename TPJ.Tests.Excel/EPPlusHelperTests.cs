using OfficeOpenXml;
using TPJ.Excel;
using TPJ.Excel.Models;

namespace TPJ.Tests.Excel;

[TestFixture]
public class EPPlusHelperTests
{
    [TestCase(0, "A")]
    [TestCase(25, "Z")]
    [TestCase(26, "AA")]
    [TestCase(51, "AZ")]
    [TestCase(52, "BA")]
    [TestCase(701, "ZZ")]
    [TestCase(702, "AAA")]
    public void GetColumnName_ReturnsExpectedName(int columnIndex, string expected)
    {
        var result = EPPlusHelper.GetColumnName(columnIndex);

        Assert.That(result, Is.EqualTo(expected));
    }

    [Test]
    public void GetColumnName_ThrowsForNegativeIndex()
    {
        Assert.That(() => EPPlusHelper.GetColumnName(-1), Throws.InstanceOf<ArgumentOutOfRangeException>());
    }

    [Test]
    public void Worksheet_WithBuilder_CreatesExpectedWorksheetModel()
    {
        var staff = new[]
        {
            new Staff { Id = 1, Name = "Alice", StartDate = new DateTime(2024, 1, 15), Salary = 1234.56m }
        };

        var worksheet = EPPlusHelper.Worksheet(
            staff,
            "Staff",
            columns => columns
                .Add("Staff #", x => x.Id)
                .Add("Name", x => x.Name)
                .Add("Start Date", x => x.StartDate, dateFormat: "dd/mm/yyyy")
                .Add("Salary", x => x.Salary, numberFormat: "#,##0.00"));

        var headers = worksheet.Headers!.ToArray();
        var rows = worksheet.Rows.ToArray();
        var data = rows[0].Data.ToArray();

        Assert.Multiple(() =>
        {
            Assert.That(worksheet.Name, Is.EqualTo("Staff"));
            Assert.That(headers, Is.EqualTo(new[] { "Staff #", "Name", "Start Date", "Salary" }));
            Assert.That(rows, Has.Length.EqualTo(1));
            Assert.That(data, Has.Length.EqualTo(4));
            Assert.That(data[0].Value, Is.EqualTo(1));
            Assert.That(data[1].Value, Is.EqualTo("Alice"));
            Assert.That(data[2].Value, Is.EqualTo(new DateTime(2024, 1, 15)));
            Assert.That(data[2].DateFormat, Is.EqualTo("dd/mm/yyyy"));
            Assert.That(data[3].Value, Is.EqualTo(1234.56m));
            Assert.That(data[3].NumberFormat, Is.EqualTo("#,##0.00"));
        });
    }

    [Test]
    public void Create_WithTypedColumns_WritesHeadersValuesAndFormats()
    {
        var staff = new[]
        {
            new Staff { Id = 1, Name = "Alice", StartDate = new DateTime(2024, 1, 15), Salary = 1234.56m }
        };

        var bytes = EPPlusHelper.Create(
            staff,
            "Staff",
            columns => columns
                .Add("Staff #", x => x.Id)
                .Add("Name", x => x.Name)
                .Add("Start Date", x => x.StartDate, dateFormat: "dd/mm/yyyy")
                .Add("Salary", x => x.Salary, numberFormat: "#,##0.00"));

        using var stream = new MemoryStream(bytes);
        using var package = new ExcelPackage(stream);
        var worksheet = package.Workbook.Worksheets["Staff"];

        Assert.That(worksheet, Is.Not.Null);
        Assert.Multiple(() =>
        {
            Assert.That(worksheet!.Cells["A1"].Value, Is.EqualTo("Staff #"));
            Assert.That(worksheet.Cells["B1"].Value, Is.EqualTo("Name"));
            Assert.That(worksheet.Cells["C1"].Value, Is.EqualTo("Start Date"));
            Assert.That(worksheet.Cells["D1"].Value, Is.EqualTo("Salary"));
            Assert.That(worksheet.Cells["A2"].Value, Is.EqualTo(1));
            Assert.That(worksheet.Cells["B2"].Value, Is.EqualTo("Alice"));
            Assert.That(DateTime.FromOADate(Convert.ToDouble(worksheet.Cells["C2"].Value)), Is.EqualTo(new DateTime(2024, 1, 15)));
            Assert.That(worksheet.Cells["C2"].Style.Numberformat.Format, Is.EqualTo("dd/mm/yyyy"));
            Assert.That(worksheet.Cells["D2"].Value, Is.EqualTo(1234.56d));
            Assert.That(worksheet.Cells["D2"].Style.Numberformat.Format, Is.EqualTo("#,##0.00"));
        });
    }

    [Test]
    public void Worksheet_ThrowsForInvalidWorksheetName()
    {
        var staff = new[]
        {
            new Staff { Id = 1, Name = "Alice", StartDate = new DateTime(2024, 1, 15), Salary = 1234.56m }
        };

        Assert.That(
            () => EPPlusHelper.Worksheet(staff, "Invalid/Name", EPPlusHelper.Column<Staff>("Staff #", x => x.Id)),
            Throws.ArgumentException);
    }

    [Test]
    public void Create_ThrowsForDuplicateWorksheetNames()
    {
        var worksheets = new[]
        {
            new EPPlusWorksheet { Name = "Staff", Rows = Array.Empty<EPPlusRow>() },
            new EPPlusWorksheet { Name = "staff", Rows = Array.Empty<EPPlusRow>() }
        };

        Assert.That(() => EPPlusHelper.Create(worksheets), Throws.ArgumentException);
    }

    private sealed class Staff
    {
        public int Id { get; init; }

        public required string Name { get; init; }

        public DateTime StartDate { get; init; }

        public decimal Salary { get; init; }
    }
}
