using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Globalization;
using TPJ.Excel;

namespace TPJ.Tests.Excel;

[TestFixture]
public class ExcelDocumentTests
{
    [Test]
    public void Create_WithScalarValues_UsesValueHeaderAndWorksheetName()
    {
        var bytes = ExcelDocument.Create(new[] { "Alice", "Bob" });

        using var stream = new MemoryStream(bytes);
        using var document = SpreadsheetDocument.Open(stream, false);
        var workbookPart = document.WorkbookPart!;
        var sheet = workbookPart.Workbook.Sheets!.Elements<Sheet>().Single();

        Assert.Multiple(() =>
        {
            Assert.That(sheet.Name!.Value, Is.EqualTo("Sheet1"));
            Assert.That(GetCellText(workbookPart, sheet, "A1"), Is.EqualTo("Value"));
            Assert.That(GetCellText(workbookPart, sheet, "A2"), Is.EqualTo("Alice"));
            Assert.That(GetCellText(workbookPart, sheet, "A3"), Is.EqualTo("Bob"));
        });
    }

    [Test]
    public void Create_ToFile_WithNamedWorksheet_WritesPropertyHeadersAndValues()
    {
        var staff = new[]
        {
            new Staff { Id = 1, Name = "Alice", StartDate = new DateTime(2024, 1, 15) },
            new Staff { Id = 2, Name = "Bob", StartDate = new DateTime(2024, 6, 1) }
        };

        var tempFile = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid():N}.xlsx");

        try
        {
            ExcelDocument.Create(staff, tempFile, "Staff");

            using var document = SpreadsheetDocument.Open(tempFile, false);
            var workbookPart = document.WorkbookPart!;
            var sheet = workbookPart.Workbook.Sheets!.Elements<Sheet>().Single();

            Assert.Multiple(() =>
            {
                Assert.That(sheet.Name!.Value, Is.EqualTo("Staff"));
                Assert.That(GetCellText(workbookPart, sheet, "A1"), Is.EqualTo("Id"));
                Assert.That(GetCellText(workbookPart, sheet, "B1"), Is.EqualTo("Name"));
                Assert.That(GetCellText(workbookPart, sheet, "C1"), Is.EqualTo("StartDate"));
                Assert.That(GetCellText(workbookPart, sheet, "A2"), Is.EqualTo("1"));
                Assert.That(GetCellText(workbookPart, sheet, "B2"), Is.EqualTo("Alice"));
                Assert.That(GetCellText(workbookPart, sheet, "C2"), Is.Not.Empty);
                Assert.That(GetCellText(workbookPart, sheet, "A3"), Is.EqualTo("2"));
                Assert.That(GetCellText(workbookPart, sheet, "B3"), Is.EqualTo("Bob"));
            });
        }
        finally
        {
            if (File.Exists(tempFile))
                File.Delete(tempFile);
        }
    }

    [Test]
    public void Create_WithNullableReferenceProperty_WritesBlankCellForNull()
    {
        var rows = new[]
        {
            new StaffWithManager { Name = "Alice", Manager = null }
        };

        var bytes = ExcelDocument.Create(rows);

        using var stream = new MemoryStream(bytes);
        using var document = SpreadsheetDocument.Open(stream, false);
        var workbookPart = document.WorkbookPart!;
        var sheet = workbookPart.Workbook.Sheets!.Elements<Sheet>().Single();

        Assert.Multiple(() =>
        {
            Assert.That(GetCellText(workbookPart, sheet, "A1"), Is.EqualTo("Name"));
            Assert.That(GetCellText(workbookPart, sheet, "B1"), Is.EqualTo("Manager"));
            Assert.That(GetCellText(workbookPart, sheet, "A2"), Is.EqualTo("Alice"));
            Assert.That(GetCellText(workbookPart, sheet, "B2"), Is.EqualTo(string.Empty));
        });
    }

    [Test]
    public void Create_WithDateAndBooleanValues_WritesNativeCellTypes()
    {
        var rows = new[]
        {
            new StatusRow { CreatedOn = new DateTime(2024, 1, 15, 13, 45, 0), IsActive = true }
        };

        var bytes = ExcelDocument.Create(rows);

        using var stream = new MemoryStream(bytes);
        using var document = SpreadsheetDocument.Open(stream, false);
        var workbookPart = document.WorkbookPart!;
        var sheet = workbookPart.Workbook.Sheets!.Elements<Sheet>().Single();
        var createdOnCell = GetCell(workbookPart, sheet, "A2");
        var isActiveCell = GetCell(workbookPart, sheet, "B2");

        Assert.Multiple(() =>
        {
            Assert.That(createdOnCell.DataType, Is.Null);
            Assert.That(double.Parse(createdOnCell.CellValue!.Text, CultureInfo.InvariantCulture), Is.EqualTo(rows[0].CreatedOn.ToOADate()).Within(0.0000001));
            Assert.That(isActiveCell.DataType?.Value, Is.EqualTo(CellValues.Boolean));
            Assert.That(isActiveCell.CellValue?.Text, Is.EqualTo("1"));
        });
    }

    private static string GetCellText(WorkbookPart workbookPart, Sheet sheet, string cellReference) =>
        GetCell(workbookPart, sheet, cellReference).CellValue?.Text ?? string.Empty;

    private static Cell GetCell(WorkbookPart workbookPart, Sheet sheet, string cellReference)
    {
        var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id!);
        return worksheetPart.Worksheet.Descendants<Cell>()
            .Single(x => string.Equals(x.CellReference?.Value, cellReference, StringComparison.OrdinalIgnoreCase));
    }

    private sealed class Staff
    {
        public int Id { get; init; }

        public required string Name { get; init; }

        public DateTime StartDate { get; init; }
    }

    private sealed class StaffWithManager
    {
        public required string Name { get; init; }

        public string? Manager { get; init; }
    }

    private sealed class StatusRow
    {
        public DateTime CreatedOn { get; init; }

        public bool IsActive { get; init; }
    }
}
