# TPJ.Excel

`TPJ.Excel` is a small helper library for generating Excel files in three different ways:

1. `ExcelDocument` for quick exports from a list, `DataTable`, or `DataSet`
2. `EPPlusHelper` when you want custom headers and simple formatting
3. `AddWorksheet` / `ExcelWorksheetExtension` when you want full control while still writing concise EPPlus code

If you want to see working examples, check `TPJ.ExcelTest`.

## When to use each option

### Use `ExcelDocument` when you want speed and simplicity
- Create a workbook directly from a `List<T>`, `DataTable`, or `DataSet`
- Column names come from property names or table columns
- Good for quick exports where default formatting is enough

### Use `EPPlusHelper` when you want clean control over output
- Set worksheet names
- Choose your own headers
- Format dates and numbers
- Good for most business export scenarios

### Use full EPPlus support when you need custom layout
- Write exactly where you want in the sheet
- Move cell by cell using `CurrentCell()`, `NextColumn()`, and `NextRow()`
- Good for reports, grouped sections, totals, and custom styling

## Sample model

All examples below use this simple model:

```csharp
public class Staff
{
    public int Id { get; set; }
    public string Name { get; set; } = string.Empty;
    public DateTime StartDate { get; set; }
    public decimal Salary { get; set; }
}

var staff = new List<Staff>
{
    new() { Id = 1, Name = "Jeff Bob", StartDate = new DateTime(2024, 1, 15), Salary = 32500m },
    new() { Id = 2, Name = "Thomas James", StartDate = new DateTime(2022, 6, 1), Salary = 41000m },
    new() { Id = 3, Name = "Sophie Hall", StartDate = new DateTime(2025, 2, 10), Salary = 29500m }
};
```

## 1. Quick export with `ExcelDocument`

`ExcelDocument` is the quickest way to generate an Excel file.

### Save a list to disk

```csharp
using TPJ.Excel;

ExcelDocument.Create(staff, @"C:\Test\Staff.xlsx");
```

You can also pass any `IEnumerable<T>` and choose the worksheet name:

```csharp
using TPJ.Excel;

ExcelDocument.Create(staff.Where(x => x.Salary >= 30000m), @"C:\Test\StaffFiltered.xlsx", "Staff");
```

### Return the file as bytes

```csharp
using TPJ.Excel;

byte[] excelBytes = ExcelDocument.Create(staff);
```

### Notes

- Each public property becomes a column
- Property names are used as the header row
- Simple scalar values such as `string`, `int`, `decimal`, and `DateTime` are exported as a single `Value` column
- You can also create files from `DataTable` and `DataSet`

Example with a `DataTable`:

```csharp
using System.Data;
using TPJ.Excel;

var table = new DataTable("Staff");
table.Columns.Add("Staff #", typeof(int));
table.Columns.Add("Name", typeof(string));
table.Columns.Add("Start Date", typeof(DateTime));

table.Rows.Add(1, "Jeff Bob", new DateTime(2024, 1, 15));
table.Rows.Add(2, "Thomas James", new DateTime(2022, 6, 1));

ExcelDocument.Create(table, @"C:\Test\StaffFromTable.xlsx");
```

## 2. Custom headers and formatting with `EPPlusHelper`

Use `EPPlusHelper` when you want a simple API but need more control over how the sheet looks.

### Main types

For most cases, use the typed builder overloads:

```csharp
using TPJ.Excel;

EPPlusHelper.Create(
    staff,
    "Staff",
    @"C:\Test\StaffFormatted.xlsx",
    columns => columns
        .Add("Staff #", x => x.Id)
        .Add("Name", x => x.Name)
        .Add("Start Date", x => x.StartDate, dateFormat: "dd/mm/yyyy")
        .Add("Salary", x => x.Salary, numberFormat: "#,##0.00"));
```

You can still use `EPPlusHelper.Column<T>(...)` if you prefer building the column list yourself.

If you want to build multiple typed worksheets before creating the workbook, use `EPPlusHelper.Worksheet(...)`:

```csharp
using TPJ.Excel;

var summarySheet = EPPlusHelper.Worksheet(
    staff.Select(x => new { Label = x.Name, Created = DateTime.Today }),
    "Summary",
    columns => columns
        .Add("Label", x => x.Label)
        .Add("Created", x => x.Created, dateFormat: "dd/mm/yyyy"));

var staffSheet = EPPlusHelper.Worksheet(
    staff,
    "Staff",
    columns => columns
        .Add("Staff #", x => x.Id)
        .Add("Name", x => x.Name)
        .Add("Start Date", x => x.StartDate, dateFormat: "dd/mm/yyyy")
        .Add("Salary", x => x.Salary, numberFormat: "#,##0.00"));

EPPlusHelper.Create(new[] { summarySheet, staffSheet }, @"C:\Test\StaffReport.xlsx");
```

If you want to build the worksheet model yourself, these are the core types:

- `EPPlusWorksheet`
  - `Name`: worksheet name
  - `Headers`: column headers
  - `Rows`: row data
- `EPPlusRow`
  - `Data`: the cells for that row, in column order
- `EPPlusData`
  - `Value`: cell value
  - `DateFormat`: Excel date format string
  - `NumberFormat`: Excel number format string

Formatting options are passed straight to Excel cell formatting. Common examples include `dd/mm/yyyy` and `#,##0.00`.

### Save to disk with the worksheet model

```csharp
using TPJ.Excel;
using TPJ.Excel.Models;

var worksheet = new EPPlusWorksheet
{
    Name = "Staff",
    Headers =
    [
        "Staff #",
        "Name",
        "Start Date",
        "Salary"
    ],
    Rows = staff.Select(x => new EPPlusRow
    {
        Data =
        [
            new EPPlusData { Value = x.Id },
            new EPPlusData { Value = x.Name },
            new EPPlusData
            {
                Value = x.StartDate,
                DateFormat = "dd/mm/yyyy"
            },
            new EPPlusData
            {
                Value = x.Salary,
                NumberFormat = "#,##0.00"
            }
        ]
    })
};

EPPlusHelper.Create(worksheet, @"C:\Test\StaffFormatted.xlsx");
```

### Return the file as bytes

```csharp
using TPJ.Excel;

byte[] excelBytes = EPPlusHelper.Create(worksheet);
```

### Multiple worksheets

```csharp
using TPJ.Excel;
using TPJ.Excel.Models;

var summarySheet = new EPPlusWorksheet
{
    Name = "Summary",
    Headers = ["Report", "Created"],
    Rows =
    [
        new EPPlusRow
        {
            Data =
            [
                new EPPlusData { Value = "Staff Export" },
                new EPPlusData { Value = DateTime.Today, DateFormat = "dd/mm/yyyy" }
            ]
        }
    ]
};

EPPlusHelper.Create(new[] { summarySheet, worksheet }, @"C:\Test\StaffReport.xlsx");
```

## 3. Full control with EPPlus and `AddWorksheet`

If you need a more custom workbook, use EPPlus directly and let `TPJ.Excel` help with navigation.

`AddWorksheet` returns an `ExcelWorksheetExtension`, which tracks the current row and column for you.
It also includes `Write(...)`, `WriteNext(...)`, `WriteHeader(...)`, and `WriteRow(...)` helpers to cut down repetitive code.

### Example

```csharp
using OfficeOpenXml;
using TPJ.Excel;

using var package = new ExcelPackage();
var ws = package.Workbook.AddWorksheet("Staff");

ws.WriteHeader("Staff #", "Name", "Start Date", "Salary");

foreach (var item in staff)
{
    ws.WriteNext(item.Id)
      .WriteNext(item.Name)
      .WriteNext(item.StartDate, dateFormat: "dd/mm/yyyy")
      .Write(item.Salary, numberFormat: "#,##0.00");

    ws.NextRow();
}

package.SaveAs(new FileInfo(@"C:\Test\StaffAdvanced.xlsx"));
```

### Useful navigation methods

- `CurrentCell()` gets the current cell
- `NextColumn()` moves one column to the right
- `NextRow()` moves to the next row and resets the column back to the first column
- `PreviousColumn()` and `PreviousRow()` let you move backwards
- `Reset()` returns to cell `A1`

## Summary

- Use `ExcelDocument` for the fastest export
- Use `EPPlusHelper` for clean, readable exports with headers and formatting
- Use full EPPlus support when you need a custom layout or advanced workbook logic

For complete working examples, see `TPJ.ExcelTest\Program.cs`.
