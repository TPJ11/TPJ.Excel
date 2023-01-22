# TPJ.Excel
Simple helper library that covers almost all scenarios where you need to create an excel file.

# Examples
There are three main types of excel scenarios that this library covers. Take a look at `TPJ.ExcelTest` within the code to see a simple example of each.

## Super Simple
If all you want to do is create an excel file using the given list of data and are not worried about the format or the headings then you can achieve this by calling `ExcelDocument.Create` there are a number of overloads for this method but they simply group down to do you want to save the file to disk or get the file as bytes?

### Save to disk
`ExcelDocument.Create(staff, @"C:\Test\SimpleExcelDocument.xlsx");`

### Get bytes
`var excelBytes = ExcelDocument.Create(staff);`

## Some control
If you want to have some control over the headings and the format of the dates and numbers then you can use the `EPPlusHelper` class's and methods. This basically wraps what you can do in the more complex example below in an easy to read and understand format, very useful for 90% of scenarios you are asked to create an excel file for.

First you need to create a worksheet this is done using `EPPlusWorksheet` class. Within the worksheet you have three properties:
1. Name - This is the name of the worksheet (tabs at the bottom of excel)
2. Headers - This is the list of names you want for your headers in order of the data you are adding
3. Rows - This contains your data

The row object `EPPlusRow` has one property `Data` this contains the data to add to the row in the same order as your headings.

The data object `EPPlusData` has three properties:
1. Value - Contains the value that will go into the cell
2. DateFormat - If set the format of the value e.g. `dd/mm/yyyy`
2. NumberFormat - If set the format of the value e.g. `#0.00`

Both DateFormat and NumberFormat apply to the same propery so they are the same the only reason to have them as two properties is to make it clear to people reading your code. Formats can be found on the [EPPlus documentation](https://github.com/EPPlusSoftware/EPPlus/wiki/Formatting-and-styling)

```
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
}, @"C:\Test\SimpleEPPlus.xlsx");
```

## Full Control
Lastly if you have a more complex file you need to create then you can use the full power of EPPlus (note TPJ.Excel uses the last open source EPPlus version 4.5.3.3, V5+ of EPPlus is a paid for product).

Calling the extension method `AddWorksheet` on the workbook returns an object containing the worksheet that has the ability to track the currently 'selected' cell to make moving though the worksheet simple and clean.

Note - the below will produce the same as the 'Some control' above

```
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

p.SaveAs(new FileInfo(@"C:\Test\ComplexEPPlus.xlsx"));
```