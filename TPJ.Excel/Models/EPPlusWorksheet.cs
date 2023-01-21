namespace TPJ.Excel.Models;

public class EPPlusWorksheet
{
    public string Name { get; set; }
    public IEnumerable<string> Headers { get; set; }
    public IEnumerable<EPPlusRow> Rows { get; set; }

    public EPPlusWorksheet()
    {
    }

    public EPPlusWorksheet(string name)
    {
        Name = name;
    }
}
