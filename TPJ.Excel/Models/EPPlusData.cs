namespace TPJ.Excel.Models;

public class EPPlusData
{
    public object Value { get; set; }
    public string? DateFormat { get; set; }
    public string? NumberFormat { get; set; }

    public EPPlusData()
    {
    }

    public EPPlusData(object value)
    {
        Value = value;
    }
}
