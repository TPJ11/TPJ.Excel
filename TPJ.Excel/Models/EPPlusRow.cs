namespace TPJ.Excel.Models;

public class EPPlusRow
{
    public IEnumerable<EPPlusData> Data { get; set; }

    public EPPlusRow()
    {
    }

    public EPPlusRow(IEnumerable<EPPlusData> data)
    {
        Data = data;
    }
}
