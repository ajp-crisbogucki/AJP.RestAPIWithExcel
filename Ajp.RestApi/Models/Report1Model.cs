namespace Ajp.RestApi.Models;

public class Report1Model
{
    public IEnumerable<ExcelModel> Items { get; set; }
    public decimal SummaryUnitsSold { get; set; }
}