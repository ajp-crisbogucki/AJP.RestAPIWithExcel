using Ajp.RestApi.Models;
using Ajp.RestApi.Storage;
using Microsoft.AspNetCore.Mvc;

namespace Ajp.RestApi.Controllers;

[ApiController]
[Route("[controller]")]
public class ExcelController : ControllerBase
{
    public ExcelController()
    {
        ExcelStorage.Load();
    }

    [HttpGet("get-by-segment/{segment}")]
    public IEnumerable<ExcelModel> GetBySegment(string segment)
    {
        return ExcelStorage.Data.Where(x => x.Segment == segment);
    }
    
    [HttpGet("get-by-country/{country}")]
    public IEnumerable<ExcelModel> GetByCountry(string country)
    {
        return ExcelStorage.Data.Where(x => x.Country == country);
    }
    
    [HttpGet("get-by-product/{product}")]
    public IEnumerable<ExcelModel> GetByProduct(string product)
    {
        return ExcelStorage.Data.Where(x => x.Product.ToLower() == product.ToLower());
    }
    
    [HttpGet("get-report-by-segment-and-country/{segment}/{country}")]
    public Report1Model GetReport(string segment, string country)
    {
        var result = new Report1Model();
        result.Items =  ExcelStorage.Data.Where(x => x.Segment == segment && x.Country == country);
        result.SummaryUnitsSold = result.Items.Select(x => Convert.ToDecimal(x.UnitsSold)).Sum();
        return result;
    }
    
    [HttpPost("add")]
    public ActionResult Add(ExcelModel model)
    {
        ExcelStorage.Add(model);
        return Ok();
    }
    
    [HttpDelete("remove-by-index/{index}")]
    public ActionResult RemoveByIndex(int index)
    {
        ExcelStorage.Remove(index);
        return Ok();
    }
    
    [HttpPost("search-by-model")]
    public ActionResult<ExcelModel> Search(ExcelModel model)
    {
        var result = ExcelStorage.Data.FirstOrDefault(x => x == model);
        return Ok(result);
    }
    
    
}