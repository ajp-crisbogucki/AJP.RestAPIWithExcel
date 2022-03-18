using System.Dynamic;
using System.Net.Security;
using System.Reflection;
using Ajp.RestApi.Models;
using OfficeOpenXml;

namespace Ajp.RestApi.Storage;

public static class ExcelStorage
{
    private static byte[]? ExcelDataRaw { get; set; }
    public static List<ExcelModel> Data { get; set; }

    public static void Load()
    {
        ExcelDataRaw = File.ReadAllBytes("sample-xlsx-file-for-testing.xlsx");
        LoadDataFromExcelFile();
    }

    public static void Remove(int id)
    {
        id++;
        using MemoryStream stream = new MemoryStream(ExcelDataRaw);
        using (var pack = new ExcelPackage(stream))
        {
            var ws = pack.Workbook.Worksheets.FirstOrDefault();
            ws.DeleteRow(id, id, true);
            pack.SaveAs(new FileInfo("sample-xlsx-file-for-testing.xlsx"));
        }
    }

    public static void LoadDataFromExcelFile()
    {
        Data = new List<ExcelModel>();
        Data.Clear();

        using MemoryStream stream = new MemoryStream(ExcelDataRaw);
        using ExcelPackage excelPackage = new ExcelPackage(stream);
        var id = 0;

        //loop all worksheets
        foreach (var worksheet in excelPackage.Workbook.Worksheets)
        {
            //loop all rows
            for (int i = worksheet.Dimension.Start.Row; i <= worksheet.Dimension.End.Row; i++)
            {
                var row = new ExcelModel();

                //loop all columns in a row
                for (var j = worksheet.Dimension.Start.Column; j <= worksheet.Dimension.End.Column; j++)
                {
                    if (id <= 0) continue;
                    row.Id = id;
                    var propName = worksheet.Cells[1, j].Text.Trim().Replace(" ", "");
                    var val = worksheet.Cells[i, j].Value.ToString();
                    SetPropValue(row, propName, val);
                }

                if (id > 0)
                    Data.Add(row);

                id++;
            }
        }
    }


    private static void SetPropValue(object src, string propName, object? val)
    {
        src.GetType().GetProperty(propName)?.SetValue(src, val);
    }

    public static void Add(ExcelModel model)
    {
        using MemoryStream stream = new MemoryStream(ExcelDataRaw);
        using ExcelPackage excelPackage = new ExcelPackage(stream);
        var i = Data.Count + 1;
        //loop all worksheets
        foreach (var worksheet in excelPackage.Workbook.Worksheets)
        {
            //loop all columns in a row
            for (var j = worksheet.Dimension.Start.Column; j <= worksheet.Dimension.End.Column; j++)
            {
                foreach (var prop in model.GetType().GetProperties())
                {
                    var propName = worksheet.Cells[1, j].Text.Trim().Replace(" ", "");
                    if (prop.Name == propName)
                    {
                        worksheet.Cells[i, j].Value = prop.GetValue(model)?.ToString();
                    }
                }
            }
        }
        
        excelPackage.SaveAs(new FileInfo("sample-xlsx-file-for-testing.xlsx"));
    }
}