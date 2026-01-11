using System.Text.Json;
using Aspose.Cells;
using AsposeMcpServer.Handlers.Excel.Formula;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Excel.Formula;

public class GetFormulasHandlerTests : ExcelHandlerTestBase
{
    private readonly GetFormulasHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Get()
    {
        Assert.Equal("get", _handler.Operation);
    }

    #endregion

    #region Sheet Index

    [Fact]
    public void Execute_WithSheetIndex_GetsFormulasFromCorrectSheet()
    {
        var workbook = CreateWorkbookWithFormulas();
        workbook.Worksheets.Add("Sheet2");
        workbook.Worksheets[1].Cells["A1"].Formula = "=10+20";
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sheetIndex", 1 }
        });

        var result = _handler.Execute(context, parameters);
        var json = JsonDocument.Parse(result);

        Assert.Equal(1, json.RootElement.GetProperty("count").GetInt32());
    }

    #endregion

    #region No Formulas

    [Fact]
    public void Execute_NoFormulas_ReturnsEmptyResult()
    {
        var workbook = CreateWorkbookWithData(new object[,]
        {
            { 10, 20, 30 }
        });
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);
        var json = JsonDocument.Parse(result);

        Assert.Equal(0, json.RootElement.GetProperty("count").GetInt32());
        Assert.Contains("No formulas found", json.RootElement.GetProperty("message").GetString());
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithInvalidRange_ThrowsArgumentException()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "range", "InvalidRange" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("Invalid range", ex.Message);
    }

    #endregion

    #region Helper Methods

    private static Workbook CreateWorkbookWithFormulas()
    {
        var workbook = new Workbook();
        var sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].Value = 10;
        sheet.Cells["A2"].Value = 20;
        sheet.Cells["B1"].Value = 30;
        sheet.Cells["B2"].Value = 40;
        sheet.Cells["C1"].Formula = "=A1+B1";
        sheet.Cells["C2"].Formula = "=A2+B2";
        sheet.Cells["C3"].Formula = "=SUM(C1:C2)";
        workbook.CalculateFormula();
        return workbook;
    }

    #endregion

    #region Basic Get Operations

    [Fact]
    public void Execute_GetsFormulas()
    {
        var workbook = CreateWorkbookWithFormulas();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);
        var json = JsonDocument.Parse(result);

        Assert.True(json.RootElement.GetProperty("count").GetInt32() > 0);
    }

    [Fact]
    public void Execute_ReturnsCorrectCount()
    {
        var workbook = CreateWorkbookWithFormulas();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);
        var json = JsonDocument.Parse(result);

        Assert.Equal(3, json.RootElement.GetProperty("count").GetInt32());
    }

    [Fact]
    public void Execute_ReturnsItems()
    {
        var workbook = CreateWorkbookWithFormulas();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);
        var json = JsonDocument.Parse(result);

        Assert.Equal(3, json.RootElement.GetProperty("items").GetArrayLength());
    }

    [Fact]
    public void Execute_ReturnsWorksheetName()
    {
        var workbook = CreateWorkbookWithFormulas();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);
        var json = JsonDocument.Parse(result);

        Assert.NotNull(json.RootElement.GetProperty("worksheetName").GetString());
    }

    #endregion

    #region Range Filtering

    [Fact]
    public void Execute_WithRange_ReturnsOnlyFormulasInRange()
    {
        var workbook = CreateWorkbookWithFormulas();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "range", "C1:C2" }
        });

        var result = _handler.Execute(context, parameters);
        var json = JsonDocument.Parse(result);

        Assert.Equal(2, json.RootElement.GetProperty("count").GetInt32());
    }

    [Fact]
    public void Execute_WithRangeNoFormulas_ReturnsEmptyResult()
    {
        var workbook = CreateWorkbookWithFormulas();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "range", "D1:D10" }
        });

        var result = _handler.Execute(context, parameters);
        var json = JsonDocument.Parse(result);

        Assert.Equal(0, json.RootElement.GetProperty("count").GetInt32());
        Assert.Contains("No formulas found", json.RootElement.GetProperty("message").GetString());
    }

    #endregion
}
