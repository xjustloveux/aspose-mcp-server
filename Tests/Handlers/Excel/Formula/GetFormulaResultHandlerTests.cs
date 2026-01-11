using System.Text.Json;
using Aspose.Cells;
using AsposeMcpServer.Handlers.Excel.Formula;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Excel.Formula;

public class GetFormulaResultHandlerTests : ExcelHandlerTestBase
{
    private readonly GetFormulaResultHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_GetResult()
    {
        Assert.Equal("get_result", _handler.Operation);
    }

    #endregion

    #region CalculateBeforeRead

    [Fact]
    public void Execute_WithCalculateBeforeReadFalse_DoesNotRecalculate()
    {
        var workbook = new Workbook();
        var sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].Value = 10;
        sheet.Cells["B1"].Value = 20;
        sheet.Cells["C1"].Formula = "=A1+B1";
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "cell", "C1" },
            { "calculateBeforeRead", false }
        });

        var result = _handler.Execute(context, parameters);
        var json = JsonDocument.Parse(result);

        Assert.Equal("=A1+B1", json.RootElement.GetProperty("formula").GetString());
    }

    #endregion

    #region Sheet Index

    [Fact]
    public void Execute_WithSheetIndex_GetsResultFromCorrectSheet()
    {
        var workbook = new Workbook();
        workbook.Worksheets.Add("Sheet2");
        workbook.Worksheets[1].Cells["A1"].Value = 100;
        workbook.Worksheets[1].Cells["B1"].Value = 200;
        workbook.Worksheets[1].Cells["C1"].Formula = "=A1+B1";
        workbook.CalculateFormula();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "cell", "C1" },
            { "sheetIndex", 1 }
        });

        var result = _handler.Execute(context, parameters);
        var json = JsonDocument.Parse(result);

        Assert.Equal("300", json.RootElement.GetProperty("calculatedValue").GetString());
    }

    #endregion

    #region Non-Formula Cells

    [Fact]
    public void Execute_NonFormulaCell_ReturnsNullFormula()
    {
        var workbook = CreateWorkbookWithData(new object[,]
        {
            { 10, 20, 30 }
        });
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "cell", "A1" }
        });

        var result = _handler.Execute(context, parameters);
        var json = JsonDocument.Parse(result);

        Assert.Null(json.RootElement.GetProperty("formula").GetString());
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutCell_ThrowsArgumentException()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("cell", ex.Message);
    }

    #endregion

    #region Helper Methods

    private static Workbook CreateWorkbookWithFormula()
    {
        var workbook = new Workbook();
        var sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].Value = 10;
        sheet.Cells["B1"].Value = 20;
        sheet.Cells["C1"].Formula = "=A1+B1";
        workbook.CalculateFormula();
        return workbook;
    }

    #endregion

    #region Basic Get Result Operations

    [Fact]
    public void Execute_GetsFormulaResult()
    {
        var workbook = CreateWorkbookWithFormula();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "cell", "C1" }
        });

        var result = _handler.Execute(context, parameters);
        var json = JsonDocument.Parse(result);

        Assert.Equal("C1", json.RootElement.GetProperty("cell").GetString());
    }

    [Fact]
    public void Execute_ReturnsFormula()
    {
        var workbook = CreateWorkbookWithFormula();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "cell", "C1" }
        });

        var result = _handler.Execute(context, parameters);
        var json = JsonDocument.Parse(result);

        Assert.Equal("=A1+B1", json.RootElement.GetProperty("formula").GetString());
    }

    [Fact]
    public void Execute_ReturnsCalculatedValue()
    {
        var workbook = CreateWorkbookWithFormula();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "cell", "C1" }
        });

        var result = _handler.Execute(context, parameters);
        var json = JsonDocument.Parse(result);

        Assert.Equal("30", json.RootElement.GetProperty("calculatedValue").GetString());
    }

    [Fact]
    public void Execute_ReturnsValueType()
    {
        var workbook = CreateWorkbookWithFormula();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "cell", "C1" }
        });

        var result = _handler.Execute(context, parameters);
        var json = JsonDocument.Parse(result);

        Assert.NotNull(json.RootElement.GetProperty("valueType").GetString());
    }

    #endregion
}
