using Aspose.Cells;
using Aspose.Cells.Tables;
using AsposeMcpServer.Handlers.Excel.Table;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Excel.Table;

public class AddTotalRowExcelTableHandlerTests : ExcelHandlerTestBase
{
    private readonly AddTotalRowExcelTableHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_ShouldBeAddTotalRow()
    {
        Assert.Equal("add_total_row", _handler.Operation);
    }

    #endregion

    #region Helper Methods

    private static Workbook CreateWorkbookWithTable()
    {
        var workbook = CreateWorkbookWithData(new object[,]
        {
            { "Name", "Value", "Count" },
            { "A", 1, 10 },
            { "B", 2, 20 },
            { "C", 3, 30 }
        });
        workbook.Worksheets[0].ListObjects.Add("A1", "C4", true);
        return workbook;
    }

    #endregion

    #region Basic Add Total Row Operations

    [Fact]
    public void Execute_ShouldEnableTotalRow()
    {
        var workbook = CreateWorkbookWithTable();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "tableIndex", 0 }
        });

        _handler.Execute(context, parameters);

        Assert.True(workbook.Worksheets[0].ListObjects[0].ShowTotals);
    }

    [Fact]
    public void Execute_WithColumnFunction_ShouldSetTotalsCalculation()
    {
        var workbook = CreateWorkbookWithTable();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "tableIndex", 0 },
            { "columnIndex", 1 },
            { "totalFunction", "sum" }
        });

        _handler.Execute(context, parameters);

        var listObject = workbook.Worksheets[0].ListObjects[0];
        Assert.True(listObject.ShowTotals);
        Assert.Equal(TotalsCalculation.Sum, listObject.ListColumns[1].TotalsCalculation);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithMissingTableIndex_ShouldThrow()
    {
        var workbook = CreateWorkbookWithTable();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("tableIndex", ex.Message);
    }

    [Fact]
    public void Execute_WithInvalidColumnIndex_ShouldThrow()
    {
        var workbook = CreateWorkbookWithTable();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "tableIndex", 0 },
            { "columnIndex", 99 },
            { "totalFunction", "sum" }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion

    #region ResolveTotalsCalculation

    [Theory]
    [InlineData("sum", TotalsCalculation.Sum)]
    [InlineData("count", TotalsCalculation.Count)]
    [InlineData("average", TotalsCalculation.Average)]
    [InlineData("max", TotalsCalculation.Max)]
    [InlineData("min", TotalsCalculation.Min)]
    public void ResolveTotalsCalculation_WithValidNames_ShouldReturn(string functionName,
        TotalsCalculation expected)
    {
        var result = AddTotalRowExcelTableHandler.ResolveTotalsCalculation(functionName);

        Assert.Equal(expected, result);
    }

    [Fact]
    public void ResolveTotalsCalculation_WithInvalidName_ShouldThrow()
    {
        Assert.Throws<ArgumentException>(() =>
            AddTotalRowExcelTableHandler.ResolveTotalsCalculation("invalid_function"));
    }

    #endregion
}
