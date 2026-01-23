using AsposeMcpServer.Handlers.Excel.ViewSettings;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Excel.ViewSettings;

public class SetGridlinesExcelViewHandlerTests : ExcelHandlerTestBase
{
    private readonly SetGridlinesExcelViewHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_SetGridlines()
    {
        Assert.Equal("set_gridlines", _handler.Operation);
    }

    #endregion

    #region Basic Set Operations

    [Fact]
    public void Execute_ShowsGridlines()
    {
        var workbook = CreateEmptyWorkbook();
        workbook.Worksheets[0].IsGridlinesVisible = false;
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "visible", true }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("visible", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.True(workbook.Worksheets[0].IsGridlinesVisible);
        AssertModified(context);
    }

    [Fact]
    public void Execute_HidesGridlines()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "visible", false }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("hidden", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.False(workbook.Worksheets[0].IsGridlinesVisible);
    }

    [Fact]
    public void Execute_WithSheetIndex_SetsGridlinesOnSpecificSheet()
    {
        var workbook = CreateEmptyWorkbook();
        workbook.Worksheets.Add("Sheet2");
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sheetIndex", 1 },
            { "visible", false }
        });

        _handler.Execute(context, parameters);

        Assert.False(workbook.Worksheets[1].IsGridlinesVisible);
    }

    #endregion
}
