using AsposeMcpServer.Handlers.Excel.ViewSettings;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Excel.ViewSettings;

public class SetAllExcelViewHandlerTests : ExcelHandlerTestBase
{
    private readonly SetAllExcelViewHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_SetAll()
    {
        Assert.Equal("set_all", _handler.Operation);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithInvalidZoom_ThrowsArgumentException()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "zoom", 5 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion

    #region Basic Set Operations

    [Fact]
    public void Execute_SetsAllViewSettings()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "zoom", 150 },
            { "showGridlines", false },
            { "showRowColumnHeaders", false },
            { "showZeroValues", false }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("updated", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(150, workbook.Worksheets[0].Zoom);
        Assert.False(workbook.Worksheets[0].IsGridlinesVisible);
        Assert.False(workbook.Worksheets[0].IsRowColumnHeadersVisible);
        Assert.False(workbook.Worksheets[0].DisplayZeros);
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithSheetIndex_SetsOnSpecificSheet()
    {
        var workbook = CreateEmptyWorkbook();
        workbook.Worksheets.Add("Sheet2");
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sheetIndex", 1 },
            { "zoom", 200 },
            { "showGridlines", false }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("updated", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(200, workbook.Worksheets[1].Zoom);
        Assert.False(workbook.Worksheets[1].IsGridlinesVisible);
    }

    [Fact]
    public void Execute_WithDisplayRightToLeft_SetsDirection()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "displayRightToLeft", true }
        });

        _handler.Execute(context, parameters);

        Assert.True(workbook.Worksheets[0].DisplayRightToLeft);
        AssertModified(context);
    }

    #endregion
}
