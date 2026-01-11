using AsposeMcpServer.Handlers.Excel.ViewSettings;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Excel.ViewSettings;

public class SetHeadersExcelViewHandlerTests : ExcelHandlerTestBase
{
    private readonly SetHeadersExcelViewHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_SetHeaders()
    {
        Assert.Equal("set_headers", _handler.Operation);
    }

    #endregion

    #region Basic Set Operations

    [Fact]
    public void Execute_ShowsHeaders()
    {
        var workbook = CreateEmptyWorkbook();
        workbook.Worksheets[0].IsRowColumnHeadersVisible = false;
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "visible", true }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("visible", result.ToLower());
        Assert.True(workbook.Worksheets[0].IsRowColumnHeadersVisible);
        AssertModified(context);
    }

    [Fact]
    public void Execute_HidesHeaders()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "visible", false }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("hidden", result.ToLower());
        Assert.False(workbook.Worksheets[0].IsRowColumnHeadersVisible);
    }

    [Fact]
    public void Execute_WithSheetIndex_SetsHeadersOnSpecificSheet()
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

        Assert.False(workbook.Worksheets[1].IsRowColumnHeadersVisible);
    }

    #endregion
}
