using AsposeMcpServer.Handlers.Excel.ViewSettings;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Excel.ViewSettings;

public class SetTabColorExcelViewHandlerTests : ExcelHandlerTestBase
{
    private readonly SetTabColorExcelViewHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_SetTabColor()
    {
        Assert.Equal("set_tab_color", _handler.Operation);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutColor_ThrowsArgumentException()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion

    #region Basic Set Operations

    [Fact]
    public void Execute_SetsTabColor()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "color", "Red" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("red", result.ToLower());
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithSheetIndex_SetsTabColorOnSpecificSheet()
    {
        var workbook = CreateEmptyWorkbook();
        workbook.Worksheets.Add("Sheet2");
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sheetIndex", 1 },
            { "color", "Blue" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("blue", result.ToLower());
    }

    [Fact]
    public void Execute_WithHexColor_SetsTabColor()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "color", "#FF5733" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("#FF5733", result);
        AssertModified(context);
    }

    #endregion
}
