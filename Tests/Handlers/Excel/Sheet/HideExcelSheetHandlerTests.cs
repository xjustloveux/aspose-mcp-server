using AsposeMcpServer.Handlers.Excel.Sheet;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Excel.Sheet;

public class HideExcelSheetHandlerTests : ExcelHandlerTestBase
{
    private readonly HideExcelSheetHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Hide()
    {
        Assert.Equal("hide", _handler.Operation);
    }

    #endregion

    #region Preserve Other Sheets

    [Fact]
    public void Execute_PreservesOtherSheetsVisibility()
    {
        var workbook = CreateEmptyWorkbook();
        workbook.Worksheets.Add("Sheet2");
        workbook.Worksheets.Add("Sheet3");
        workbook.Worksheets[2].IsVisible = false;
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sheetIndex", 0 }
        });

        _handler.Execute(context, parameters);

        Assert.False(workbook.Worksheets[0].IsVisible);
        Assert.True(workbook.Worksheets[1].IsVisible);
        Assert.False(workbook.Worksheets[2].IsVisible);
    }

    #endregion

    #region Hide Operations

    [Fact]
    public void Execute_HidesVisibleSheet()
    {
        var workbook = CreateEmptyWorkbook();
        workbook.Worksheets.Add("Sheet2");
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sheetIndex", 0 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("hidden", result, StringComparison.OrdinalIgnoreCase);
        Assert.False(workbook.Worksheets[0].IsVisible);
        AssertModified(context);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    public void Execute_HidesSheetAtVariousIndices(int sheetIndex)
    {
        var workbook = CreateEmptyWorkbook();
        workbook.Worksheets.Add("Sheet2");
        workbook.Worksheets.Add("Sheet3");
        workbook.Worksheets.Add("Sheet4");
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sheetIndex", sheetIndex }
        });

        _handler.Execute(context, parameters);

        Assert.False(workbook.Worksheets[sheetIndex].IsVisible);
    }

    #endregion

    #region Show Operations

    [Fact]
    public void Execute_ShowsHiddenSheet()
    {
        var workbook = CreateEmptyWorkbook();
        workbook.Worksheets.Add("Sheet2");
        workbook.Worksheets[0].IsVisible = false;
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sheetIndex", 0 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("shown", result, StringComparison.OrdinalIgnoreCase);
        Assert.True(workbook.Worksheets[0].IsVisible);
        AssertModified(context);
    }

    [Fact]
    public void Execute_TogglesVisibility()
    {
        var workbook = CreateEmptyWorkbook();
        workbook.Worksheets.Add("Sheet2");
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sheetIndex", 0 }
        });

        _handler.Execute(context, parameters);
        Assert.False(workbook.Worksheets[0].IsVisible);

        _handler.Execute(context, parameters);
        Assert.True(workbook.Worksheets[0].IsVisible);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutSheetIndex_ThrowsArgumentException()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("sheetIndex", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData(-1)]
    [InlineData(10)]
    public void Execute_WithInvalidIndex_ThrowsArgumentException(int invalidIndex)
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sheetIndex", invalidIndex }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("index", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    #endregion

    #region Result Message

    [Fact]
    public void Execute_WhenHiding_ReturnsHiddenMessage()
    {
        var workbook = CreateEmptyWorkbook();
        workbook.Worksheets.Add("Sheet2");
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sheetIndex", 0 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("Sheet1", result);
        Assert.Contains("hidden", result, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Execute_WhenShowing_ReturnsShownMessage()
    {
        var workbook = CreateEmptyWorkbook();
        workbook.Worksheets.Add("Sheet2");
        workbook.Worksheets[0].IsVisible = false;
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sheetIndex", 0 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("Sheet1", result);
        Assert.Contains("shown", result, StringComparison.OrdinalIgnoreCase);
    }

    #endregion
}
