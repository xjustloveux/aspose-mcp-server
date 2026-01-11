using AsposeMcpServer.Handlers.Excel.PrintSettings;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Excel.PrintSettings;

public class SetPrintAreaHandlerTests : ExcelHandlerTestBase
{
    private readonly SetPrintAreaHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_SetPrintArea()
    {
        Assert.Equal("set_print_area", _handler.Operation);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutRangeOrClear_ThrowsArgumentException()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion

    #region Basic Set Print Area Operations

    [Fact]
    public void Execute_SetsPrintArea()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "range", "A1:D10" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("print area set", result.ToLower());
        Assert.Contains("A1:D10", result);
        Assert.Equal("A1:D10", workbook.Worksheets[0].PageSetup.PrintArea);
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithSheetIndex_SetsPrintAreaInSpecificSheet()
    {
        var workbook = CreateWorkbookWithSheets(2);
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sheetIndex", 1 },
            { "range", "B2:E5" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("sheet 1", result.ToLower());
        Assert.Equal("B2:E5", workbook.Worksheets[1].PageSetup.PrintArea);
    }

    [Fact]
    public void Execute_WithClearPrintArea_ClearsPrintArea()
    {
        var workbook = CreateEmptyWorkbook();
        workbook.Worksheets[0].PageSetup.PrintArea = "A1:D10";
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "clearPrintArea", true }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("cleared", result.ToLower());
        Assert.True(string.IsNullOrEmpty(workbook.Worksheets[0].PageSetup.PrintArea));
    }

    #endregion
}
