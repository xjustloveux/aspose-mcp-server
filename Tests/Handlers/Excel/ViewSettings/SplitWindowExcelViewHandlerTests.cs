using Aspose.Cells;
using AsposeMcpServer.Handlers.Excel.ViewSettings;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Excel.ViewSettings;

public class SplitWindowExcelViewHandlerTests : ExcelHandlerTestBase
{
    private readonly SplitWindowExcelViewHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_SplitWindow()
    {
        Assert.Equal("split_window", _handler.Operation);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithNoParameters_ThrowsArgumentException()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion

    #region Helper Methods

    private static Workbook CreateWorkbookWithSplit()
    {
        var workbook = new Workbook();
        workbook.Worksheets[0].ActiveCell = "E5";
        workbook.Worksheets[0].Split();
        return workbook;
    }

    #endregion

    #region Basic Split Operations

    [Fact]
    public void Execute_SplitsWindow()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "splitRow", 5 },
            { "splitColumn", 3 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("split", result.ToLower());
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithSplitRowOnly_SplitsWindow()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "splitRow", 10 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("split", result.ToLower());
        Assert.Contains("row 10", result.ToLower());
        AssertModified(context);
    }

    [Fact]
    public void Execute_RemovesSplit()
    {
        var workbook = CreateWorkbookWithSplit();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "removeSplit", true }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("removed", result.ToLower());
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithSheetIndex_SplitsOnSpecificSheet()
    {
        var workbook = CreateEmptyWorkbook();
        workbook.Worksheets.Add("Sheet2");
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sheetIndex", 1 },
            { "splitRow", 5 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("split", result.ToLower());
    }

    #endregion
}
