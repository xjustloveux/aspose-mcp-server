using Aspose.Cells;
using AsposeMcpServer.Handlers.Excel.ViewSettings;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

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

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        if (!IsEvaluationMode(AsposeLibraryType.Cells))
        {
            var ws = workbook.Worksheets[0];
            var expectedCell = CellsHelper.CellIndexToName(5, 3);
            Assert.Equal(expectedCell, ws.ActiveCell);
        }

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

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        if (!IsEvaluationMode(AsposeLibraryType.Cells))
        {
            var ws = workbook.Worksheets[0];
            var expectedCell = CellsHelper.CellIndexToName(10, 0);
            Assert.Equal(expectedCell, ws.ActiveCell);
        }

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

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        if (!IsEvaluationMode(AsposeLibraryType.Cells))
        {
            var ws = workbook.Worksheets[0];
            Assert.NotEqual(PaneStateType.Split, ws.PaneState);
        }

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

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        if (!IsEvaluationMode(AsposeLibraryType.Cells))
        {
            var ws = workbook.Worksheets[1];
            var expectedCell = CellsHelper.CellIndexToName(5, 0);
            Assert.Equal(expectedCell, ws.ActiveCell);
        }
    }

    #endregion
}
