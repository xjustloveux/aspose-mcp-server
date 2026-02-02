using Aspose.Cells;
using AsposeMcpServer.Handlers.Excel.ViewSettings;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Excel.ViewSettings;

public class FreezePanesExcelViewHandlerTests : ExcelHandlerTestBase
{
    private readonly FreezePanesExcelViewHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_FreezePanes()
    {
        Assert.Equal("freeze_panes", _handler.Operation);
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

    private static Workbook CreateWorkbookWithFrozenPanes()
    {
        var workbook = new Workbook();
        workbook.Worksheets[0].FreezePanes(2, 2, 1, 1);
        return workbook;
    }

    #endregion

    #region Basic Freeze Operations

    [Fact]
    public void Execute_FreezesPanes()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "freezeRow", 1 },
            { "freezeColumn", 1 }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        if (!IsEvaluationMode(AsposeLibraryType.Cells))
        {
            var ws = workbook.Worksheets[0];
            Assert.Equal(PaneStateType.Frozen, ws.PaneState);
            ws.GetFreezedPanes(out _, out _, out var frozenRows, out var frozenCols);
            Assert.Equal(1, frozenRows);
            Assert.Equal(1, frozenCols);
        }

        AssertModified(context);
    }

    [Fact]
    public void Execute_WithFreezeRowOnly_FreezesPanes()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "freezeRow", 2 }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        if (!IsEvaluationMode(AsposeLibraryType.Cells))
        {
            var ws = workbook.Worksheets[0];
            Assert.Equal(PaneStateType.Frozen, ws.PaneState);
            ws.GetFreezedPanes(out _, out _, out var frozenRows, out var frozenCols);
            Assert.Equal(2, frozenRows);
            Assert.Equal(0, frozenCols);
        }

        AssertModified(context);
    }

    [Fact]
    public void Execute_UnfreezesPanes()
    {
        var workbook = CreateWorkbookWithFrozenPanes();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "unfreeze", true }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        if (!IsEvaluationMode(AsposeLibraryType.Cells))
        {
            var ws = workbook.Worksheets[0];
            Assert.NotEqual(PaneStateType.Frozen, ws.PaneState);
        }

        AssertModified(context);
    }

    [Fact]
    public void Execute_WithSheetIndex_FreezesOnSpecificSheet()
    {
        var workbook = CreateEmptyWorkbook();
        workbook.Worksheets.Add("Sheet2");
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sheetIndex", 1 },
            { "freezeRow", 3 }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        if (!IsEvaluationMode(AsposeLibraryType.Cells))
        {
            var ws = workbook.Worksheets[1];
            Assert.Equal(PaneStateType.Frozen, ws.PaneState);
            ws.GetFreezedPanes(out _, out _, out var frozenRows, out var frozenCols);
            Assert.Equal(3, frozenRows);
            Assert.Equal(0, frozenCols);
        }
    }

    #endregion
}
