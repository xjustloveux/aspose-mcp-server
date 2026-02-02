using Aspose.Cells;
using AsposeMcpServer.Handlers.Excel.ViewSettings;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Excel.ViewSettings;

public class AutoFitRowExcelViewHandlerTests : ExcelHandlerTestBase
{
    private readonly AutoFitRowExcelViewHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_AutoFitRow()
    {
        Assert.Equal("auto_fit_row", _handler.Operation);
    }

    #endregion

    #region Helper Methods

    private static Workbook CreateWorkbookWithData()
    {
        var workbook = new Workbook();
        var ws = workbook.Worksheets[0];
        var style = ws.Cells["A1"].GetStyle();
        style.IsTextWrapped = true;
        ws.Cells["A1"].SetStyle(style);
        ws.Cells["A1"].PutValue("This is a long text\nwith multiple lines\nfor testing auto fit row");
        return workbook;
    }

    #endregion

    #region Basic AutoFit Operations

    [Fact]
    public void Execute_AutoFitsRow()
    {
        var workbook = CreateWorkbookWithData();
        var heightBefore = workbook.Worksheets[0].Cells.GetRowHeight(0);
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "rowIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        if (!IsEvaluationMode(AsposeLibraryType.Cells))
        {
            var heightAfter = workbook.Worksheets[0].Cells.GetRowHeight(0);
            Assert.True(heightAfter > heightBefore,
                $"Row height should increase after auto-fit with wrapped text. Before: {heightBefore}, After: {heightAfter}");
        }

        AssertModified(context);
    }

    [Fact]
    public void Execute_WithColumnRange_AutoFitsRow()
    {
        var workbook = CreateWorkbookWithData();
        var heightBefore = workbook.Worksheets[0].Cells.GetRowHeight(0);
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "rowIndex", 0 },
            { "startColumn", 0 },
            { "endColumn", 3 }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        if (!IsEvaluationMode(AsposeLibraryType.Cells))
        {
            var heightAfter = workbook.Worksheets[0].Cells.GetRowHeight(0);
            Assert.True(heightAfter > heightBefore,
                $"Row height should increase after auto-fit with column range. Before: {heightBefore}, After: {heightAfter}");
        }

        AssertModified(context);
    }

    [Fact]
    public void Execute_WithSheetIndex_AutoFitsOnSpecificSheet()
    {
        var workbook = CreateWorkbookWithData();
        workbook.Worksheets.Add("Sheet2");
        var style = workbook.Worksheets[1].Cells["A1"].GetStyle();
        style.IsTextWrapped = true;
        workbook.Worksheets[1].Cells["A1"].SetStyle(style);
        workbook.Worksheets[1].Cells["A1"].PutValue("Long text\nwith multiple lines\nfor auto fit row");
        var heightBefore = workbook.Worksheets[1].Cells.GetRowHeight(0);
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sheetIndex", 1 },
            { "rowIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        if (!IsEvaluationMode(AsposeLibraryType.Cells))
        {
            var heightAfter = workbook.Worksheets[1].Cells.GetRowHeight(0);
            Assert.True(heightAfter > heightBefore,
                $"Row height on Sheet2 should increase after auto-fit. Before: {heightBefore}, After: {heightAfter}");
        }
    }

    #endregion
}
