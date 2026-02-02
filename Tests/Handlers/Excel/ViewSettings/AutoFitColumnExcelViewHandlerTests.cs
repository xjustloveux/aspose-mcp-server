using Aspose.Cells;
using AsposeMcpServer.Handlers.Excel.ViewSettings;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Excel.ViewSettings;

public class AutoFitColumnExcelViewHandlerTests : ExcelHandlerTestBase
{
    private readonly AutoFitColumnExcelViewHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_AutoFitColumn()
    {
        Assert.Equal("auto_fit_column", _handler.Operation);
    }

    #endregion

    #region Helper Methods

    private static Workbook CreateWorkbookWithData()
    {
        var workbook = new Workbook();
        var ws = workbook.Worksheets[0];
        ws.Cells["A1"].PutValue("This is a long text for testing auto fit");
        ws.Cells["A2"].PutValue("Short");
        ws.Cells["A3"].PutValue("Medium length text");
        return workbook;
    }

    #endregion

    #region Basic AutoFit Operations

    [Fact]
    public void Execute_AutoFitsColumn()
    {
        var workbook = CreateWorkbookWithData();
        var widthBefore = workbook.Worksheets[0].Cells.GetColumnWidth(0);
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "columnIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        if (!IsEvaluationMode(AsposeLibraryType.Cells))
        {
            var widthAfter = workbook.Worksheets[0].Cells.GetColumnWidth(0);
            Assert.True(widthAfter > widthBefore,
                $"Column width should increase after auto-fit. Before: {widthBefore}, After: {widthAfter}");
        }

        AssertModified(context);
    }

    [Fact]
    public void Execute_WithRowRange_AutoFitsColumn()
    {
        var workbook = CreateWorkbookWithData();
        var widthBefore = workbook.Worksheets[0].Cells.GetColumnWidth(0);
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "columnIndex", 0 },
            { "startRow", 0 },
            { "endRow", 5 }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        if (!IsEvaluationMode(AsposeLibraryType.Cells))
        {
            var widthAfter = workbook.Worksheets[0].Cells.GetColumnWidth(0);
            Assert.True(widthAfter > widthBefore,
                $"Column width should increase after auto-fit with row range. Before: {widthBefore}, After: {widthAfter}");
        }

        AssertModified(context);
    }

    [Fact]
    public void Execute_WithSheetIndex_AutoFitsOnSpecificSheet()
    {
        var workbook = CreateWorkbookWithData();
        workbook.Worksheets.Add("Sheet2");
        workbook.Worksheets[1].Cells["A1"].PutValue("Long text for auto fit");
        var widthBefore = workbook.Worksheets[1].Cells.GetColumnWidth(0);
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sheetIndex", 1 },
            { "columnIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        if (!IsEvaluationMode(AsposeLibraryType.Cells))
        {
            var widthAfter = workbook.Worksheets[1].Cells.GetColumnWidth(0);
            Assert.True(widthAfter > widthBefore,
                $"Column width on Sheet2 should increase after auto-fit. Before: {widthBefore}, After: {widthAfter}");
        }
    }

    #endregion
}
