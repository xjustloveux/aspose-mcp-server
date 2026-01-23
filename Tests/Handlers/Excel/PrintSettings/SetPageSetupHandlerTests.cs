using Aspose.Cells;
using AsposeMcpServer.Handlers.Excel.PrintSettings;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Excel.PrintSettings;

public class SetPageSetupHandlerTests : ExcelHandlerTestBase
{
    private readonly SetPageSetupHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_SetPageSetup()
    {
        Assert.Equal("set_page_setup", _handler.Operation);
    }

    #endregion

    #region Basic Set Page Setup Operations

    [Fact]
    public void Execute_SetsOrientation()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "orientation", "landscape" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("page setup updated", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(PageOrientationType.Landscape, workbook.Worksheets[0].PageSetup.Orientation);
        AssertModified(context);
    }

    [Fact]
    public void Execute_SetsMargins()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "leftMargin", 1.0 },
            { "rightMargin", 1.0 },
            { "topMargin", 0.5 },
            { "bottomMargin", 0.5 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("page setup updated", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.True(Math.Abs(1.0 - workbook.Worksheets[0].PageSetup.LeftMargin) < 0.01);
        Assert.True(Math.Abs(1.0 - workbook.Worksheets[0].PageSetup.RightMargin) < 0.01);
    }

    [Fact]
    public void Execute_SetsHeaderAndFooter()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "header", "&CHeader Text" },
            { "footer", "&CPage &P" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("page setup updated", result.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Execute_SetsFitToPage()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "fitToPage", true },
            { "fitToPagesWide", 1 },
            { "fitToPagesTall", 2 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("page setup updated", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(1, workbook.Worksheets[0].PageSetup.FitToPagesWide);
        Assert.Equal(2, workbook.Worksheets[0].PageSetup.FitToPagesTall);
    }

    [Fact]
    public void Execute_WithSheetIndex_UpdatesSpecificSheet()
    {
        var workbook = CreateWorkbookWithSheets(2);
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sheetIndex", 1 },
            { "orientation", "portrait" }
        });

        _handler.Execute(context, parameters);

        Assert.Equal(PageOrientationType.Portrait, workbook.Worksheets[1].PageSetup.Orientation);
    }

    [Fact]
    public void Execute_WithNoParameters_ReportsNoChanges()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("no changes", result.Message, StringComparison.OrdinalIgnoreCase);
    }

    #endregion
}
