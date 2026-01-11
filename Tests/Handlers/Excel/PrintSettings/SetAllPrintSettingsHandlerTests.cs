using Aspose.Cells;
using AsposeMcpServer.Handlers.Excel.PrintSettings;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Excel.PrintSettings;

public class SetAllPrintSettingsHandlerTests : ExcelHandlerTestBase
{
    private readonly SetAllPrintSettingsHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_SetAll()
    {
        Assert.Equal("set_all", _handler.Operation);
    }

    #endregion

    #region Basic Set All Print Settings Operations

    [Fact]
    public void Execute_SetsAllPrintSettings()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "range", "A1:D10" },
            { "rows", "$1:$1" },
            { "columns", "$A:$A" },
            { "orientation", "landscape" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("print settings updated", result.ToLower());
        Assert.Equal("A1:D10", workbook.Worksheets[0].PageSetup.PrintArea);
        Assert.Equal("$1:$1", workbook.Worksheets[0].PageSetup.PrintTitleRows);
        Assert.Equal("$A:$A", workbook.Worksheets[0].PageSetup.PrintTitleColumns);
        Assert.Equal(PageOrientationType.Landscape, workbook.Worksheets[0].PageSetup.Orientation);
        AssertModified(context);
    }

    [Fact]
    public void Execute_SetsPrintAreaOnly()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "range", "B2:E5" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("printArea=B2:E5", result);
        Assert.Equal("B2:E5", workbook.Worksheets[0].PageSetup.PrintArea);
    }

    [Fact]
    public void Execute_SetsMarginsAndOrientation()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "leftMargin", 0.75 },
            { "rightMargin", 0.75 },
            { "topMargin", 1.0 },
            { "bottomMargin", 1.0 },
            { "orientation", "portrait" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("print settings updated", result.ToLower());
        Assert.Equal(PageOrientationType.Portrait, workbook.Worksheets[0].PageSetup.Orientation);
    }

    [Fact]
    public void Execute_WithSheetIndex_UpdatesSpecificSheet()
    {
        var workbook = CreateWorkbookWithSheets(2);
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sheetIndex", 1 },
            { "range", "A1:C5" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("sheet 1", result.ToLower());
        Assert.Equal("A1:C5", workbook.Worksheets[1].PageSetup.PrintArea);
    }

    [Fact]
    public void Execute_WithNoParameters_ReportsNoChanges()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("no changes", result.ToLower());
    }

    #endregion
}
