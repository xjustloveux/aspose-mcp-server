using AsposeMcpServer.Handlers.Excel.PrintSettings;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Excel.PrintSettings;

public class SetPrintTitlesHandlerTests : ExcelHandlerTestBase
{
    private readonly SetPrintTitlesHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_SetPrintTitles()
    {
        Assert.Equal("set_print_titles", _handler.Operation);
    }

    #endregion

    #region Basic Set Print Titles Operations

    [Fact]
    public void Execute_SetsPrintTitleRows()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "rows", "$1:$2" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("print titles updated", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Equal("$1:$2", workbook.Worksheets[0].PageSetup.PrintTitleRows);
        AssertModified(context);
    }

    [Fact]
    public void Execute_SetsPrintTitleColumns()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "columns", "$A:$B" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("print titles updated", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Equal("$A:$B", workbook.Worksheets[0].PageSetup.PrintTitleColumns);
    }

    [Fact]
    public void Execute_SetsBothRowsAndColumns()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "rows", "$1:$1" },
            { "columns", "$A:$A" }
        });

        _handler.Execute(context, parameters);

        Assert.Equal("$1:$1", workbook.Worksheets[0].PageSetup.PrintTitleRows);
        Assert.Equal("$A:$A", workbook.Worksheets[0].PageSetup.PrintTitleColumns);
    }

    [Fact]
    public void Execute_WithClearTitles_ClearsPrintTitles()
    {
        var workbook = CreateEmptyWorkbook();
        workbook.Worksheets[0].PageSetup.PrintTitleRows = "$1:$2";
        workbook.Worksheets[0].PageSetup.PrintTitleColumns = "$A:$B";
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "clearTitles", true }
        });

        _handler.Execute(context, parameters);

        Assert.True(string.IsNullOrEmpty(workbook.Worksheets[0].PageSetup.PrintTitleRows));
        Assert.True(string.IsNullOrEmpty(workbook.Worksheets[0].PageSetup.PrintTitleColumns));
    }

    [Fact]
    public void Execute_WithSheetIndex_UpdatesSpecificSheet()
    {
        var workbook = CreateWorkbookWithSheets(2);
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sheetIndex", 1 },
            { "rows", "$1:$3" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("sheet 1", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Equal("$1:$3", workbook.Worksheets[1].PageSetup.PrintTitleRows);
    }

    #endregion
}
