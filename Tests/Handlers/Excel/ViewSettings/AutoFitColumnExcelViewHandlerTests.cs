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
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "columnIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("auto-fitted", result.Message, StringComparison.OrdinalIgnoreCase);
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithRowRange_AutoFitsColumn()
    {
        var workbook = CreateWorkbookWithData();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "columnIndex", 0 },
            { "startRow", 0 },
            { "endRow", 5 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("auto-fitted", result.Message, StringComparison.OrdinalIgnoreCase);
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithSheetIndex_AutoFitsOnSpecificSheet()
    {
        var workbook = CreateWorkbookWithData();
        workbook.Worksheets.Add("Sheet2");
        workbook.Worksheets[1].Cells["A1"].PutValue("Long text for auto fit");
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sheetIndex", 1 },
            { "columnIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("column 0", result.Message, StringComparison.OrdinalIgnoreCase);
    }

    #endregion
}
