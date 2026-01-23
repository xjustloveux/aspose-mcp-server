using Aspose.Cells;
using AsposeMcpServer.Handlers.Excel.Style;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Excel.Style;

public class CopySheetFormatHandlerTests : ExcelHandlerTestBase
{
    private readonly CopySheetFormatHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_CopySheetFormat()
    {
        Assert.Equal("copy_sheet_format", _handler.Operation);
    }

    #endregion

    #region Helper Methods

    private static Workbook CreateWorkbookWithTwoSheets()
    {
        var workbook = new Workbook();
        workbook.Worksheets.Add("Sheet2");
        return workbook;
    }

    #endregion

    #region Basic Copy Operations

    [Fact]
    public void Execute_CopiesSheetFormat()
    {
        var workbook = CreateWorkbookWithTwoSheets();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sourceSheetIndex", 0 },
            { "targetSheetIndex", 1 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("copied", result.Message, StringComparison.OrdinalIgnoreCase);
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithCopyColumnWidths_CopiesColumnWidths()
    {
        var workbook = CreateWorkbookWithTwoSheets();
        workbook.Worksheets[0].Cells.SetColumnWidth(0, 20);
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sourceSheetIndex", 0 },
            { "targetSheetIndex", 1 },
            { "copyColumnWidths", true }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("copied", result.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Execute_WithCopyRowHeights_CopiesRowHeights()
    {
        var workbook = CreateWorkbookWithTwoSheets();
        workbook.Worksheets[0].Cells.SetRowHeight(0, 30);
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sourceSheetIndex", 0 },
            { "targetSheetIndex", 1 },
            { "copyRowHeights", true }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("copied", result.Message, StringComparison.OrdinalIgnoreCase);
    }

    #endregion
}
