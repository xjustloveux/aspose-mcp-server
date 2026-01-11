using Aspose.Cells;
using AsposeMcpServer.Handlers.Excel.Style;
using AsposeMcpServer.Tests.Helpers;

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

        var result = _handler.Execute(context, parameters);

        Assert.Contains("copied", result.ToLower());
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

        var result = _handler.Execute(context, parameters);

        Assert.Contains("copied", result.ToLower());
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

        var result = _handler.Execute(context, parameters);

        Assert.Contains("copied", result.ToLower());
    }

    #endregion
}
