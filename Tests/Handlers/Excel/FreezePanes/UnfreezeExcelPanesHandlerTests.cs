using Aspose.Cells;
using AsposeMcpServer.Handlers.Excel.FreezePanes;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Excel.FreezePanes;

public class UnfreezeExcelPanesHandlerTests : ExcelHandlerTestBase
{
    private readonly UnfreezeExcelPanesHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Unfreeze()
    {
        Assert.Equal("unfreeze", _handler.Operation);
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

    #region Basic Unfreeze Operations

    [Fact]
    public void Execute_UnfreezesPanes()
    {
        var workbook = CreateWorkbookWithFrozenPanes();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("unfrozen", result.ToLower());
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithSheetIndex_UnfreezesOnSpecificSheet()
    {
        var workbook = CreateWorkbookWithFrozenPanes();
        workbook.Worksheets.Add("Sheet2");
        workbook.Worksheets[1].FreezePanes(2, 2, 1, 1);
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sheetIndex", 1 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("unfrozen", result.ToLower());
    }

    #endregion
}
