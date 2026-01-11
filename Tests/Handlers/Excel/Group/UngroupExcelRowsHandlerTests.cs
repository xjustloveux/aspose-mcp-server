using Aspose.Cells;
using AsposeMcpServer.Handlers.Excel.Group;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Excel.Group;

public class UngroupExcelRowsHandlerTests : ExcelHandlerTestBase
{
    private readonly UngroupExcelRowsHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_UngroupRows()
    {
        Assert.Equal("ungroup_rows", _handler.Operation);
    }

    #endregion

    #region Basic Ungroup Operations

    [Fact]
    public void Execute_UngroupsRows()
    {
        var workbook = CreateWorkbookWithGroupedRows();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "startRow", 0 },
            { "endRow", 5 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("ungrouped", result.ToLower());
        AssertModified(context);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutStartRow_ThrowsArgumentException()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "endRow", 5 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("startRow", ex.Message);
    }

    #endregion

    #region Helper Methods

    private static Workbook CreateWorkbookWithGroupedRows()
    {
        var workbook = new Workbook();
        workbook.Worksheets[0].Cells.GroupRows(0, 5);
        return workbook;
    }

    #endregion
}
