using Aspose.Cells;
using AsposeMcpServer.Handlers.Excel.Group;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Excel.Group;

public class UngroupExcelColumnsHandlerTests : ExcelHandlerTestBase
{
    private readonly UngroupExcelColumnsHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_UngroupColumns()
    {
        Assert.Equal("ungroup_columns", _handler.Operation);
    }

    #endregion

    #region Basic Ungroup Operations

    [Fact]
    public void Execute_UngroupsColumns()
    {
        var workbook = CreateWorkbookWithGroupedColumns();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "startColumn", 0 },
            { "endColumn", 3 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("ungrouped", result.Message, StringComparison.OrdinalIgnoreCase);
        AssertModified(context);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutStartColumn_ThrowsArgumentException()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "endColumn", 3 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("startColumn", ex.Message);
    }

    #endregion

    #region Helper Methods

    private static Workbook CreateWorkbookWithGroupedColumns()
    {
        var workbook = new Workbook();
        workbook.Worksheets[0].Cells.GroupColumns(0, 3);
        return workbook;
    }

    #endregion
}
