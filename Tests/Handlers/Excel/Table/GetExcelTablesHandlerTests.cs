using Aspose.Cells;
using AsposeMcpServer.Handlers.Excel.Table;
using AsposeMcpServer.Results.Excel.Table;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Excel.Table;

public class GetExcelTablesHandlerTests : ExcelHandlerTestBase
{
    private readonly GetExcelTablesHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_ShouldBeGet()
    {
        Assert.Equal("get", _handler.Operation);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithInvalidTableIndex_ShouldThrow()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "tableIndex", 99 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion

    #region Helper Methods

    private static Workbook CreateWorkbookWithTable()
    {
        var workbook = CreateWorkbookWithData(new object[,]
        {
            { "Name", "Value", "Count" },
            { "A", 1, 10 },
            { "B", 2, 20 },
            { "C", 3, 30 }
        });
        workbook.Worksheets[0].ListObjects.Add("A1", "C4", true);
        return workbook;
    }

    #endregion

    #region Basic Get Operations

    [Fact]
    public void Execute_WithNoTables_ShouldReturnEmpty()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetTablesExcelResult>(res);
        Assert.Equal(0, result.Count);
        Assert.Empty(result.Items);
    }

    [Fact]
    public void Execute_WithTable_ShouldReturnTableInfo()
    {
        var workbook = CreateWorkbookWithTable();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetTablesExcelResult>(res);
        Assert.Equal(1, result.Count);
        Assert.Single(result.Items);
        Assert.Equal(0, result.Items[0].Index);
        Assert.Equal(3, result.Items[0].ColumnCount);
    }

    [Fact]
    public void Execute_WithSpecificTableIndex_ShouldReturnSingleTable()
    {
        var workbook = CreateWorkbookWithTable();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "tableIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetTablesExcelResult>(res);
        Assert.Equal(1, result.Count);
        Assert.Single(result.Items);
    }

    #endregion
}
