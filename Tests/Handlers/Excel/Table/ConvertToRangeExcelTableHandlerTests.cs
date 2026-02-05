using Aspose.Cells;
using AsposeMcpServer.Handlers.Excel.Table;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Excel.Table;

public class ConvertToRangeExcelTableHandlerTests : ExcelHandlerTestBase
{
    private readonly ConvertToRangeExcelTableHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_ShouldBeConvertToRange()
    {
        Assert.Equal("convert_to_range", _handler.Operation);
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

    #region Basic Convert Operations

    [Fact]
    public void Execute_ShouldConvertTableToRange()
    {
        var workbook = CreateWorkbookWithTable();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "tableIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);
        Assert.Contains("converted", result.Message);
        Assert.Empty(workbook.Worksheets[0].ListObjects);
    }

    [Fact]
    public void Execute_ShouldMarkModified()
    {
        var workbook = CreateWorkbookWithTable();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "tableIndex", 0 }
        });

        _handler.Execute(context, parameters);

        AssertModified(context);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithMissingTableIndex_ShouldThrow()
    {
        var workbook = CreateWorkbookWithTable();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("tableIndex", ex.Message);
    }

    [Fact]
    public void Execute_WithInvalidTableIndex_ShouldThrow()
    {
        var workbook = CreateWorkbookWithTable();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "tableIndex", 99 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion
}
