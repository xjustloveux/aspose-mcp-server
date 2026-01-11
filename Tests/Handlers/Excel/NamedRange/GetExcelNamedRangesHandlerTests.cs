using Aspose.Cells;
using AsposeMcpServer.Handlers.Excel.NamedRange;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Excel.NamedRange;

public class GetExcelNamedRangesHandlerTests : ExcelHandlerTestBase
{
    private readonly GetExcelNamedRangesHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Get()
    {
        Assert.Equal("get", _handler.Operation);
    }

    #endregion

    #region Helper Methods

    private static Workbook CreateWorkbookWithNamedRange()
    {
        var workbook = new Workbook();
        workbook.Worksheets[0].Cells.CreateRange("A1:B5").Name = "TestRange";
        return workbook;
    }

    #endregion

    #region Basic Get Operations

    [Fact]
    public void Execute_WithNoNamedRanges_ReturnsEmptyList()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("\"count\": 0", result);
        Assert.Contains("No named ranges found", result);
        AssertNotModified(context);
    }

    [Fact]
    public void Execute_WithNamedRanges_ReturnsRangeInfo()
    {
        var workbook = CreateWorkbookWithNamedRange();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("TestRange", result);
        Assert.Contains("count", result.ToLower());
    }

    [Fact]
    public void Execute_ReturnsJsonFormat()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("{", result);
        Assert.Contains("}", result);
    }

    #endregion
}
