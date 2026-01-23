using Aspose.Cells;
using AsposeMcpServer.Handlers.Excel.NamedRange;
using AsposeMcpServer.Results.Excel.NamedRange;
using AsposeMcpServer.Tests.Infrastructure;

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

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetNamedRangesResult>(res);

        Assert.Equal(0, result.Count);
        Assert.Equal("No named ranges found", result.Message);
        AssertNotModified(context);
    }

    [Fact]
    public void Execute_WithNamedRanges_ReturnsRangeInfo()
    {
        var workbook = CreateWorkbookWithNamedRange();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetNamedRangesResult>(res);

        Assert.True(result.Count > 0);
        Assert.Contains(result.Items, item => item.Name == "TestRange");
    }

    [Fact]
    public void Execute_ReturnsItems()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetNamedRangesResult>(res);

        Assert.NotNull(result.Items);
    }

    #endregion
}
