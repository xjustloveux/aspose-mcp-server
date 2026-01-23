using Aspose.Cells;
using AsposeMcpServer.Handlers.Excel.FreezePanes;
using AsposeMcpServer.Results.Excel.FreezePanes;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Excel.FreezePanes;

public class GetExcelFreezePanesHandlerTests : ExcelHandlerTestBase
{
    private readonly GetExcelFreezePanesHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Get()
    {
        Assert.Equal("get", _handler.Operation);
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

    #region Basic Get Operations

    [Fact]
    public void Execute_WithNoFreezePanes_ReturnsNotFrozen()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetFreezePanesResult>(res);

        Assert.False(result.IsFrozen);
        Assert.Contains("not frozen", result.Status.ToLower());
        AssertNotModified(context);
    }

    [Fact]
    public void Execute_WithFreezePanes_ReturnsFrozenStatus()
    {
        var workbook = CreateWorkbookWithFrozenPanes();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetFreezePanesResult>(res);

        Assert.True(result.IsFrozen);
        Assert.Contains("frozen", result.Status.ToLower());
    }

    [Fact]
    public void Execute_ReturnsWorksheetName()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetFreezePanesResult>(res);

        Assert.NotNull(result.WorksheetName);
    }

    #endregion
}
