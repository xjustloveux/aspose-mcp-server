using Aspose.Cells;
using AsposeMcpServer.Handlers.Excel.FreezePanes;
using AsposeMcpServer.Tests.Helpers;

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

        var result = _handler.Execute(context, parameters);

        Assert.Contains("not frozen", result.ToLower());
        AssertNotModified(context);
    }

    [Fact]
    public void Execute_WithFreezePanes_ReturnsFrozenStatus()
    {
        var workbook = CreateWorkbookWithFrozenPanes();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("frozen", result.ToLower());
        Assert.Contains("isFrozen", result);
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
        Assert.Contains("worksheetName", result);
    }

    #endregion
}
