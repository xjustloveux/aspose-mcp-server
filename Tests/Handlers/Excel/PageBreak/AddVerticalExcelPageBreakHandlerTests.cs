using AsposeMcpServer.Handlers.Excel.PageBreak;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Excel.PageBreak;

public class AddVerticalExcelPageBreakHandlerTests : ExcelHandlerTestBase
{
    private readonly AddVerticalExcelPageBreakHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_ShouldBeAddVertical()
    {
        Assert.Equal("add_vertical", _handler.Operation);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithMissingColumn_ShouldThrow()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("column", ex.Message);
    }

    #endregion

    #region Basic Add Operations

    [Fact]
    public void Execute_ShouldAddVerticalPageBreak()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "column", 5 }
        });

        _handler.Execute(context, parameters);

        Assert.Single(workbook.Worksheets[0].VerticalPageBreaks);
    }

    [Fact]
    public void Execute_ShouldMarkModified()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "column", 5 }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        AssertModified(context);
    }

    #endregion
}
