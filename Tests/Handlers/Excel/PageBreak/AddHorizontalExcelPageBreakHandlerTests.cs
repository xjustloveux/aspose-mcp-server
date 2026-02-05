using AsposeMcpServer.Handlers.Excel.PageBreak;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Excel.PageBreak;

public class AddHorizontalExcelPageBreakHandlerTests : ExcelHandlerTestBase
{
    private readonly AddHorizontalExcelPageBreakHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_ShouldBeAddHorizontal()
    {
        Assert.Equal("add_horizontal", _handler.Operation);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithMissingRow_ShouldThrow()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("row", ex.Message);
    }

    #endregion

    #region Basic Add Operations

    [Fact]
    public void Execute_ShouldAddHorizontalPageBreak()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "row", 10 }
        });

        _handler.Execute(context, parameters);

        Assert.Single(workbook.Worksheets[0].HorizontalPageBreaks);
    }

    [Fact]
    public void Execute_ShouldMarkModified()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "row", 10 }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        AssertModified(context);
    }

    #endregion
}
