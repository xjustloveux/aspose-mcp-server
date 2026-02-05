using AsposeMcpServer.Handlers.Excel.Sparkline;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Excel.Sparkline;

public class DeleteExcelSparklineHandlerTests : ExcelHandlerTestBase
{
    private readonly DeleteExcelSparklineHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_ShouldBeDelete()
    {
        Assert.Equal("delete", _handler.Operation);
    }

    #endregion

    #region Execute Tests

    [Fact]
    public void Execute_WithMissingGroupIndex_ShouldThrow()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("groupIndex", ex.Message);
    }

    [Fact]
    public void Execute_WithInvalidGroupIndex_ShouldThrow()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "groupIndex", 0 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("out of range", ex.Message);
    }

    #endregion
}
