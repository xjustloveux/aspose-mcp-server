using AsposeMcpServer.Handlers.Word.HeaderFooter;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Word.HeaderFooter;

public class SetFooterLineHandlerTests : WordHandlerTestBase
{
    private readonly SetFooterLineHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_SetFooterLine()
    {
        Assert.Equal("set_footer_line", _handler.Operation);
    }

    #endregion

    #region Basic Operations

    [Fact]
    public void Execute_SetsFooterLine()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "showLine", true }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("line", result.Message, StringComparison.OrdinalIgnoreCase);
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithLineWidth_SetsWidth()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "showLine", true },
            { "lineWidth", 1.5 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("line", result.Message, StringComparison.OrdinalIgnoreCase);
    }

    #endregion
}
