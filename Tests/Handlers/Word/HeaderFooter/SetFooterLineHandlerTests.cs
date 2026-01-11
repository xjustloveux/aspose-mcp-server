using AsposeMcpServer.Handlers.Word.HeaderFooter;
using AsposeMcpServer.Tests.Helpers;

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

        var result = _handler.Execute(context, parameters);

        Assert.Contains("line", result.ToLower());
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

        var result = _handler.Execute(context, parameters);

        Assert.Contains("line", result.ToLower());
    }

    #endregion
}
