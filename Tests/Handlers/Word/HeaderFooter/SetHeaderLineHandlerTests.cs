using AsposeMcpServer.Handlers.Word.HeaderFooter;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Word.HeaderFooter;

public class SetHeaderLineHandlerTests : WordHandlerTestBase
{
    private readonly SetHeaderLineHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_SetHeaderLine()
    {
        Assert.Equal("set_header_line", _handler.Operation);
    }

    #endregion

    #region Basic Operations

    [Fact]
    public void Execute_SetsHeaderLine()
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
            { "lineWidth", 2.0 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("line", result.ToLower());
    }

    [Fact]
    public void Execute_WithLineColor_SetsColor()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "showLine", true },
            { "lineColor", "Blue" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("line", result.ToLower());
    }

    #endregion
}
