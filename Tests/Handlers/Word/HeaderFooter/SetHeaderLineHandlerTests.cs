using AsposeMcpServer.Handlers.Word.HeaderFooter;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

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
            { "lineWidth", 2.0 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("line", result.Message, StringComparison.OrdinalIgnoreCase);
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

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("line", result.Message, StringComparison.OrdinalIgnoreCase);
    }

    #endregion
}
