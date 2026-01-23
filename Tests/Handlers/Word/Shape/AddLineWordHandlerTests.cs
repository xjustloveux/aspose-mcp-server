using AsposeMcpServer.Handlers.Word.Shape;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Word.Shape;

public class AddLineWordHandlerTests : WordHandlerTestBase
{
    private readonly AddLineWordHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_AddLine()
    {
        Assert.Equal("add_line", _handler.Operation);
    }

    #endregion

    #region Basic Add Line Operations

    [Fact]
    public void Execute_AddsLineToBody()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("successfully inserted line", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("document body", result.Message, StringComparison.OrdinalIgnoreCase);
        AssertModified(context);
    }

    [Fact]
    public void Execute_AddsLineToHeader()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "location", "header" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("successfully inserted line", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("header", result.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Execute_AddsLineToFooter()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "location", "footer" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("successfully inserted line", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("footer", result.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Execute_AddsLineAtStart()
    {
        var doc = CreateDocumentWithText("Some content.");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "position", "start" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("start position", result.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Execute_WithCustomLineStyle()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "lineStyle", "border" },
            { "lineWidth", 2.0 },
            { "lineColor", "FF0000" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("successfully inserted line", result.Message, StringComparison.OrdinalIgnoreCase);
    }

    #endregion
}
