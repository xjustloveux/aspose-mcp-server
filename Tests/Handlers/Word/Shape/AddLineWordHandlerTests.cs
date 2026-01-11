using AsposeMcpServer.Handlers.Word.Shape;
using AsposeMcpServer.Tests.Helpers;

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

        var result = _handler.Execute(context, parameters);

        Assert.Contains("successfully inserted line", result.ToLower());
        Assert.Contains("document body", result.ToLower());
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

        var result = _handler.Execute(context, parameters);

        Assert.Contains("successfully inserted line", result.ToLower());
        Assert.Contains("header", result.ToLower());
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

        var result = _handler.Execute(context, parameters);

        Assert.Contains("successfully inserted line", result.ToLower());
        Assert.Contains("footer", result.ToLower());
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

        var result = _handler.Execute(context, parameters);

        Assert.Contains("start position", result.ToLower());
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

        var result = _handler.Execute(context, parameters);

        Assert.Contains("successfully inserted line", result.ToLower());
    }

    #endregion
}
