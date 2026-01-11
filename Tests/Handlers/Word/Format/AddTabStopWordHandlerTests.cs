using AsposeMcpServer.Handlers.Word.Format;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Word.Format;

public class AddTabStopWordHandlerTests : WordHandlerTestBase
{
    private readonly AddTabStopWordHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_AddTabStop()
    {
        Assert.Equal("add_tab_stop", _handler.Operation);
    }

    #endregion

    #region Basic Add Operations

    [Fact]
    public void Execute_AddsTabStop()
    {
        var doc = CreateDocumentWithText("Sample text with tab.");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "paragraphIndex", 0 },
            { "tabPosition", 72.0 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("tab stop added", result.ToLower());
        Assert.Contains("72", result);
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithAlignment_AddsTabStopWithAlignment()
    {
        var doc = CreateDocumentWithText("Sample text.");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "paragraphIndex", 0 },
            { "tabPosition", 144.0 },
            { "tabAlignment", "center" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("center", result.ToLower());
    }

    [Fact]
    public void Execute_WithLeader_AddsTabStopWithLeader()
    {
        var doc = CreateDocumentWithText("Sample text.");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "paragraphIndex", 0 },
            { "tabPosition", 200.0 },
            { "tabLeader", "dots" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("dots", result.ToLower());
    }

    #endregion
}
