using AsposeMcpServer.Handlers.Word.HeaderFooter;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Word.HeaderFooter;

public class SetFooterTabsHandlerTests : WordHandlerTestBase
{
    private readonly SetFooterTabsHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_SetFooterTabs()
    {
        Assert.Equal("set_footer_tabs", _handler.Operation);
    }

    #endregion

    #region Basic Operations

    [Fact]
    public void Execute_SetsFooterTabs()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "centerTabPosition", 200.0 },
            { "rightTabPosition", 400.0 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("tab", result.ToLower());
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithOnlyRightTab_SetsRightTab()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "rightTabPosition", 350.0 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("tab", result.ToLower());
    }

    #endregion
}
