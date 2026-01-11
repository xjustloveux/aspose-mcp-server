using AsposeMcpServer.Handlers.Word.HeaderFooter;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Word.HeaderFooter;

public class SetHeaderTabsHandlerTests : WordHandlerTestBase
{
    private readonly SetHeaderTabsHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_SetHeaderTabs()
    {
        Assert.Equal("set_header_tabs", _handler.Operation);
    }

    #endregion

    #region Basic Operations

    [Fact]
    public void Execute_SetsHeaderTabs()
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
    public void Execute_WithOnlyCenterTab_SetsCenterTab()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "centerTabPosition", 150.0 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("tab", result.ToLower());
    }

    #endregion
}
