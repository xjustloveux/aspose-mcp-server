using AsposeMcpServer.Handlers.PowerPoint.Handout;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.Handout;

public class SetHeaderFooterPptHandoutHandlerTests : PptHandlerTestBase
{
    private readonly SetHeaderFooterPptHandoutHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_SetHeaderFooter()
    {
        Assert.Equal("set_header_footer", _handler.Operation);
    }

    #endregion

    #region Basic Set Header Footer Operations

    [Fact]
    public void Execute_WithNoHandoutMaster_ThrowsInvalidOperationException()
    {
        var presentation = CreateEmptyPresentation();
        var context = CreateContext(presentation);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "headerText", "Test Header" }
        });

        Assert.Throws<InvalidOperationException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithNoHandoutMaster_WithFooterText_ThrowsInvalidOperationException()
    {
        var presentation = CreateEmptyPresentation();
        var context = CreateContext(presentation);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "footerText", "Test Footer" }
        });

        Assert.Throws<InvalidOperationException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithNoHandoutMaster_WithDateText_ThrowsInvalidOperationException()
    {
        var presentation = CreateEmptyPresentation();
        var context = CreateContext(presentation);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "dateText", "2026-01-11" }
        });

        Assert.Throws<InvalidOperationException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithNoHandoutMaster_WithAllSettings_ThrowsInvalidOperationException()
    {
        var presentation = CreateEmptyPresentation();
        var context = CreateContext(presentation);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "headerText", "Header" },
            { "footerText", "Footer" },
            { "dateText", "Date" },
            { "showPageNumber", true }
        });

        Assert.Throws<InvalidOperationException>(() => _handler.Execute(context, parameters));
    }

    #endregion
}
