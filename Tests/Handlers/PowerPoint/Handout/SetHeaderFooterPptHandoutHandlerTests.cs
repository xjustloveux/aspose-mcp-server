using AsposeMcpServer.Handlers.PowerPoint.Handout;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

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

    #region Auto-Create Handout Master

    [Fact]
    public void Execute_WithNoHandoutMaster_AutoCreatesAndSetsHeader()
    {
        var presentation = CreateEmptyPresentation();
        var context = CreateContext(presentation);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "headerText", "Test Header" }
        });

        var result = _handler.Execute(context, parameters);

        var success = Assert.IsType<SuccessResult>(result);
        Assert.Contains("header", success.Message);
    }

    [Fact]
    public void Execute_WithNoHandoutMaster_AutoCreatesAndSetsFooter()
    {
        var presentation = CreateEmptyPresentation();
        var context = CreateContext(presentation);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "footerText", "Test Footer" }
        });

        var result = _handler.Execute(context, parameters);

        var success = Assert.IsType<SuccessResult>(result);
        Assert.Contains("footer", success.Message);
    }

    [Fact]
    public void Execute_WithNoHandoutMaster_AutoCreatesAndSetsDate()
    {
        var presentation = CreateEmptyPresentation();
        var context = CreateContext(presentation);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "dateText", "2026-01-11" }
        });

        var result = _handler.Execute(context, parameters);

        var success = Assert.IsType<SuccessResult>(result);
        Assert.Contains("date", success.Message);
    }

    [Fact]
    public void Execute_WithNoHandoutMaster_AutoCreatesAndSetsAllSettings()
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

        var result = _handler.Execute(context, parameters);

        var success = Assert.IsType<SuccessResult>(result);
        Assert.Contains("header", success.Message);
        Assert.Contains("footer", success.Message);
        Assert.Contains("date", success.Message);
        Assert.Contains("page number shown", success.Message);
    }

    #endregion
}
