using AsposeMcpServer.Handlers.PowerPoint.Notes;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.Notes;

public class SetNotesHeaderFooterHandlerTests : PptHandlerTestBase
{
    private readonly SetNotesHeaderFooterHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_SetHeaderFooter()
    {
        Assert.Equal("set_header_footer", _handler.Operation);
    }

    #endregion

    #region Basic Set Notes Header Footer Operations

    [Fact]
    public void Execute_SetsHeaderText()
    {
        var presentation = CreateEmptyPresentation();
        var context = CreateContext(presentation);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "headerText", "Test Header" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("header", result.ToLower());
        AssertModified(context);
    }

    [Fact]
    public void Execute_SetsFooterText()
    {
        var presentation = CreateEmptyPresentation();
        var context = CreateContext(presentation);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "footerText", "Test Footer" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("footer", result.ToLower());
    }

    [Fact]
    public void Execute_SetsDateText()
    {
        var presentation = CreateEmptyPresentation();
        var context = CreateContext(presentation);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "dateText", "2026-01-11" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("date", result.ToLower());
    }

    [Fact]
    public void Execute_SetsPageNumberVisibility()
    {
        var presentation = CreateEmptyPresentation();
        var context = CreateContext(presentation);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "showPageNumber", false }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("page number hidden", result.ToLower());
    }

    [Fact]
    public void Execute_WithAllSettings()
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

        Assert.Contains("header", result.ToLower());
        Assert.Contains("footer", result.ToLower());
        Assert.Contains("date", result.ToLower());
    }

    #endregion
}
