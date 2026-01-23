using AsposeMcpServer.Handlers.PowerPoint.Notes;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

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

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("header", result.Message, StringComparison.OrdinalIgnoreCase);
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

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("footer", result.Message, StringComparison.OrdinalIgnoreCase);
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

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("date", result.Message, StringComparison.OrdinalIgnoreCase);
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

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("page number hidden", result.Message, StringComparison.OrdinalIgnoreCase);
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

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("header", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("footer", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("date", result.Message, StringComparison.OrdinalIgnoreCase);
    }

    #endregion
}
