using AsposeMcpServer.Handlers.Word.HeaderFooter;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Word.HeaderFooter;

public class SetFooterTextHandlerTests : WordHandlerTestBase
{
    private readonly SetFooterTextHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_SetFooterText()
    {
        Assert.Equal("set_footer_text", _handler.Operation);
    }

    #endregion

    #region No Content Warning

    [Fact]
    public void Execute_WithNoContent_ReturnsWarning()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("warning", result.ToLower());
        Assert.Contains("no footer text", result.ToLower());
    }

    #endregion

    #region Basic Set Operations

    [Fact]
    public void Execute_SetsFooterText()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "footerLeft", "Left Text" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("footer text set", result.ToLower());
        Assert.Contains("left", result.ToLower());
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithCenterText_SetsCenter()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "footerCenter", "Center Text" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("center", result.ToLower());
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithRightText_SetsRight()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "footerRight", "Right Text" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("right", result.ToLower());
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithAllPositions_SetsAll()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "footerLeft", "Left" },
            { "footerCenter", "Center" },
            { "footerRight", "Right" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("left", result.ToLower());
        Assert.Contains("center", result.ToLower());
        Assert.Contains("right", result.ToLower());
    }

    #endregion

    #region Section Index

    [Fact]
    public void Execute_WithSectionIndex_SetsOnSpecificSection()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "footerLeft", "Test" },
            { "sectionIndex", 0 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("section 0", result.ToLower());
    }

    [Fact]
    public void Execute_WithSectionIndexMinus1_SetsOnAllSections()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "footerLeft", "Test" },
            { "sectionIndex", -1 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("all sections", result.ToLower());
    }

    #endregion
}
