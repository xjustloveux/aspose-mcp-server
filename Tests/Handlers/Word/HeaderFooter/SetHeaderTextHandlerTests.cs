using AsposeMcpServer.Handlers.Word.HeaderFooter;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Word.HeaderFooter;

public class SetHeaderTextHandlerTests : WordHandlerTestBase
{
    private readonly SetHeaderTextHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_SetHeaderText()
    {
        Assert.Equal("set_header_text", _handler.Operation);
    }

    #endregion

    #region No Content Warning

    [Fact]
    public void Execute_WithNoContent_ReturnsWarning()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("warning", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("no header text", result.Message, StringComparison.OrdinalIgnoreCase);
    }

    #endregion

    #region Basic Set Operations

    [Fact]
    public void Execute_SetsHeaderText()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "headerLeft", "Left Text" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("header text set", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("left", result.Message, StringComparison.OrdinalIgnoreCase);
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithCenterText_SetsCenter()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "headerCenter", "Center Text" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("center", result.Message, StringComparison.OrdinalIgnoreCase);
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithRightText_SetsRight()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "headerRight", "Right Text" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("right", result.Message, StringComparison.OrdinalIgnoreCase);
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithAllPositions_SetsAll()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "headerLeft", "Left" },
            { "headerCenter", "Center" },
            { "headerRight", "Right" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("left", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("center", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("right", result.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Execute_WithFontName_SetsFont()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "headerLeft", "Text" },
            { "fontName", "Arial" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("header text set", result.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Execute_WithFontSize_SetsSize()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "headerLeft", "Text" },
            { "fontSize", 14.0 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("header text set", result.Message, StringComparison.OrdinalIgnoreCase);
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
            { "headerLeft", "Test" },
            { "sectionIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("section 0", result.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Execute_WithSectionIndexMinus1_SetsOnAllSections()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "headerLeft", "Test" },
            { "sectionIndex", -1 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("all sections", result.Message, StringComparison.OrdinalIgnoreCase);
    }

    #endregion
}
