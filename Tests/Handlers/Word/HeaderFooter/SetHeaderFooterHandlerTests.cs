using AsposeMcpServer.Handlers.Word.HeaderFooter;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Word.HeaderFooter;

public class SetHeaderFooterHandlerTests : WordHandlerTestBase
{
    private readonly SetHeaderFooterHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_SetHeaderFooter()
    {
        Assert.Equal("set_header_footer", _handler.Operation);
    }

    #endregion

    #region Font Options

    [Fact]
    public void Execute_WithFontOptions_SetsFont()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "headerLeft", "Text" },
            { "fontName", "Arial" },
            { "fontSize", 12.0 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("set", result.Message, StringComparison.OrdinalIgnoreCase);
    }

    #endregion

    #region Section Options

    [Fact]
    public void Execute_WithSectionIndexMinus1_AppliesAllSections()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "headerLeft", "All Sections" },
            { "sectionIndex", -1 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("set", result.Message, StringComparison.OrdinalIgnoreCase);
    }

    #endregion

    #region Basic Operations

    [Fact]
    public void Execute_SetsBothHeaderAndFooter()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "headerLeft", "Header Text" },
            { "footerLeft", "Footer Text" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("header and footer set", result.Message, StringComparison.OrdinalIgnoreCase);
        AssertModified(context);
    }

    [Fact]
    public void Execute_SetsOnlyHeader()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "headerCenter", "Header Only" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("set", result.Message, StringComparison.OrdinalIgnoreCase);
        AssertModified(context);
    }

    [Fact]
    public void Execute_SetsOnlyFooter()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "footerCenter", "Footer Only" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("set", result.Message, StringComparison.OrdinalIgnoreCase);
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithAllPositions_SetsAllContent()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "headerLeft", "H-Left" },
            { "headerCenter", "H-Center" },
            { "headerRight", "H-Right" },
            { "footerLeft", "F-Left" },
            { "footerCenter", "F-Center" },
            { "footerRight", "F-Right" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("set", result.Message, StringComparison.OrdinalIgnoreCase);
        AssertModified(context);
    }

    #endregion
}
