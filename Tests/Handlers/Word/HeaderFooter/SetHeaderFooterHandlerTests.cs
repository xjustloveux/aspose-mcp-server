using Aspose.Words;
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

        Assert.IsType<SuccessResult>(res);

        var header = doc.FirstSection.HeadersFooters[HeaderFooterType.HeaderPrimary];
        Assert.NotNull(header);
        if (!IsEvaluationMode(AsposeLibraryType.Words)) Assert.Contains("Text", header.GetText());
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

        Assert.IsType<SuccessResult>(res);

        var header = doc.FirstSection.HeadersFooters[HeaderFooterType.HeaderPrimary];
        Assert.NotNull(header);
        if (!IsEvaluationMode(AsposeLibraryType.Words)) Assert.Contains("All Sections", header.GetText());
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

        Assert.IsType<SuccessResult>(res);
        AssertModified(context);

        var header = doc.FirstSection.HeadersFooters[HeaderFooterType.HeaderPrimary];
        Assert.NotNull(header);
        var footer = doc.FirstSection.HeadersFooters[HeaderFooterType.FooterPrimary];
        Assert.NotNull(footer);
        if (!IsEvaluationMode(AsposeLibraryType.Words))
        {
            Assert.Contains("Header Text", header.GetText());
            Assert.Contains("Footer Text", footer.GetText());
        }
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

        Assert.IsType<SuccessResult>(res);
        AssertModified(context);

        var header = doc.FirstSection.HeadersFooters[HeaderFooterType.HeaderPrimary];
        Assert.NotNull(header);
        if (!IsEvaluationMode(AsposeLibraryType.Words)) Assert.Contains("Header Only", header.GetText());
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

        Assert.IsType<SuccessResult>(res);
        AssertModified(context);

        var footer = doc.FirstSection.HeadersFooters[HeaderFooterType.FooterPrimary];
        Assert.NotNull(footer);
        if (!IsEvaluationMode(AsposeLibraryType.Words)) Assert.Contains("Footer Only", footer.GetText());
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

        Assert.IsType<SuccessResult>(res);
        AssertModified(context);

        var header = doc.FirstSection.HeadersFooters[HeaderFooterType.HeaderPrimary];
        Assert.NotNull(header);
        var footer = doc.FirstSection.HeadersFooters[HeaderFooterType.FooterPrimary];
        Assert.NotNull(footer);
        if (!IsEvaluationMode(AsposeLibraryType.Words))
        {
            var headerText = header.GetText();
            Assert.Contains("H-Left", headerText);
            Assert.Contains("H-Center", headerText);
            Assert.Contains("H-Right", headerText);
            var footerText = footer.GetText();
            Assert.Contains("F-Left", footerText);
            Assert.Contains("F-Center", footerText);
            Assert.Contains("F-Right", footerText);
        }
    }

    #endregion
}
