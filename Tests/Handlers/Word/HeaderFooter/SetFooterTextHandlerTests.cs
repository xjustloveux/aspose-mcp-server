using Aspose.Words;
using AsposeMcpServer.Handlers.Word.HeaderFooter;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

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

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("warning", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("no footer text", result.Message, StringComparison.OrdinalIgnoreCase);
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

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        AssertModified(context);

        var footer = doc.FirstSection.HeadersFooters[HeaderFooterType.FooterPrimary];
        Assert.NotNull(footer);
        if (!IsEvaluationMode(AsposeLibraryType.Words)) Assert.Contains("Left Text", footer.GetText());
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

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        AssertModified(context);

        var footer = doc.FirstSection.HeadersFooters[HeaderFooterType.FooterPrimary];
        Assert.NotNull(footer);
        if (!IsEvaluationMode(AsposeLibraryType.Words)) Assert.Contains("Center Text", footer.GetText());
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

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        AssertModified(context);

        var footer = doc.FirstSection.HeadersFooters[HeaderFooterType.FooterPrimary];
        Assert.NotNull(footer);
        if (!IsEvaluationMode(AsposeLibraryType.Words)) Assert.Contains("Right Text", footer.GetText());
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

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        var footer = doc.FirstSection.HeadersFooters[HeaderFooterType.FooterPrimary];
        Assert.NotNull(footer);
        if (!IsEvaluationMode(AsposeLibraryType.Words))
        {
            var footerText = footer.GetText();
            Assert.Contains("Left", footerText);
            Assert.Contains("Center", footerText);
            Assert.Contains("Right", footerText);
        }
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

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        var footer = doc.FirstSection.HeadersFooters[HeaderFooterType.FooterPrimary];
        Assert.NotNull(footer);
        if (!IsEvaluationMode(AsposeLibraryType.Words)) Assert.Contains("Test", footer.GetText());
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

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        var footer = doc.FirstSection.HeadersFooters[HeaderFooterType.FooterPrimary];
        Assert.NotNull(footer);
        if (!IsEvaluationMode(AsposeLibraryType.Words)) Assert.Contains("Test", footer.GetText());
    }

    #endregion
}
