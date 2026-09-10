using Aspose.Words;
using AsposeMcpServer.Handlers.Word.HeaderFooter;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;
using WordParagraph = Aspose.Words.Paragraph;

namespace AsposeMcpServer.Tests.Handlers.Word.HeaderFooter;

public class SetFooterLineHandlerTests : WordHandlerTestBase
{
    private readonly SetFooterLineHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_SetFooterLine()
    {
        Assert.Equal("set_footer_line", _handler.Operation);
    }

    #endregion

    #region Basic Operations

    /// <summary>
    ///     The operation always appends a bordered paragraph; there is no toggle for it.
    ///     <para>
    ///         These cases used to pass a <c>showLine</c> key, but no handler or tool declares
    ///         one, so it was dropped and changed nothing. <c>lineStyle</c> only selects between
    ///         single, double and thick, with single as the fallback, so the line cannot be
    ///         switched off through this operation at all.
    ///     </para>
    /// </summary>
    [Fact]
    public void Execute_SetsFooterLine()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>());

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        if (!IsEvaluationMode(AsposeLibraryType.Words))
        {
            var footer = doc.FirstSection.HeadersFooters[HeaderFooterType.FooterPrimary];
            Assert.NotNull(footer);
            var para = footer.GetChildNodes(NodeType.Paragraph, true).Cast<WordParagraph>().LastOrDefault();
            Assert.NotNull(para);
            Assert.Equal(LineStyle.Single, para.ParagraphFormat.Borders.Top.LineStyle);
        }

        AssertModified(context);
    }

    [Fact]
    public void Execute_WithLineWidth_SetsWidth()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "lineWidth", 1.5 }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        if (!IsEvaluationMode(AsposeLibraryType.Words))
        {
            var footer = doc.FirstSection.HeadersFooters[HeaderFooterType.FooterPrimary];
            Assert.NotNull(footer);
            var para = footer.GetChildNodes(NodeType.Paragraph, true).Cast<WordParagraph>().LastOrDefault();
            Assert.NotNull(para);
            Assert.Equal(1.5, para.ParagraphFormat.Borders.Top.LineWidth);
        }
    }

    #endregion
}
