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

    [Fact]
    public void Execute_SetsFooterLine()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "showLine", true }
        });

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
            { "showLine", true },
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
