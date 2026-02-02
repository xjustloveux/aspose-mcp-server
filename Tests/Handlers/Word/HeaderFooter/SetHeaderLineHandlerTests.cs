using Aspose.Words;
using AsposeMcpServer.Handlers.Word.HeaderFooter;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;
using WordParagraph = Aspose.Words.Paragraph;

namespace AsposeMcpServer.Tests.Handlers.Word.HeaderFooter;

public class SetHeaderLineHandlerTests : WordHandlerTestBase
{
    private readonly SetHeaderLineHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_SetHeaderLine()
    {
        Assert.Equal("set_header_line", _handler.Operation);
    }

    #endregion

    #region Basic Operations

    [Fact]
    public void Execute_SetsHeaderLine()
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
            var header = doc.FirstSection.HeadersFooters[HeaderFooterType.HeaderPrimary];
            Assert.NotNull(header);
            var para = header.GetChildNodes(NodeType.Paragraph, true).Cast<WordParagraph>().LastOrDefault();
            Assert.NotNull(para);
            Assert.Equal(LineStyle.Single, para.ParagraphFormat.Borders.Bottom.LineStyle);
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
            { "lineWidth", 2.0 }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        if (!IsEvaluationMode(AsposeLibraryType.Words))
        {
            var header = doc.FirstSection.HeadersFooters[HeaderFooterType.HeaderPrimary];
            Assert.NotNull(header);
            var para = header.GetChildNodes(NodeType.Paragraph, true).Cast<WordParagraph>().LastOrDefault();
            Assert.NotNull(para);
            Assert.Equal(2.0, para.ParagraphFormat.Borders.Bottom.LineWidth);
        }
    }

    [Fact]
    public void Execute_WithLineColor_SetsColor()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "showLine", true },
            { "lineColor", "Blue" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        if (!IsEvaluationMode(AsposeLibraryType.Words))
        {
            var header = doc.FirstSection.HeadersFooters[HeaderFooterType.HeaderPrimary];
            Assert.NotNull(header);
            var para = header.GetChildNodes(NodeType.Paragraph, true).Cast<WordParagraph>().LastOrDefault();
            Assert.NotNull(para);
            Assert.Equal(LineStyle.Single, para.ParagraphFormat.Borders.Bottom.LineStyle);
        }
    }

    #endregion
}
