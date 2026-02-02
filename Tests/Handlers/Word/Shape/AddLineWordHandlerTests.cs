using Aspose.Words;
using Aspose.Words.Drawing;
using AsposeMcpServer.Handlers.Word.Shape;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Word.Shape;

public class AddLineWordHandlerTests : WordHandlerTestBase
{
    private readonly AddLineWordHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_AddLine()
    {
        Assert.Equal("add_line", _handler.Operation);
    }

    #endregion

    #region Helper Methods

    private static List<Aspose.Words.Drawing.Shape> GetLineShapes(Document doc)
    {
        return doc.GetChildNodes(NodeType.Shape, true)
            .Cast<Aspose.Words.Drawing.Shape>()
            .Where(s => s.ShapeType == ShapeType.Line)
            .ToList();
    }

    #endregion

    #region Basic Add Line Operations

    [Fact]
    public void Execute_AddsLineToBody()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        if (!IsEvaluationMode(AsposeLibraryType.Words))
        {
            var lines = GetLineShapes(doc);
            Assert.NotEmpty(lines);
        }

        AssertModified(context);
    }

    [Fact]
    public void Execute_AddsLineToHeader()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "location", "header" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        if (!IsEvaluationMode(AsposeLibraryType.Words))
        {
            var header = doc.FirstSection.HeadersFooters[HeaderFooterType.HeaderPrimary];
            Assert.NotNull(header);
            var headerShapes = header.GetChildNodes(NodeType.Shape, true)
                .Cast<Aspose.Words.Drawing.Shape>()
                .Where(s => s.ShapeType == ShapeType.Line)
                .ToList();
            Assert.NotEmpty(headerShapes);
        }
    }

    [Fact]
    public void Execute_AddsLineToFooter()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "location", "footer" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        if (!IsEvaluationMode(AsposeLibraryType.Words))
        {
            var footer = doc.FirstSection.HeadersFooters[HeaderFooterType.FooterPrimary];
            Assert.NotNull(footer);
            var footerShapes = footer.GetChildNodes(NodeType.Shape, true)
                .Cast<Aspose.Words.Drawing.Shape>()
                .Where(s => s.ShapeType == ShapeType.Line)
                .ToList();
            Assert.NotEmpty(footerShapes);
        }
    }

    [Fact]
    public void Execute_AddsLineAtStart()
    {
        var doc = CreateDocumentWithText("Some content.");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "position", "start" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        if (!IsEvaluationMode(AsposeLibraryType.Words))
        {
            var body = doc.FirstSection.Body;
            var firstPara = body.FirstParagraph;
            var lineShapes = firstPara.GetChildNodes(NodeType.Shape, true)
                .Cast<Aspose.Words.Drawing.Shape>()
                .Where(s => s.ShapeType == ShapeType.Line)
                .ToList();
            Assert.NotEmpty(lineShapes);
        }
    }

    [Fact]
    public void Execute_WithCustomLineStyle()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "lineStyle", "border" },
            { "lineWidth", 2.0 },
            { "lineColor", "FF0000" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        if (!IsEvaluationMode(AsposeLibraryType.Words))
        {
            var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true)
                .Cast<Aspose.Words.Paragraph>()
                .ToList();
            var hasBorderPara = paragraphs.Any(p =>
                p.ParagraphFormat.Borders.Bottom.LineStyle == LineStyle.Single);
            Assert.True(hasBorderPara);
        }
    }

    #endregion
}
