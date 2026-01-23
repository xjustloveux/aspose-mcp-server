using Aspose.Words;
using Aspose.Words.Drawing;
using AsposeMcpServer.Handlers.Word.Shape;
using AsposeMcpServer.Results.Word.Shape;
using AsposeMcpServer.Tests.Infrastructure;
using WordParagraph = Aspose.Words.Paragraph;

namespace AsposeMcpServer.Tests.Handlers.Word.Shape;

public class GetTextboxesWordHandlerTests : WordHandlerTestBase
{
    private readonly GetTextboxesWordHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_GetTextboxes()
    {
        Assert.Equal("get_textboxes", _handler.Operation);
    }

    #endregion

    #region Helper Methods

    private static Document CreateDocumentWithTextbox()
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        var textBox = new Aspose.Words.Drawing.Shape(doc, ShapeType.TextBox)
        {
            Width = 200,
            Height = 100
        };

        var para = new WordParagraph(doc);
        var run = new Run(doc, "Textbox content");
        para.AppendChild(run);
        textBox.AppendChild(para);

        builder.InsertNode(textBox);
        return doc;
    }

    #endregion

    #region Basic Get Textboxes Operations

    [Fact]
    public void Execute_WithNoTextboxes_ReturnsNoTextboxesMessage()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetTextboxesWordResult>(res);

        Assert.Contains("no textboxes found", result.Content.ToLower());
    }

    [Fact]
    public void Execute_WithTextbox_ReturnsTextboxInfo()
    {
        var doc = CreateDocumentWithTextbox();
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetTextboxesWordResult>(res);

        Assert.Contains("total textboxes:", result.Content.ToLower());
        Assert.Contains("width:", result.Content.ToLower());
        Assert.Contains("height:", result.Content.ToLower());
    }

    [Fact]
    public void Execute_WithIncludeContent_ReturnsTextboxContent()
    {
        var doc = CreateDocumentWithTextbox();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "includeContent", true }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetTextboxesWordResult>(res);

        Assert.Contains("content", result.Content.ToLower());
    }

    #endregion
}
