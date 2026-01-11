using Aspose.Words;
using Aspose.Words.Drawing;
using AsposeMcpServer.Handlers.Word.Shape;
using AsposeMcpServer.Tests.Helpers;
using WordParagraph = Aspose.Words.Paragraph;

namespace AsposeMcpServer.Tests.Handlers.Word.Shape;

public class SetTextBoxBorderWordHandlerTests : WordHandlerTestBase
{
    private readonly SetTextBoxBorderWordHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_SetTextboxBorder()
    {
        Assert.Equal("set_textbox_border", _handler.Operation);
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
        var run = new Run(doc, "Content");
        para.AppendChild(run);
        textBox.AppendChild(para);

        builder.InsertNode(textBox);
        return doc;
    }

    #endregion

    #region Basic Set TextBox Border Operations

    [Fact]
    public void Execute_SetsBorderColor()
    {
        var doc = CreateDocumentWithTextbox();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "textboxIndex", 0 },
            { "borderColor", "#FF0000" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("successfully", result.ToLower());
        AssertModified(context);
    }

    [Fact]
    public void Execute_SetsBorderWidth()
    {
        var doc = CreateDocumentWithTextbox();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "textboxIndex", 0 },
            { "borderWidth", 2.0 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("successfully", result.ToLower());
    }

    [Fact]
    public void Execute_WithoutTextboxIndex_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithTextbox();
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithInvalidIndex_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithTextbox();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "textboxIndex", 999 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion
}
