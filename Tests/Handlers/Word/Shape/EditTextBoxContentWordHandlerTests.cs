using Aspose.Words;
using Aspose.Words.Drawing;
using AsposeMcpServer.Handlers.Word.Shape;
using AsposeMcpServer.Helpers.Word;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;
using WordParagraph = Aspose.Words.Paragraph;

namespace AsposeMcpServer.Tests.Handlers.Word.Shape;

public class EditTextBoxContentWordHandlerTests : WordHandlerTestBase
{
    private readonly EditTextBoxContentWordHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_EditTextboxContent()
    {
        Assert.Equal("edit_textbox_content", _handler.Operation);
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
        var run = new Run(doc, "Original content");
        para.AppendChild(run);
        textBox.AppendChild(para);

        builder.InsertNode(textBox);
        return doc;
    }

    #endregion

    #region Basic Edit TextBox Content Operations

    [Fact]
    public void Execute_UpdatesTextboxContent()
    {
        var doc = CreateDocumentWithTextbox();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "textboxIndex", 0 },
            { "text", "New Content" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        var textboxes = WordShapeHelper.FindAllTextboxes(doc);
        Assert.NotEmpty(textboxes);
        if (!IsEvaluationMode(AsposeLibraryType.Words))
        {
            var runs = textboxes[0].GetChildNodes(NodeType.Run, true).Cast<Run>().ToList();
            Assert.NotEmpty(runs);
            Assert.Equal("New Content", runs[0].Text);
        }

        AssertModified(context);
    }

    [Fact]
    public void Execute_AppendsTextToTextbox()
    {
        var doc = CreateDocumentWithTextbox();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "textboxIndex", 0 },
            { "text", " Appended" },
            { "appendText", true }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        var textboxes = WordShapeHelper.FindAllTextboxes(doc);
        Assert.NotEmpty(textboxes);
        if (!IsEvaluationMode(AsposeLibraryType.Words))
        {
            var fullText = textboxes[0].GetText();
            Assert.Contains("Original content", fullText);
            Assert.Contains(" Appended", fullText);
        }
    }

    [Fact]
    public void Execute_WithFormatting_AppliesFormatting()
    {
        var doc = CreateDocumentWithTextbox();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "textboxIndex", 0 },
            { "text", "Formatted" },
            { "fontSize", 16.0 },
            { "bold", true }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        var textboxes = WordShapeHelper.FindAllTextboxes(doc);
        Assert.NotEmpty(textboxes);
        if (!IsEvaluationMode(AsposeLibraryType.Words))
        {
            var runs = textboxes[0].GetChildNodes(NodeType.Run, true).Cast<Run>().ToList();
            Assert.NotEmpty(runs);
            Assert.Equal("Formatted", runs[0].Text);
            Assert.Equal(16.0, runs[0].Font.Size);
            Assert.True(runs[0].Font.Bold);
        }
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
