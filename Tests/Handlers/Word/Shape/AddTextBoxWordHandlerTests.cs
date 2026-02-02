using Aspose.Words;
using Aspose.Words.Drawing;
using AsposeMcpServer.Handlers.Word.Shape;
using AsposeMcpServer.Helpers.Word;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Word.Shape;

public class AddTextBoxWordHandlerTests : WordHandlerTestBase
{
    private readonly AddTextBoxWordHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_AddTextbox()
    {
        Assert.Equal("add_textbox", _handler.Operation);
    }

    #endregion

    #region Basic Add TextBox Operations

    [Fact]
    public void Execute_AddsTextBox()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "Hello World" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        var textboxes = WordShapeHelper.FindAllTextboxes(doc);
        Assert.NotEmpty(textboxes);
        if (!IsEvaluationMode(AsposeLibraryType.Words))
        {
            var tb = textboxes[0];
            Assert.Equal(ShapeType.TextBox, tb.ShapeType);
            Assert.Contains("Hello World", tb.GetText());
        }

        AssertModified(context);
    }

    [Fact]
    public void Execute_WithDimensions_AddsTextBoxWithSize()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "Test Text" },
            { "textboxWidth", 300.0 },
            { "textboxHeight", 150.0 }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        var textboxes = WordShapeHelper.FindAllTextboxes(doc);
        Assert.NotEmpty(textboxes);
        if (!IsEvaluationMode(AsposeLibraryType.Words))
        {
            var tb = textboxes[0];
            Assert.Equal(300.0, tb.Width);
            Assert.Equal(150.0, tb.Height);
        }
    }

    [Fact]
    public void Execute_WithPosition_AddsTextBoxAtPosition()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "Positioned Text" },
            { "positionX", 200.0 },
            { "positionY", 300.0 }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        var textboxes = WordShapeHelper.FindAllTextboxes(doc);
        Assert.NotEmpty(textboxes);
        if (!IsEvaluationMode(AsposeLibraryType.Words))
        {
            var tb = textboxes[0];
            Assert.Equal(200.0, tb.Left);
            Assert.Equal(300.0, tb.Top);
        }
    }

    [Fact]
    public void Execute_WithBackgroundColor_AddsStyledTextBox()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "Styled Text" },
            { "backgroundColor", "#FF0000" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        var textboxes = WordShapeHelper.FindAllTextboxes(doc);
        Assert.NotEmpty(textboxes);
        if (!IsEvaluationMode(AsposeLibraryType.Words))
        {
            var tb = textboxes[0];
            Assert.True(tb.Fill.Visible);
        }
    }

    [Fact]
    public void Execute_WithFontSettings_AddsFormattedTextBox()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "Formatted Text" },
            { "fontSize", 14.0 },
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
            Assert.Equal(14.0, runs[0].Font.Size);
            Assert.True(runs[0].Font.Bold);
        }
    }

    [Fact]
    public void Execute_WithoutText_ThrowsArgumentException()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion
}
