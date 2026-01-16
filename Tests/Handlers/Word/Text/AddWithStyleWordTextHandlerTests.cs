using Aspose.Words;
using AsposeMcpServer.Handlers.Word.Text;
using AsposeMcpServer.Tests.Helpers;
using WordParagraph = Aspose.Words.Paragraph;

namespace AsposeMcpServer.Tests.Handlers.Word.Text;

public class AddWithStyleWordTextHandlerTests : WordHandlerTestBase
{
    private readonly AddWithStyleWordTextHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_AddWithStyle()
    {
        Assert.Equal("add_with_style", _handler.Operation);
    }

    #endregion

    #region Basic Text Addition

    [Theory]
    [InlineData("Simple text")]
    [InlineData("Text with numbers 123")]
    [InlineData("Unicode: 中文測試")]
    public void Execute_AddsTextToDocument(string text)
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", text }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("added", result, StringComparison.OrdinalIgnoreCase);
        AssertContainsText(doc, text);
        AssertModified(context);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutText_ThrowsArgumentException()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("text", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    #endregion

    #region Combined Formatting

    [Fact]
    public void Execute_WithMultipleFormattingOptions_AppliesAll()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "Fully formatted" },
            { "bold", true },
            { "italic", true },
            { "fontSize", 16.0 },
            { "alignment", "center" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("Bold", result);
        Assert.Contains("Italic", result);
        var runs = doc.GetChildNodes(NodeType.Run, true).Cast<Run>().ToList();
        var run = runs.FirstOrDefault(r => r.Text.Contains("Fully formatted"));
        Assert.NotNull(run);
        Assert.True(run.Font.Bold);
        Assert.True(run.Font.Italic);
        Assert.Equal(16.0, run.Font.Size);
        AssertModified(context);
    }

    #endregion

    #region Style Application

    [Theory]
    [InlineData("Heading 1")]
    [InlineData("Heading 2")]
    [InlineData("Normal")]
    public void Execute_WithStyleName_AppliesStyle(string styleName)
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "Styled text" },
            { "styleName", styleName }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("style", result, StringComparison.OrdinalIgnoreCase);
        AssertContainsText(doc, "Styled text");
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithInvalidStyleName_ThrowsException()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "Test" },
            { "styleName", "NonExistentStyle12345" }
        });

        Assert.ThrowsAny<Exception>(() => _handler.Execute(context, parameters));
    }

    #endregion

    #region Font Formatting

    [Theory]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public void Execute_WithBoldAndItalic_AppliesFormatting(bool bold, bool italic)
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "Formatted text" },
            { "bold", bold },
            { "italic", italic }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("added", result, StringComparison.OrdinalIgnoreCase);
        var runs = doc.GetChildNodes(NodeType.Run, true).Cast<Run>().ToList();
        var run = runs.FirstOrDefault(r => r.Text.Contains("Formatted text"));
        Assert.NotNull(run);
        Assert.Equal(bold, run.Font.Bold);
        Assert.Equal(italic, run.Font.Italic);
        AssertModified(context);
    }

    [Theory]
    [InlineData(8.0)]
    [InlineData(12.0)]
    [InlineData(24.0)]
    public void Execute_WithFontSize_AppliesSize(double fontSize)
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "Sized text" },
            { "fontSize", fontSize }
        });

        _handler.Execute(context, parameters);

        var runs = doc.GetChildNodes(NodeType.Run, true).Cast<Run>().ToList();
        var run = runs.FirstOrDefault(r => r.Text.Contains("Sized text"));
        Assert.NotNull(run);
        Assert.Equal(fontSize, run.Font.Size);
        AssertModified(context);
    }

    [Theory]
    [InlineData("Arial")]
    [InlineData("Times New Roman")]
    public void Execute_WithFontName_AppliesFont(string fontName)
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "Custom font text" },
            { "fontName", fontName }
        });

        _handler.Execute(context, parameters);

        var runs = doc.GetChildNodes(NodeType.Run, true).Cast<Run>().ToList();
        var run = runs.FirstOrDefault(r => r.Text.Contains("Custom font text"));
        Assert.NotNull(run);
        Assert.Equal(fontName, run.Font.Name);
        AssertModified(context);
    }

    #endregion

    #region Paragraph Formatting

    [Theory]
    [InlineData("left", ParagraphAlignment.Left)]
    [InlineData("right", ParagraphAlignment.Right)]
    [InlineData("center", ParagraphAlignment.Center)]
    [InlineData("justify", ParagraphAlignment.Justify)]
    public void Execute_WithAlignment_AppliesAlignment(string alignment, ParagraphAlignment expected)
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "Aligned text" },
            { "alignment", alignment }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("Alignment", result, StringComparison.OrdinalIgnoreCase);
        var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true).Cast<WordParagraph>().ToList();
        var para = paragraphs.FirstOrDefault(p => p.GetText().Contains("Aligned text"));
        Assert.NotNull(para);
        Assert.Equal(expected, para.ParagraphFormat.Alignment);
        AssertModified(context);
    }

    [Theory]
    [InlineData(1)]
    [InlineData(2)]
    [InlineData(3)]
    public void Execute_WithIndentLevel_AppliesIndent(int indentLevel)
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "Indented text" },
            { "indentLevel", indentLevel }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("Indent level", result, StringComparison.OrdinalIgnoreCase);
        AssertModified(context);
    }

    #endregion

    #region Paragraph Position

    [Theory]
    [InlineData(-1)]
    [InlineData(0)]
    public void Execute_WithParagraphIndex_InsertsAtPosition(int paragraphIndex)
    {
        var doc = CreateDocumentWithText("Existing text");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "New text" },
            { "paragraphIndexForAdd", paragraphIndex }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("added", result, StringComparison.OrdinalIgnoreCase);
        AssertContainsText(doc, "New text");
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithInvalidParagraphIndex_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithText("Test");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "New" },
            { "paragraphIndexForAdd", 100 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion

    #region Underline and Color

    [Fact]
    public void Execute_WithUnderline_AppliesUnderline()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "Underlined text" },
            { "underline", true }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("Underline", result);
        var runs = doc.GetChildNodes(NodeType.Run, true).Cast<Run>().ToList();
        var run = runs.FirstOrDefault(r => r.Text.Contains("Underlined text"));
        Assert.NotNull(run);
        Assert.NotEqual(Underline.None, run.Font.Underline);
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithColor_AppliesColor()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "Colored text" },
            { "color", "#FF0000" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("Color", result);
        AssertModified(context);
    }

    #endregion

    #region Left and First Line Indent

    [Theory]
    [InlineData(36.0)]
    [InlineData(72.0)]
    public void Execute_WithLeftIndent_AppliesIndent(double leftIndent)
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "Left indented text" },
            { "leftIndent", leftIndent }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("Left indent", result);
        AssertModified(context);
    }

    [Theory]
    [InlineData(18.0)]
    [InlineData(36.0)]
    public void Execute_WithFirstLineIndent_AppliesIndent(double firstLineIndent)
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "First line indented" },
            { "firstLineIndent", firstLineIndent }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("First line indent", result);
        AssertModified(context);
    }

    #endregion

    #region Font Name Variants

    [Fact]
    public void Execute_WithFontNameAscii_AppliesAsciiFont()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "ASCII font text" },
            { "fontNameAscii", "Courier New" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("Font (ASCII)", result);
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithFontNameFarEast_AppliesFarEastFont()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "Far East font text" },
            { "fontNameFarEast", "MS Gothic" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("Font (Far East)", result);
        AssertModified(context);
    }

    #endregion

    #region Tab Stops

    [Fact]
    public void Execute_WithTabStops_AppliesTabStops()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var tabStopsJson = "[{\"position\": 72, \"alignment\": \"Left\", \"leader\": \"None\"}]";
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "Text with tabs" },
            { "tabStops", tabStopsJson }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("added", result, StringComparison.OrdinalIgnoreCase);
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithMultipleTabStops_AppliesAll()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var tabStopsJson =
            "[{\"position\": 72, \"alignment\": \"Left\"}, {\"position\": 144, \"alignment\": \"Center\"}, {\"position\": 216, \"alignment\": \"Right\"}]";
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "Multi-tab text" },
            { "tabStops", tabStopsJson }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("added", result, StringComparison.OrdinalIgnoreCase);
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithTabStopLeaders_AppliesLeaders()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var tabStopsJson = "[{\"position\": 144, \"alignment\": \"Right\", \"leader\": \"Dots\"}]";
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "Dotted leader" },
            { "tabStops", tabStopsJson }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("added", result, StringComparison.OrdinalIgnoreCase);
        AssertModified(context);
    }

    #endregion
}
