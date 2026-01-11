using Aspose.Words;
using AsposeMcpServer.Handlers.Word.Text;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Word.Text;

public class AddWordTextHandlerTests : WordHandlerTestBase
{
    private readonly AddWordTextHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Add()
    {
        Assert.Equal("add", _handler.Operation);
    }

    #endregion

    #region Font Formatting - Bold and Italic

    [Theory]
    [InlineData("bold", true, true, false)]
    [InlineData("bold", false, false, false)]
    [InlineData("italic", true, false, true)]
    [InlineData("italic", false, false, false)]
    public void Execute_WithBoldOrItalic_AppliesFormatting(string paramName, bool paramValue, bool expectedBold,
        bool expectedItalic)
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "Formatted Text" },
            { paramName, paramValue }
        });

        _handler.Execute(context, parameters);

        var runs = doc.GetChildNodes(NodeType.Run, true).Cast<Run>().ToList();
        var run = runs.FirstOrDefault(r => r.Text.Contains("Formatted Text"));
        Assert.NotNull(run);
        Assert.Equal(expectedBold, run.Font.Bold);
        Assert.Equal(expectedItalic, run.Font.Italic);
        if (paramValue)
            AssertModified(context);
    }

    #endregion

    #region Font Formatting - Strikethrough, Superscript, Subscript

    [Theory]
    [InlineData("strikethrough", true, true, false, false)]
    [InlineData("superscript", true, false, true, false)]
    [InlineData("subscript", true, false, false, true)]
    public void Execute_WithSpecialFormatting_AppliesCorrectly(string paramName, bool paramValue, bool expectedStrike,
        bool expectedSuper, bool expectedSub)
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "Special Text" },
            { paramName, paramValue }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains(paramName, result, StringComparison.OrdinalIgnoreCase);

        var runs = doc.GetChildNodes(NodeType.Run, true).Cast<Run>().ToList();
        var run = runs.FirstOrDefault(r => r.Text.Contains("Special Text"));
        Assert.NotNull(run);
        Assert.Equal(expectedStrike, run.Font.StrikeThrough);
        Assert.Equal(expectedSuper, run.Font.Superscript);
        Assert.Equal(expectedSub, run.Font.Subscript);
        AssertModified(context);
    }

    #endregion

    #region Font Formatting - Color

    [Theory]
    [InlineData("#FF0000", 255, 0, 0)]
    [InlineData("#00FF00", 0, 255, 0)]
    [InlineData("#0000FF", 0, 0, 255)]
    [InlineData("FF0000", 255, 0, 0)]
    [InlineData("red", 255, 0, 0)]
    public void Execute_WithColor_AppliesColor(string colorValue, int expectedR, int expectedG, int expectedB)
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "Colored Text" },
            { "color", colorValue }
        });

        _handler.Execute(context, parameters);

        var runs = doc.GetChildNodes(NodeType.Run, true).Cast<Run>().ToList();
        var run = runs.FirstOrDefault(r => r.Text.Contains("Colored Text"));
        Assert.NotNull(run);
        Assert.Equal(expectedR, run.Font.Color.R);
        Assert.Equal(expectedG, run.Font.Color.G);
        Assert.Equal(expectedB, run.Font.Color.B);
        AssertModified(context);
    }

    #endregion

    #region Combined Formatting

    [Fact]
    public void Execute_WithMultipleFormats_AppliesAllFormats()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "Fully Formatted" },
            { "bold", true },
            { "italic", true },
            { "underline", "single" },
            { "fontSize", 16.0 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("bold", result, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("italic", result, StringComparison.OrdinalIgnoreCase);

        var runs = doc.GetChildNodes(NodeType.Run, true).Cast<Run>().ToList();
        var run = runs.FirstOrDefault(r => r.Text.Contains("Fully Formatted"));
        Assert.NotNull(run);
        Assert.True(run.Font.Bold, "Should be bold");
        Assert.True(run.Font.Italic, "Should be italic");
        Assert.Equal(Underline.Single, run.Font.Underline);
        Assert.Equal(16.0, run.Font.Size);
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

    #region Basic Text Addition

    [Fact]
    public void Execute_AddsTextToDocument()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "Hello World" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("added", result, StringComparison.OrdinalIgnoreCase);
        AssertContainsText(doc, "Hello World");
        AssertModified(context);
    }

    [Theory]
    [InlineData("Simple text")]
    [InlineData("Text with numbers 12345")]
    [InlineData("Special chars: !@#$%^&*()")]
    [InlineData("Unicode: 中文測試 日本語")]
    [InlineData("")]
    public void Execute_WithVariousTextContent_AddsText(string text)
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", text }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("added", result, StringComparison.OrdinalIgnoreCase);
        if (!string.IsNullOrEmpty(text))
            AssertContainsText(doc, text);
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithNewlines_CreatesMultipleParagraphs()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "Line 1\nLine 2\nLine 3" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("added", result, StringComparison.OrdinalIgnoreCase);
        AssertContainsText(doc, "Line 1");
        AssertContainsText(doc, "Line 2");
        AssertContainsText(doc, "Line 3");

        var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
        Assert.True(paragraphs.Count >= 3, "Should create multiple paragraphs for multiline text");
        AssertModified(context);
    }

    [Theory]
    [InlineData("\r\n")]
    [InlineData("\n")]
    [InlineData("\r")]
    public void Execute_WithDifferentLineEndings_HandlesCorrectly(string lineEnding)
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", $"Line1{lineEnding}Line2" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("added", result, StringComparison.OrdinalIgnoreCase);
        AssertContainsText(doc, "Line1");
        AssertContainsText(doc, "Line2");
        AssertModified(context);
    }

    #endregion

    #region Font Formatting - Underline

    [Theory]
    [InlineData("single", Underline.Single)]
    [InlineData("double", Underline.Double)]
    [InlineData("dotted", Underline.Dotted)]
    [InlineData("dash", Underline.Dash)]
    public void Execute_WithUnderlineStyles_AppliesCorrectUnderlineType(string underlineStyle, Underline expected)
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", $"Underline {underlineStyle}" },
            { "underline", underlineStyle }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("underline", result, StringComparison.OrdinalIgnoreCase);

        var runs = doc.GetChildNodes(NodeType.Run, true).Cast<Run>().ToList();
        var run = runs.FirstOrDefault(r => r.Text.Contains($"Underline {underlineStyle}"));
        Assert.NotNull(run);
        Assert.Equal(expected, run.Font.Underline);
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithUnderlineNone_DoesNotApplyUnderline()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "No Underline" },
            { "underline", "none" }
        });

        _handler.Execute(context, parameters);

        var runs = doc.GetChildNodes(NodeType.Run, true).Cast<Run>().ToList();
        var run = runs.FirstOrDefault(r => r.Text.Contains("No Underline"));
        Assert.NotNull(run);
        Assert.Equal(Underline.None, run.Font.Underline);
    }

    #endregion

    #region Font Formatting - Font Name and Size

    [Theory]
    [InlineData("Arial")]
    [InlineData("Times New Roman")]
    [InlineData("Calibri")]
    public void Execute_WithFontName_AppliesFontName(string fontName)
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "Custom Font" },
            { "fontName", fontName }
        });

        _handler.Execute(context, parameters);

        var runs = doc.GetChildNodes(NodeType.Run, true).Cast<Run>().ToList();
        var run = runs.FirstOrDefault(r => r.Text.Contains("Custom Font"));
        Assert.NotNull(run);
        Assert.Equal(fontName, run.Font.Name);
        AssertModified(context);
    }

    [Theory]
    [InlineData(8.0)]
    [InlineData(12.0)]
    [InlineData(14.0)]
    [InlineData(24.0)]
    [InlineData(72.0)]
    public void Execute_WithFontSize_AppliesFontSize(double fontSize)
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "Sized Text" },
            { "fontSize", fontSize }
        });

        _handler.Execute(context, parameters);

        var runs = doc.GetChildNodes(NodeType.Run, true).Cast<Run>().ToList();
        var run = runs.FirstOrDefault(r => r.Text.Contains("Sized Text"));
        Assert.NotNull(run);
        Assert.Equal(fontSize, run.Font.Size);
        AssertModified(context);
    }

    #endregion

    #region Document State

    [Fact]
    public void Execute_PreservesExistingContent()
    {
        var doc = CreateDocumentWithText("Existing Content");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "New Content" }
        });

        _handler.Execute(context, parameters);

        AssertContainsText(doc, "Existing Content");
        AssertContainsText(doc, "New Content");
    }

    [Fact]
    public void Execute_AddsToEndOfDocument()
    {
        var doc = CreateDocumentWithText("First");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "Last" }
        });

        _handler.Execute(context, parameters);

        var text = GetDocumentText(doc);
        var firstIndex = text.IndexOf("First", StringComparison.Ordinal);
        var lastIndex = text.IndexOf("Last", StringComparison.Ordinal);
        Assert.True(lastIndex > firstIndex, "New text should be added after existing text");
    }

    #endregion
}
