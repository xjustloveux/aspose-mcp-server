using System.Drawing;
using Aspose.Words;
using AsposeMcpServer.Handlers.Word.Text;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;
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
    [InlineData("Unicode: 中�?測試")]
    public void Execute_AddsTextToDocument(string text)
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", text }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        if (!IsEvaluationMode(AsposeLibraryType.Words)) AssertContainsText(doc, text);
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

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        if (!IsEvaluationMode(AsposeLibraryType.Words))
        {
            var runs = doc.GetChildNodes(NodeType.Run, true).Cast<Run>().ToList();
            var run = runs.FirstOrDefault(r => r.Text.Contains("Fully formatted"));
            Assert.NotNull(run);
            Assert.True(run.Font.Bold);
            Assert.True(run.Font.Italic);
            Assert.Equal(16.0, run.Font.Size);
        }

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

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        if (!IsEvaluationMode(AsposeLibraryType.Words))
        {
            AssertContainsText(doc, "Styled text");
            var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true).Cast<WordParagraph>().ToList();
            var para = paragraphs.FirstOrDefault(p => p.GetText().Contains("Styled text"));
            Assert.NotNull(para);
            Assert.Equal(styleName, para.ParagraphFormat.StyleName);
        }

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

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        if (!IsEvaluationMode(AsposeLibraryType.Words))
        {
            var runs = doc.GetChildNodes(NodeType.Run, true).Cast<Run>().ToList();
            var run = runs.FirstOrDefault(r => r.Text.Contains("Formatted text"));
            Assert.NotNull(run);
            Assert.Equal(bold, run.Font.Bold);
            Assert.Equal(italic, run.Font.Italic);
        }

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

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        if (!IsEvaluationMode(AsposeLibraryType.Words))
        {
            var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true).Cast<WordParagraph>().ToList();
            var para = paragraphs.FirstOrDefault(p => p.GetText().Contains("Aligned text"));
            Assert.NotNull(para);
            Assert.Equal(expected, para.ParagraphFormat.Alignment);
        }

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

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        if (!IsEvaluationMode(AsposeLibraryType.Words))
        {
            var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true).Cast<WordParagraph>().ToList();
            var para = paragraphs.FirstOrDefault(p => p.GetText().Contains("Indented text"));
            Assert.NotNull(para);
            Assert.Equal(indentLevel * 36, para.ParagraphFormat.LeftIndent);
        }

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

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        if (!IsEvaluationMode(AsposeLibraryType.Words)) AssertContainsText(doc, "New text");
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

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        if (!IsEvaluationMode(AsposeLibraryType.Words))
        {
            var runs = doc.GetChildNodes(NodeType.Run, true).Cast<Run>().ToList();
            var run = runs.FirstOrDefault(r => r.Text.Contains("Underlined text"));
            Assert.NotNull(run);
            Assert.NotEqual(Underline.None, run.Font.Underline);
        }

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

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        if (!IsEvaluationMode(AsposeLibraryType.Words))
        {
            var runs = doc.GetChildNodes(NodeType.Run, true).Cast<Run>().ToList();
            var run = runs.FirstOrDefault(r => r.Text.Contains("Colored text"));
            Assert.NotNull(run);
            Assert.Equal(Color.FromArgb(255, 0, 0).ToArgb(), run.Font.Color.ToArgb());
        }

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

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        if (!IsEvaluationMode(AsposeLibraryType.Words))
        {
            var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true).Cast<WordParagraph>().ToList();
            var para = paragraphs.FirstOrDefault(p => p.GetText().Contains("Left indented text"));
            Assert.NotNull(para);
            Assert.Equal(leftIndent, para.ParagraphFormat.LeftIndent);
        }

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

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        if (!IsEvaluationMode(AsposeLibraryType.Words))
        {
            var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true).Cast<WordParagraph>().ToList();
            var para = paragraphs.FirstOrDefault(p => p.GetText().Contains("First line indented"));
            Assert.NotNull(para);
            Assert.Equal(firstLineIndent, para.ParagraphFormat.FirstLineIndent);
        }

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

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        if (!IsEvaluationMode(AsposeLibraryType.Words))
        {
            var runs = doc.GetChildNodes(NodeType.Run, true).Cast<Run>().ToList();
            var run = runs.FirstOrDefault(r => r.Text.Contains("ASCII font text"));
            Assert.NotNull(run);
            Assert.Equal("Courier New", run.Font.NameAscii);
        }

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

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        if (!IsEvaluationMode(AsposeLibraryType.Words))
        {
            var runs = doc.GetChildNodes(NodeType.Run, true).Cast<Run>().ToList();
            var run = runs.FirstOrDefault(r => r.Text.Contains("Far East font text"));
            Assert.NotNull(run);
            Assert.Equal("MS Gothic", run.Font.NameFarEast);
        }

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

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        if (!IsEvaluationMode(AsposeLibraryType.Words))
        {
            var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true).Cast<WordParagraph>().ToList();
            var para = paragraphs.FirstOrDefault(p => p.GetText().Contains("Text with tabs"));
            Assert.NotNull(para);
            Assert.True(para.ParagraphFormat.TabStops.Count > 0);
        }

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

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        if (!IsEvaluationMode(AsposeLibraryType.Words))
        {
            var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true).Cast<WordParagraph>().ToList();
            var para = paragraphs.FirstOrDefault(p => p.GetText().Contains("Multi-tab text"));
            Assert.NotNull(para);
            Assert.Equal(3, para.ParagraphFormat.TabStops.Count);
        }

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

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        if (!IsEvaluationMode(AsposeLibraryType.Words))
        {
            var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true).Cast<WordParagraph>().ToList();
            var para = paragraphs.FirstOrDefault(p => p.GetText().Contains("Dotted leader"));
            Assert.NotNull(para);
            Assert.True(para.ParagraphFormat.TabStops.Count > 0);
        }

        AssertModified(context);
    }

    #endregion
}
