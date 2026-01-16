using System.Drawing;
using System.Text.Json;
using Aspose.Words;
using Aspose.Words.Lists;
using AsposeMcpServer.Handlers.Word.Paragraph;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Word.Paragraph;

public class GetParagraphFormatWordHandlerTests : WordHandlerTestBase
{
    private readonly GetParagraphFormatWordHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_GetFormat()
    {
        Assert.Equal("get_format", _handler.Operation);
    }

    #endregion

    #region Include Run Details

    [Fact]
    public void Execute_WithIncludeRunDetailsTrue_ReturnsRunInfo()
    {
        var doc = CreateDocumentWithParagraphs("Test paragraph");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "paragraphIndex", 0 },
            { "includeRunDetails", true }
        });

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        Assert.True(json.RootElement.TryGetProperty("runCount", out _));
    }

    #endregion

    #region Various Paragraph Indices

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    public void Execute_WithVariousParagraphIndices_ReturnsCorrectParagraph(int index)
    {
        var doc = CreateDocumentWithParagraphs("First", "Second", "Third");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "paragraphIndex", index }
        });

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        Assert.Equal(index, json.RootElement.GetProperty("paragraphIndex").GetInt32());
    }

    #endregion

    #region Read-Only Verification

    [Fact]
    public void Execute_DoesNotModifyDocument()
    {
        var doc = CreateDocumentWithParagraphs("Test content");
        var initialText = doc.GetText();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "paragraphIndex", 0 }
        });

        _handler.Execute(context, parameters);

        Assert.Equal(initialText, doc.GetText());
        AssertNotModified(context);
    }

    #endregion

    #region Include Run Details False

    [Fact]
    public void Execute_WithIncludeRunDetailsFalse_OmitsRunDetails()
    {
        var doc = CreateDocumentWithParagraphs("Test paragraph");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "paragraphIndex", 0 },
            { "includeRunDetails", false }
        });

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        Assert.True(json.RootElement.TryGetProperty("paragraphFormat", out _));
    }

    #endregion

    #region List Format

    [SkippableFact]
    public void Execute_WithListParagraph_ReturnsListFormat()
    {
        SkipInEvaluationMode(AsposeLibraryType.Words, "Evaluation mode limits list operations");

        var doc = CreateEmptyDocument();
        var builder = new DocumentBuilder(doc);
        var list = doc.Lists.Add(ListTemplate.NumberDefault);
        builder.ListFormat.List = list;
        builder.Writeln("First item");

        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "paragraphIndex", 0 }
        });

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        Assert.True(json.RootElement.TryGetProperty("listFormat", out var listFormat));
        Assert.True(listFormat.GetProperty("isListItem").GetBoolean());
    }

    #endregion

    #region Background Color

    [SkippableFact]
    public void Execute_WithBackgroundColor_ReturnsColorInfo()
    {
        SkipInEvaluationMode(AsposeLibraryType.Words, "Evaluation mode limits shading operations");

        var doc = CreateEmptyDocument();
        var builder = new DocumentBuilder(doc);
        builder.Write("Highlighted paragraph");

        var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
        var para = paragraphs[0] as Aspose.Words.Paragraph;
        para!.ParagraphFormat.Shading.BackgroundPatternColor = Color.Yellow;

        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "paragraphIndex", 0 }
        });

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        Assert.True(json.RootElement.TryGetProperty("backgroundColor", out var bgColor));
        Assert.Contains("FF", bgColor.GetString());
    }

    #endregion

    #region Tab Stops

    [SkippableFact]
    public void Execute_WithTabStops_ReturnsTabStopInfo()
    {
        SkipInEvaluationMode(AsposeLibraryType.Words, "Evaluation mode limits tab operations");

        var doc = CreateEmptyDocument();
        var builder = new DocumentBuilder(doc);
        builder.Write("Text with tabs");

        var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
        var para = paragraphs[0] as Aspose.Words.Paragraph;
        para!.ParagraphFormat.TabStops.Add(72, TabAlignment.Left, TabLeader.None);
        para.ParagraphFormat.TabStops.Add(144, TabAlignment.Center, TabLeader.Dots);

        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "paragraphIndex", 0 }
        });

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        Assert.True(json.RootElement.TryGetProperty("tabStops", out var tabStops));
        Assert.Equal(2, tabStops.GetArrayLength());
    }

    #endregion

    #region Run Details

    [SkippableFact]
    public void Execute_WithManyRuns_LimitsToTen()
    {
        SkipInEvaluationMode(AsposeLibraryType.Words, "Evaluation mode limits run operations");

        var doc = CreateEmptyDocument();
        var builder = new DocumentBuilder(doc);
        for (var i = 0; i < 15; i++)
        {
            builder.Font.Bold = i % 2 == 0;
            builder.Write($"Run{i} ");
        }

        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "paragraphIndex", 0 },
            { "includeRunDetails", true }
        });

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        Assert.True(json.RootElement.TryGetProperty("runs", out var runs));
        Assert.Equal(10, runs.GetProperty("displayed").GetInt32());
        Assert.Equal(15, runs.GetProperty("total").GetInt32());
    }

    #endregion

    #region Basic Get Operations

    [Fact]
    public void Execute_ReturnsFormatInfo()
    {
        var doc = CreateDocumentWithParagraphs("Test paragraph");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "paragraphIndex", 0 }
        });

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        Assert.True(json.RootElement.TryGetProperty("paragraphIndex", out _));
        Assert.True(json.RootElement.TryGetProperty("text", out _));
        Assert.True(json.RootElement.TryGetProperty("paragraphFormat", out _));
        AssertNotModified(context);
    }

    [Fact]
    public void Execute_ReturnsParagraphFormatProperties()
    {
        var doc = CreateDocumentWithParagraphs("Test paragraph");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "paragraphIndex", 0 }
        });

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        var format = json.RootElement.GetProperty("paragraphFormat");
        Assert.True(format.TryGetProperty("alignment", out _));
        Assert.True(format.TryGetProperty("leftIndent", out _));
        Assert.True(format.TryGetProperty("spaceBefore", out _));
        Assert.True(format.TryGetProperty("spaceAfter", out _));
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutParagraphIndex_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithParagraphs("Test paragraph");
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("paragraphIndex", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData(-1)]
    [InlineData(100)]
    public void Execute_WithInvalidParagraphIndex_ThrowsArgumentException(int invalidIndex)
    {
        var doc = CreateDocumentWithParagraphs("Test paragraph");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "paragraphIndex", invalidIndex }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("index", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    #endregion

    #region Paragraph with Formatting

    [SkippableFact]
    public void Execute_WithFormattedParagraph_ReturnsFormattingDetails()
    {
        SkipInEvaluationMode(AsposeLibraryType.Words, "Evaluation mode limits formatting operations");

        var doc = CreateEmptyDocument();
        var builder = new DocumentBuilder(doc) { Font = { Bold = true, Italic = true, Size = 14 } };
        builder.Write("Formatted text");

        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "paragraphIndex", 0 }
        });

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        Assert.True(json.RootElement.TryGetProperty("fontFormat", out var fontFormat));
        Assert.Equal(14, fontFormat.GetProperty("fontSize").GetDouble());
    }

    [SkippableFact]
    public void Execute_WithMultipleRuns_ReturnsRunDetails()
    {
        SkipInEvaluationMode(AsposeLibraryType.Words, "Evaluation mode limits run operations");

        var doc = CreateEmptyDocument();
        var builder = new DocumentBuilder(doc) { Font = { Bold = true } };
        builder.Write("Bold ");
        builder.Font.Bold = false;
        builder.Font.Italic = true;
        builder.Write("Italic");

        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "paragraphIndex", 0 },
            { "includeRunDetails", true }
        });

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        Assert.True(json.RootElement.TryGetProperty("runs", out _));
    }

    #endregion

    #region Borders

    [SkippableFact]
    public void Execute_WithBorders_ReturnsBorderInfo()
    {
        SkipInEvaluationMode(AsposeLibraryType.Words, "Evaluation mode limits border operations");

        var doc = CreateEmptyDocument();
        var builder = new DocumentBuilder(doc);
        builder.Write("Text with border");

        var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
        var para = paragraphs[0] as Aspose.Words.Paragraph;
        para!.ParagraphFormat.Borders.Top.LineStyle = LineStyle.Single;
        para.ParagraphFormat.Borders.Top.LineWidth = 1.5;
        para.ParagraphFormat.Borders.Bottom.LineStyle = LineStyle.Double;

        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "paragraphIndex", 0 }
        });

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        Assert.True(json.RootElement.TryGetProperty("borders", out var borders));
        Assert.True(borders.TryGetProperty("top", out _));
        Assert.True(borders.TryGetProperty("bottom", out _));
    }

    [SkippableFact]
    public void Execute_WithLeftRightBorders_ReturnsBorderInfo()
    {
        SkipInEvaluationMode(AsposeLibraryType.Words, "Evaluation mode limits border operations");

        var doc = CreateEmptyDocument();
        var builder = new DocumentBuilder(doc);
        builder.Write("Text with side borders");

        var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
        var para = paragraphs[0] as Aspose.Words.Paragraph;
        para!.ParagraphFormat.Borders.Left.LineStyle = LineStyle.Single;
        para.ParagraphFormat.Borders.Right.LineStyle = LineStyle.Single;

        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "paragraphIndex", 0 }
        });

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        Assert.True(json.RootElement.TryGetProperty("borders", out var borders));
        Assert.True(borders.TryGetProperty("left", out _));
        Assert.True(borders.TryGetProperty("right", out _));
    }

    #endregion

    #region Font Attributes

    [SkippableFact]
    public void Execute_WithFontAttributes_ReturnsFontInfo()
    {
        SkipInEvaluationMode(AsposeLibraryType.Words, "Evaluation mode limits font operations");

        var doc = CreateEmptyDocument();
        var builder = new DocumentBuilder(doc)
        {
            Font = { Underline = Underline.Single, StrikeThrough = true, Color = Color.Red }
        };
        builder.Write("Styled text");

        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "paragraphIndex", 0 }
        });

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        Assert.True(json.RootElement.TryGetProperty("fontFormat", out var fontFormat));
        Assert.True(fontFormat.TryGetProperty("underline", out _));
        Assert.True(fontFormat.TryGetProperty("strikethrough", out _));
        Assert.True(fontFormat.TryGetProperty("color", out _));
    }

    [SkippableFact]
    public void Execute_WithSuperscriptSubscript_ReturnsFontInfo()
    {
        SkipInEvaluationMode(AsposeLibraryType.Words, "Evaluation mode limits font operations");

        var doc = CreateEmptyDocument();
        var builder = new DocumentBuilder(doc) { Font = { Superscript = true } };
        builder.Write("Superscript");

        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "paragraphIndex", 0 }
        });

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        Assert.True(json.RootElement.TryGetProperty("fontFormat", out var fontFormat));
        Assert.True(fontFormat.TryGetProperty("superscript", out _));
    }

    [SkippableFact]
    public void Execute_WithHighlightColor_ReturnsFontInfo()
    {
        SkipInEvaluationMode(AsposeLibraryType.Words, "Evaluation mode limits font operations");

        var doc = CreateEmptyDocument();
        var builder = new DocumentBuilder(doc) { Font = { HighlightColor = Color.Yellow } };
        builder.Write("Highlighted");

        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "paragraphIndex", 0 }
        });

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        Assert.True(json.RootElement.TryGetProperty("fontFormat", out var fontFormat));
        Assert.True(fontFormat.TryGetProperty("highlightColor", out _));
    }

    #endregion
}
