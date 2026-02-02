using System.Text.Json.Nodes;
using Aspose.Words;
using AsposeMcpServer.Handlers.Word.Paragraph;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Word.Paragraph;

public class EditParagraphWordHandlerTests : WordHandlerTestBase
{
    private readonly EditParagraphWordHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Edit()
    {
        Assert.Equal("edit", _handler.Operation);
    }

    #endregion

    #region Section Index

    [Fact]
    public void Execute_WithSectionIndex_EditsInSpecificSection()
    {
        var doc = CreateDocumentWithParagraphs("Section content");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "paragraphIndex", 0 },
            { "sectionIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        if (!IsEvaluationMode(AsposeLibraryType.Words)) AssertContainsText(doc, "Section content");
        AssertModified(context);
    }

    #endregion

    #region Style Settings

    [Fact]
    public void Execute_WithValidStyleName_AppliesStyle()
    {
        var doc = CreateDocumentWithParagraphs("Test paragraph");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "paragraphIndex", 0 },
            { "styleName", "Normal" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        var para = doc.GetChildNodes(NodeType.Paragraph, true).Cast<Aspose.Words.Paragraph>().First();
        Assert.Equal("Normal", para.ParagraphFormat.StyleName);
        AssertModified(context);
    }

    #endregion

    #region Tab Stops

    [Fact]
    public void Execute_WithTabStops_AppliesTabStops()
    {
        var doc = CreateDocumentWithParagraphs("Test paragraph");
        var context = CreateContext(doc);
        var tabStops = new JsonArray
        {
            new JsonObject
            {
                ["position"] = 100.0,
                ["alignment"] = "center",
                ["leader"] = "dots"
            }
        };
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "paragraphIndex", 0 },
            { "tabStops", tabStops }
        });

        _handler.Execute(context, parameters);

        var para = doc.GetChildNodes(NodeType.Paragraph, true).Cast<Aspose.Words.Paragraph>().First();
        Assert.True(para.ParagraphFormat.TabStops.Count > 0);
        Assert.Equal(100.0, para.ParagraphFormat.TabStops[0].Position);
        Assert.Equal(TabAlignment.Center, para.ParagraphFormat.TabStops[0].Alignment);
        Assert.Equal(TabLeader.Dots, para.ParagraphFormat.TabStops[0].Leader);
        AssertModified(context);
    }

    #endregion

    #region Basic Edit Operations

    [Fact]
    public void Execute_EditsParagraph()
    {
        var doc = CreateDocumentWithParagraphs("Original text");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "paragraphIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        if (!IsEvaluationMode(AsposeLibraryType.Words)) AssertContainsText(doc, "Original text");
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithText_UpdatesContent()
    {
        var doc = CreateDocumentWithParagraphs("Original text");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "paragraphIndex", 0 },
            { "text", "Updated text" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        if (!IsEvaluationMode(AsposeLibraryType.Words))
        {
            AssertContainsText(doc, "Updated text");
            AssertDoesNotContainText(doc, "Original text");
        }

        AssertModified(context);
    }

    #endregion

    #region Paragraph Index

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    public void Execute_WithVariousParagraphIndices_EditsCorrectParagraph(int index)
    {
        var doc = CreateDocumentWithParagraphs("First", "Second", "Third");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "paragraphIndex", index }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true).Cast<Aspose.Words.Paragraph>().ToList();
        Assert.True(index < paragraphs.Count);
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithParagraphIndexMinus1_EditsLastParagraph()
    {
        var doc = CreateDocumentWithParagraphs("First", "Second", "Last");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "paragraphIndex", -1 },
            { "text", "Modified last" }
        });

        _handler.Execute(context, parameters);

        AssertContainsText(doc, "Modified last");
        AssertModified(context);
    }

    #endregion

    #region Formatting Options

    [Fact]
    public void Execute_WithAlignment_AppliesAlignment()
    {
        var doc = CreateDocumentWithParagraphs("Test paragraph");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "paragraphIndex", 0 },
            { "alignment", "center" }
        });

        _handler.Execute(context, parameters);

        var para = doc.GetChildNodes(NodeType.Paragraph, true).Cast<Aspose.Words.Paragraph>().First();
        Assert.Equal(ParagraphAlignment.Center, para.ParagraphFormat.Alignment);
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithBold_AppliesBold()
    {
        var doc = CreateDocumentWithParagraphs("Test paragraph");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "paragraphIndex", 0 },
            { "bold", true }
        });

        _handler.Execute(context, parameters);

        var runs = doc.GetChildNodes(NodeType.Run, true).Cast<Run>().ToList();
        Assert.NotEmpty(runs);
        Assert.True(runs[0].Font.Bold);
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithFontSize_AppliesFontSize()
    {
        var doc = CreateDocumentWithParagraphs("Test paragraph");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "paragraphIndex", 0 },
            { "fontSize", 14.0 }
        });

        _handler.Execute(context, parameters);

        var runs = doc.GetChildNodes(NodeType.Run, true).Cast<Run>().ToList();
        Assert.NotEmpty(runs);
        Assert.Equal(14.0, runs[0].Font.Size);
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithIndentation_AppliesIndentation()
    {
        var doc = CreateDocumentWithParagraphs("Test paragraph");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "paragraphIndex", 0 },
            { "indentLeft", 36.0 },
            { "indentRight", 18.0 }
        });

        _handler.Execute(context, parameters);

        var para = doc.GetChildNodes(NodeType.Paragraph, true).Cast<Aspose.Words.Paragraph>().First();
        Assert.Equal(36.0, para.ParagraphFormat.LeftIndent);
        Assert.Equal(18.0, para.ParagraphFormat.RightIndent);
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithSpacing_AppliesSpacing()
    {
        var doc = CreateDocumentWithParagraphs("Test paragraph");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "paragraphIndex", 0 },
            { "spaceBefore", 12.0 },
            { "spaceAfter", 12.0 }
        });

        _handler.Execute(context, parameters);

        var para = doc.GetChildNodes(NodeType.Paragraph, true).Cast<Aspose.Words.Paragraph>().First();
        Assert.Equal(12.0, para.ParagraphFormat.SpaceBefore);
        Assert.Equal(12.0, para.ParagraphFormat.SpaceAfter);
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithLineSpacing_AppliesLineSpacing()
    {
        var doc = CreateDocumentWithParagraphs("Test paragraph");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "paragraphIndex", 0 },
            { "lineSpacingRule", "double" }
        });

        _handler.Execute(context, parameters);

        var para = doc.GetChildNodes(NodeType.Paragraph, true).Cast<Aspose.Words.Paragraph>().First();
        Assert.Equal(LineSpacingRule.Multiple, para.ParagraphFormat.LineSpacingRule);
        Assert.Equal(2.0, para.ParagraphFormat.LineSpacing);
        AssertModified(context);
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

    [Fact]
    public void Execute_WithInvalidParagraphIndex_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithParagraphs("Test paragraph");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "paragraphIndex", 100 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("index", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Execute_WithInvalidSectionIndex_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithParagraphs("Test paragraph");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "paragraphIndex", 0 },
            { "sectionIndex", 100 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("Section", ex.Message);
    }

    [Fact]
    public void Execute_WithInvalidStyleName_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithParagraphs("Test paragraph");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "paragraphIndex", 0 },
            { "styleName", "NonExistentStyle" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("Style", ex.Message);
    }

    #endregion

    #region Font Settings

    [Fact]
    public void Execute_WithItalic_AppliesItalic()
    {
        var doc = CreateDocumentWithParagraphs("Test paragraph");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "paragraphIndex", 0 },
            { "italic", true }
        });

        _handler.Execute(context, parameters);

        var runs = doc.GetChildNodes(NodeType.Run, true).Cast<Run>().ToList();
        Assert.NotEmpty(runs);
        Assert.True(runs[0].Font.Italic);
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithUnderline_AppliesUnderline()
    {
        var doc = CreateDocumentWithParagraphs("Test paragraph");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "paragraphIndex", 0 },
            { "underline", true }
        });

        _handler.Execute(context, parameters);

        var runs = doc.GetChildNodes(NodeType.Run, true).Cast<Run>().ToList();
        Assert.NotEmpty(runs);
        Assert.Equal(Underline.Single, runs[0].Font.Underline);
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithColor_AppliesFontColor()
    {
        var doc = CreateDocumentWithParagraphs("Test paragraph");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "paragraphIndex", 0 },
            { "color", "#FF0000" }
        });

        _handler.Execute(context, parameters);

        var runs = doc.GetChildNodes(NodeType.Run, true).Cast<Run>().ToList();
        Assert.NotEmpty(runs);
        Assert.Equal(255, runs[0].Font.Color.R);
        Assert.Equal(0, runs[0].Font.Color.G);
        Assert.Equal(0, runs[0].Font.Color.B);
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithFontNameAsciiAndFarEast_SetsBothFonts()
    {
        var doc = CreateDocumentWithParagraphs("Test paragraph");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "paragraphIndex", 0 },
            { "fontNameAscii", "Arial" },
            { "fontNameFarEast", "MS Gothic" }
        });

        _handler.Execute(context, parameters);

        var runs = doc.GetChildNodes(NodeType.Run, true).Cast<Run>().ToList();
        Assert.NotEmpty(runs);
        Assert.Equal("Arial", runs[0].Font.NameAscii);
        Assert.Equal("MS Gothic", runs[0].Font.NameFarEast);
        AssertModified(context);
    }

    #endregion

    #region Paragraph Format Settings

    [Fact]
    public void Execute_WithFirstLineIndent_AppliesFirstLineIndent()
    {
        var doc = CreateDocumentWithParagraphs("Test paragraph");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "paragraphIndex", 0 },
            { "firstLineIndent", 36.0 }
        });

        _handler.Execute(context, parameters);

        var para = doc.GetChildNodes(NodeType.Paragraph, true).Cast<Aspose.Words.Paragraph>().First();
        Assert.Equal(36.0, para.ParagraphFormat.FirstLineIndent);
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithLineSpacingValue_AppliesLineSpacing()
    {
        var doc = CreateDocumentWithParagraphs("Test paragraph");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "paragraphIndex", 0 },
            { "lineSpacing", 18.0 },
            { "lineSpacingRule", "exactly" }
        });

        _handler.Execute(context, parameters);

        var para = doc.GetChildNodes(NodeType.Paragraph, true).Cast<Aspose.Words.Paragraph>().First();
        Assert.Equal(LineSpacingRule.Exactly, para.ParagraphFormat.LineSpacingRule);
        Assert.Equal(18.0, para.ParagraphFormat.LineSpacing);
        AssertModified(context);
    }

    [Theory]
    [InlineData("single", 1.0)]
    [InlineData("oneandhalf", 1.5)]
    [InlineData("double", 2.0)]
    public void Execute_WithLineSpacingRuleOnly_AppliesDefaultSpacing(string rule, double expectedSpacing)
    {
        var doc = CreateDocumentWithParagraphs("Test paragraph");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "paragraphIndex", 0 },
            { "lineSpacingRule", rule }
        });

        _handler.Execute(context, parameters);

        var para = doc.GetChildNodes(NodeType.Paragraph, true).Cast<Aspose.Words.Paragraph>().First();
        Assert.Equal(LineSpacingRule.Multiple, para.ParagraphFormat.LineSpacingRule);
        Assert.Equal(expectedSpacing, para.ParagraphFormat.LineSpacing);
        AssertModified(context);
    }

    #endregion
}
