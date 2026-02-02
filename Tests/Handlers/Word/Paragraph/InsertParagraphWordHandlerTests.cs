using Aspose.Words;
using AsposeMcpServer.Handlers.Word.Paragraph;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Word.Paragraph;

public class InsertParagraphWordHandlerTests : WordHandlerTestBase
{
    private readonly InsertParagraphWordHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Insert()
    {
        Assert.Equal("insert", _handler.Operation);
    }

    #endregion

    #region Result Message

    [Fact]
    public void Execute_ReturnsParagraphCountInMessage()
    {
        var doc = CreateDocumentWithParagraphs("Existing");
        var initialCount = doc.GetChildNodes(NodeType.Paragraph, true).Count;
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "New paragraph" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        var newCount = doc.GetChildNodes(NodeType.Paragraph, true).Count;
        Assert.Equal(initialCount + 1, newCount);
        AssertModified(context);
    }

    #endregion

    #region Basic Insert Operations

    [Fact]
    public void Execute_InsertsParagraph()
    {
        var doc = CreateDocumentWithParagraphs("Existing");
        var initialCount = doc.GetChildNodes(NodeType.Paragraph, true).Count;
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "New paragraph" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        var newCount = doc.GetChildNodes(NodeType.Paragraph, true).Count;
        Assert.Equal(initialCount + 1, newCount);
        if (!IsEvaluationMode(AsposeLibraryType.Words)) AssertContainsText(doc, "New paragraph");
        AssertModified(context);
    }

    [Theory]
    [InlineData("Hello World")]
    [InlineData("Test content")]
    [InlineData("Special chars: !@#$%")]
    public void Execute_InsertsVariousTexts(string text)
    {
        var doc = CreateDocumentWithParagraphs("Existing");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", text }
        });

        _handler.Execute(context, parameters);

        AssertContainsText(doc, text);
    }

    #endregion

    #region Insert Position

    [Fact]
    public void Execute_WithParagraphIndexMinus1_InsertsAtBeginning()
    {
        var doc = CreateDocumentWithParagraphs("Existing paragraph");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "New at beginning" },
            { "paragraphIndex", -1 }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        if (!IsEvaluationMode(AsposeLibraryType.Words))
        {
            var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true).Cast<Aspose.Words.Paragraph>().ToList();
            Assert.Contains(paragraphs, p => p.GetText().Contains("New at beginning"));
            Assert.Contains("New at beginning", paragraphs[0].GetText());
        }

        AssertModified(context);
    }

    [Fact]
    public void Execute_WithParagraphIndex_InsertsAfterSpecifiedParagraph()
    {
        var doc = CreateDocumentWithParagraphs("First", "Second");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "Inserted after first" },
            { "paragraphIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        if (!IsEvaluationMode(AsposeLibraryType.Words))
        {
            var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true).Cast<Aspose.Words.Paragraph>().ToList();
            Assert.Equal(3, paragraphs.Count);
            Assert.Contains("Inserted after first", paragraphs[1].GetText());
        }

        AssertModified(context);
    }

    [Fact]
    public void Execute_WithoutParagraphIndex_InsertsAtEnd()
    {
        var doc = CreateDocumentWithParagraphs("First");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "At end" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        if (!IsEvaluationMode(AsposeLibraryType.Words))
        {
            var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true).Cast<Aspose.Words.Paragraph>().ToList();
            var lastPara = paragraphs[^1];
            Assert.Contains("At end", lastPara.GetText());
        }

        AssertModified(context);
    }

    #endregion

    #region Formatting Options

    [Fact]
    public void Execute_WithStyleName_AppliesStyle()
    {
        var doc = CreateDocumentWithParagraphs("Existing");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "Styled paragraph" },
            { "styleName", "Heading 1" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        if (!IsEvaluationMode(AsposeLibraryType.Words))
        {
            var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true).Cast<Aspose.Words.Paragraph>().ToList();
            var styledPara = paragraphs.FirstOrDefault(p => p.GetText().Contains("Styled paragraph"));
            Assert.NotNull(styledPara);
            Assert.Equal("Heading 1", styledPara.ParagraphFormat.StyleName);
        }

        AssertModified(context);
    }

    [Fact]
    public void Execute_WithAlignment_AppliesAlignment()
    {
        var doc = CreateDocumentWithParagraphs("Existing");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "Centered paragraph" },
            { "alignment", "center" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        if (!IsEvaluationMode(AsposeLibraryType.Words))
        {
            var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true).Cast<Aspose.Words.Paragraph>().ToList();
            var centeredPara = paragraphs.FirstOrDefault(p => p.GetText().Contains("Centered paragraph"));
            Assert.NotNull(centeredPara);
            Assert.Equal(ParagraphAlignment.Center, centeredPara.ParagraphFormat.Alignment);
        }

        AssertModified(context);
    }

    [Fact]
    public void Execute_WithIndentation_AppliesIndentation()
    {
        var doc = CreateDocumentWithParagraphs("Existing");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "Indented paragraph" },
            { "indentLeft", 36.0 }
        });

        _handler.Execute(context, parameters);

        if (!IsEvaluationMode(AsposeLibraryType.Words))
        {
            var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true).Cast<Aspose.Words.Paragraph>().ToList();
            var indentedPara = paragraphs.FirstOrDefault(p => p.GetText().Contains("Indented paragraph"));
            Assert.NotNull(indentedPara);
            Assert.Equal(36.0, indentedPara.ParagraphFormat.LeftIndent);
        }

        AssertModified(context);
    }

    [Fact]
    public void Execute_WithSpacing_AppliesSpacing()
    {
        var doc = CreateDocumentWithParagraphs("Existing");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "Spaced paragraph" },
            { "spaceBefore", 12.0 },
            { "spaceAfter", 12.0 }
        });

        _handler.Execute(context, parameters);

        if (!IsEvaluationMode(AsposeLibraryType.Words))
        {
            var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true).Cast<Aspose.Words.Paragraph>().ToList();
            var spacedPara = paragraphs.FirstOrDefault(p => p.GetText().Contains("Spaced paragraph"));
            Assert.NotNull(spacedPara);
            Assert.Equal(12.0, spacedPara.ParagraphFormat.SpaceBefore);
            Assert.Equal(12.0, spacedPara.ParagraphFormat.SpaceAfter);
        }

        AssertModified(context);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutText_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithParagraphs("Existing");
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("text", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Execute_WithEmptyText_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithParagraphs("Existing");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("text", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Execute_WithInvalidParagraphIndex_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithParagraphs("Existing");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "New paragraph" },
            { "paragraphIndex", 100 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("index", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Execute_WithInvalidStyleName_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithParagraphs("Existing");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "New paragraph" },
            { "styleName", "NonExistentStyle" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("Style", ex.Message);
    }

    #endregion
}
