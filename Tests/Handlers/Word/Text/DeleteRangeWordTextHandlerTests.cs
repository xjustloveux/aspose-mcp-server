using Aspose.Words;
using AsposeMcpServer.Handlers.Word.Text;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Word.Text;

public class DeleteRangeWordTextHandlerTests : WordHandlerTestBase
{
    private readonly DeleteRangeWordTextHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_DeleteRange()
    {
        Assert.Equal("delete_range", _handler.Operation);
    }

    #endregion

    #region Field Preservation (Issue #7)

    [Fact]
    public void Execute_CharRangeOverlappingFieldCode_PreservesFieldCode()
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Write("Hello ");
        builder.InsertField("MERGEFIELD Name");
        builder.Write(" World");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "startParagraphIndex", 0 },
            { "startCharIndex", 3 },
            { "endParagraphIndex", 0 },
            { "endCharIndex", 10 }
        });

        _handler.Execute(context, parameters);

        Assert.Equal(1, doc.Range.Fields.Count);
        Assert.Equal("MERGEFIELD Name", doc.Range.Fields[0].GetFieldCode().Trim());
    }

    #endregion

    #region Multi-Paragraph Deletion

    [Fact]
    public void Execute_AcrossParagraphs_DeletesRange()
    {
        var doc = CreateDocumentWithParagraphs("First paragraph", "Second paragraph", "Third paragraph");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "startParagraphIndex", 0 },
            { "startCharIndex", 5 },
            { "endParagraphIndex", 1 },
            { "endCharIndex", 5 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("deleted", result.Message, StringComparison.OrdinalIgnoreCase);
        AssertContainsText(doc, "First");
        AssertContainsText(doc, "Third paragraph");
        AssertModified(context);
    }

    #endregion

    #region Section Index

    [Fact]
    public void Execute_WithSectionIndex_DeletesInCorrectSection()
    {
        var doc = CreateDocumentWithText("Hello World");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "startParagraphIndex", 0 },
            { "startCharIndex", 0 },
            { "endParagraphIndex", 0 },
            { "endCharIndex", 5 },
            { "sectionIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("deleted", result.Message, StringComparison.OrdinalIgnoreCase);
        AssertContainsText(doc, "World");
        AssertModified(context);
    }

    #endregion

    #region Document State

    [Fact]
    public void Execute_PreservesContentOutsideRange()
    {
        var doc = CreateDocumentWithText("Hello World Test");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "startParagraphIndex", 0 },
            { "startCharIndex", 6 },
            { "endParagraphIndex", 0 },
            { "endCharIndex", 12 }
        });

        _handler.Execute(context, parameters);

        AssertContainsText(doc, "Hello");
        AssertContainsText(doc, "Test");
        AssertModified(context);
    }

    #endregion

    #region Global Paragraph Index (Issue #1 unification)

    [Fact]
    public void Execute_WithStoryTypeHeader_TargetsHeaderParagraph()
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Write("Body");
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.Write("HeaderTrim");

        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "startParagraphIndex", 0 },
            { "endParagraphIndex", 0 },
            { "startCharIndex", 0 },
            { "endCharIndex", 4 },
            { "storyType", "Header" }
        });

        _handler.Execute(context, parameters);

        var headerPara = doc.FirstSection.HeadersFooters[HeaderFooterType.HeaderPrimary]
            .GetChildNodes(NodeType.Paragraph, true).Cast<Aspose.Words.Paragraph>().First();
        Assert.StartsWith("erTrim", headerPara.GetText());
    }

    #endregion

    #region Error Handling - Invalid Indices

    [Theory]
    [InlineData(-1, 0)]
    [InlineData(0, -1)]
    [InlineData(100, 0)]
    [InlineData(0, 100)]
    public void Execute_WithInvalidParagraphIndices_ThrowsArgumentException(int startPara, int endPara)
    {
        var doc = CreateDocumentWithText("Test");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "startParagraphIndex", startPara },
            { "startCharIndex", 0 },
            { "endParagraphIndex", endPara },
            { "endCharIndex", 4 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion

    #region Basic Delete Range Operations

    [Fact]
    public void Execute_DeletesTextRange()
    {
        var doc = CreateDocumentWithText("Hello World");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "startParagraphIndex", 0 },
            { "startCharIndex", 0 },
            { "endParagraphIndex", 0 },
            { "endCharIndex", 5 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("deleted", result.Message, StringComparison.OrdinalIgnoreCase);
        AssertContainsText(doc, "World");
        AssertModified(context);
    }

    [Theory]
    [InlineData(0, 5)]
    [InlineData(2, 8)]
    [InlineData(0, 11)]
    public void Execute_WithVariousRanges_DeletesCorrectly(int startChar, int endChar)
    {
        var doc = CreateDocumentWithText("Hello World");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "startParagraphIndex", 0 },
            { "startCharIndex", startChar },
            { "endParagraphIndex", 0 },
            { "endCharIndex", endChar }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("deleted", result.Message, StringComparison.OrdinalIgnoreCase);
        AssertModified(context);
    }

    #endregion

    #region Error Handling - Missing Parameters

    [Fact]
    public void Execute_WithoutStartParagraphIndex_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithText("Test");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "startCharIndex", 0 },
            { "endParagraphIndex", 0 },
            { "endCharIndex", 4 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("startParagraphIndex", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Execute_WithoutStartCharIndex_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithText("Test");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "startParagraphIndex", 0 },
            { "endParagraphIndex", 0 },
            { "endCharIndex", 4 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("startCharIndex", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Execute_WithoutEndParagraphIndex_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithText("Test");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "startParagraphIndex", 0 },
            { "startCharIndex", 0 },
            { "endCharIndex", 4 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("endParagraphIndex", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Execute_WithoutEndCharIndex_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithText("Test");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "startParagraphIndex", 0 },
            { "startCharIndex", 0 },
            { "endParagraphIndex", 0 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("endCharIndex", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    #endregion
}
