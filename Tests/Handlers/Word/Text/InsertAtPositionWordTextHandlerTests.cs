using Aspose.Words;
using AsposeMcpServer.Handlers.Word.Text;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Word.Text;

public class InsertAtPositionWordTextHandlerTests : WordHandlerTestBase
{
    private readonly InsertAtPositionWordTextHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_InsertAtPosition()
    {
        Assert.Equal("insert", _handler.Operation);
    }

    #endregion

    #region InsertBefore Option

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void Execute_WithInsertBefore_InsertsCorrectly(bool insertBefore)
    {
        var doc = CreateDocumentWithText("Hello World");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "insertParagraphIndex", 0 },
            { "charIndex", 0 },
            { "text", "Test" },
            { "insertBefore", insertBefore }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("inserted", result.Message, StringComparison.OrdinalIgnoreCase);
        AssertContainsText(doc, "Test");
        AssertModified(context);
    }

    #endregion

    #region Section Index

    [Fact]
    public void Execute_WithSectionIndex_InsertsInCorrectSection()
    {
        var doc = CreateDocumentWithText("Hello World");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "insertParagraphIndex", 0 },
            { "charIndex", 0 },
            { "text", "Test " },
            { "sectionIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("inserted", result.Message, StringComparison.OrdinalIgnoreCase);
        var text = GetDocumentText(doc);
        Assert.StartsWith("Test Hello", text);
        AssertModified(context);
    }

    #endregion

    #region Special Characters

    [Theory]
    [InlineData("中文測試")]
    [InlineData("日本語")]
    [InlineData("Special: !@#$%")]
    public void Execute_WithSpecialCharacters_InsertsCorrectly(string text)
    {
        var doc = CreateDocumentWithText("Hello World");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "insertParagraphIndex", 0 },
            { "charIndex", 5 },
            { "text", text }
        });

        _handler.Execute(context, parameters);

        AssertContainsText(doc, text);
        AssertModified(context);
    }

    #endregion

    #region Basic Insert Operations

    [Theory]
    [InlineData(0, 0, "Inserted")]
    [InlineData(0, 5, "Middle")]
    public void Execute_InsertsTextAtPosition(int paragraphIndex, int charIndex, string textToInsert)
    {
        var doc = CreateDocumentWithText("Hello World");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "insertParagraphIndex", paragraphIndex },
            { "charIndex", charIndex },
            { "text", textToInsert }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("inserted", result.Message, StringComparison.OrdinalIgnoreCase);
        AssertContainsText(doc, textToInsert);
        AssertModified(context);
    }

    [Fact]
    public void Execute_InsertsAtBeginning()
    {
        var doc = CreateDocumentWithText("World");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "insertParagraphIndex", 0 },
            { "charIndex", 0 },
            { "text", "Hello " }
        });

        _handler.Execute(context, parameters);

        var text = GetDocumentText(doc);
        Assert.StartsWith("Hello ", text);
        AssertContainsText(doc, "World");
        AssertModified(context);
    }

    [Fact]
    public void Execute_InsertsAtEnd()
    {
        var doc = CreateDocumentWithText("Hello");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "insertParagraphIndex", 0 },
            { "charIndex", 5 },
            { "text", " World" }
        });

        _handler.Execute(context, parameters);

        AssertContainsText(doc, "Hello World");
        AssertModified(context);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutParagraphIndex_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithText("Test");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "charIndex", 0 },
            { "text", "Insert" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("insertParagraphIndex", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Execute_WithoutCharIndex_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithText("Test");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "insertParagraphIndex", 0 },
            { "text", "Insert" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("charIndex", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Execute_WithoutText_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithText("Test");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "insertParagraphIndex", 0 },
            { "charIndex", 0 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("text", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData(100)]
    public void Execute_WithInvalidParagraphIndex_ThrowsArgumentException(int paragraphIndex)
    {
        var doc = CreateDocumentWithText("Test");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "insertParagraphIndex", paragraphIndex },
            { "charIndex", 0 },
            { "text", "Insert" }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion

    #region Global Paragraph Index (Issue #1 unification)

    [Fact]
    public void Execute_WithStoryTypeHeader_TargetsHeaderParagraph()
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Write("BodyOnly");
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.Write("HeaderOnly");

        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "insertParagraphIndex", 0 },
            { "storyType", "Header" },
            { "charIndex", 0 },
            { "text", "X" },
            { "insertBefore", true }
        });

        _handler.Execute(context, parameters);

        var headerPara = doc.FirstSection.HeadersFooters[HeaderFooterType.HeaderPrimary]
            .GetChildNodes(NodeType.Paragraph, true).Cast<Aspose.Words.Paragraph>().First();
        Assert.StartsWith("X", headerPara.GetText());
    }

    [Fact]
    public void Execute_WithParagraphIndexMinusOne_InsertsAtLastParagraph()
    {
        var doc = CreateDocumentWithParagraphs("First", "Last");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "insertParagraphIndex", -1 },
            { "charIndex", 0 },
            { "text", "X" },
            { "insertBefore", true }
        });

        _handler.Execute(context, parameters);

        var lastPara = doc.GetChildNodes(NodeType.Paragraph, true).Cast<Aspose.Words.Paragraph>().Last();
        Assert.StartsWith("X", lastPara.GetText());
    }

    [Fact]
    public void Execute_WithCharIndexBeyondText_AppendsAtParagraphEnd()
    {
        var doc = CreateDocumentWithText("Hello");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "insertParagraphIndex", 0 },
            { "charIndex", 999 },
            { "text", "X" },
            { "insertBefore", false }
        });

        _handler.Execute(context, parameters);

        Assert.StartsWith("HelloX", GetDocumentText(doc));
    }

    #endregion

    #region Field-Leading Paragraph (Issue #4)

    [Fact]
    public void Execute_AtCharZeroOfFieldLeadingParagraph_InsertsBeforeFieldNotIntoCode()
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.InsertField("MERGEFIELD Name");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "insertParagraphIndex", 0 },
            { "charIndex", 0 },
            { "text", "PREFIX " },
            { "insertBefore", true }
        });

        _handler.Execute(context, parameters);

        Assert.Equal(1, doc.Range.Fields.Count);
        Assert.Equal("MERGEFIELD Name", doc.Range.Fields[0].GetFieldCode().Trim());
        AssertContainsText(doc, "PREFIX");
    }

    [Fact]
    public void Execute_AfterFieldOnlyParagraph_InsertsAfterFieldNotIntoCode()
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.InsertField("MERGEFIELD Name");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "insertParagraphIndex", 0 },
            { "charIndex", 0 },
            { "text", "SUFFIX" },
            { "insertBefore", false }
        });

        _handler.Execute(context, parameters);

        Assert.Equal(1, doc.Range.Fields.Count);
        Assert.Equal("MERGEFIELD Name", doc.Range.Fields[0].GetFieldCode().Trim());
        AssertContainsText(doc, "SUFFIX");
    }

    [Fact]
    public void Execute_AfterFieldFollowedByText_InsertsAfterFieldNotIntoCode()
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.InsertField("MERGEFIELD Name");
        builder.Write(" Tail");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "insertParagraphIndex", 0 },
            { "charIndex", 0 },
            { "text", "MID" },
            { "insertBefore", false }
        });

        _handler.Execute(context, parameters);

        Assert.Equal(1, doc.Range.Fields.Count);
        Assert.Equal("MERGEFIELD Name", doc.Range.Fields[0].GetFieldCode().Trim());
        AssertContainsText(doc, "MID");
    }

    #endregion
}
