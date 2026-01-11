using AsposeMcpServer.Handlers.Word.Paragraph;
using AsposeMcpServer.Tests.Helpers;

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

        var result = _handler.Execute(context, parameters);

        Assert.Contains("0", result);
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

        var result = _handler.Execute(context, parameters);

        Assert.Contains("edited", result, StringComparison.OrdinalIgnoreCase);
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

        var result = _handler.Execute(context, parameters);

        Assert.Contains("text content updated", result, StringComparison.OrdinalIgnoreCase);
        AssertContainsText(doc, "Updated text");
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

        var result = _handler.Execute(context, parameters);

        Assert.Contains(index.ToString(), result);
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

    #endregion
}
