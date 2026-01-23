using AsposeMcpServer.Handlers.Word.List;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Word.List;

public class EditWordListItemHandlerTests : WordHandlerTestBase
{
    private readonly EditWordListItemHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_EditItem()
    {
        Assert.Equal("edit_item", _handler.Operation);
    }

    #endregion

    #region Various Paragraph Indices

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    public void Execute_EditsVariousParagraphs(int index)
    {
        var doc = CreateDocumentWithParagraphs("Item 0", "Item 1", "Item 2");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "paragraphIndex", index },
            { "text", $"Updated {index}" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("edited successfully", result.Message);
        Assert.Contains($"Paragraph index: {index}", result.Message);
    }

    #endregion

    #region Basic Edit Operations

    [Fact]
    public void Execute_EditsListItem()
    {
        var doc = CreateDocumentWithParagraphs("Item 1", "Item 2", "Item 3");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "paragraphIndex", 1 },
            { "text", "Updated Item" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("List item edited successfully", result.Message);
        AssertModified(context);
    }

    [Fact]
    public void Execute_ReturnsParagraphIndex()
    {
        var doc = CreateDocumentWithParagraphs("Item 1", "Item 2");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "paragraphIndex", 0 },
            { "text", "New Text" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("Paragraph index: 0", result.Message);
    }

    [Fact]
    public void Execute_ReturnsNewText()
    {
        var doc = CreateDocumentWithParagraphs("Item 1");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "paragraphIndex", 0 },
            { "text", "Updated Content" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("New text: Updated Content", result.Message);
    }

    [SkippableFact]
    public void Execute_UpdatesDocumentContent()
    {
        SkipInEvaluationMode(AsposeLibraryType.Words, "Evaluation mode adds watermark to text");
        var doc = CreateDocumentWithParagraphs("Original Text");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "paragraphIndex", 0 },
            { "text", "Modified Text" }
        });

        _handler.Execute(context, parameters);

        AssertContainsText(doc, "Modified Text");
    }

    #endregion

    #region Level Parameter

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    public void Execute_WithLevel_ReturnsLevelInfo(int level)
    {
        var doc = CreateDocumentWithParagraphs("Item 1");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "paragraphIndex", 0 },
            { "text", "Updated" },
            { "level", level }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains($"Level: {level}", result.Message);
    }

    [Fact]
    public void Execute_WithoutLevel_DoesNotShowLevelInfo()
    {
        var doc = CreateDocumentWithParagraphs("Item 1");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "paragraphIndex", 0 },
            { "text", "Updated" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.DoesNotContain("Level:", result.Message);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutParagraphIndex_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithParagraphs("Item 1");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "Updated" }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithoutText_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithParagraphs("Item 1");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "paragraphIndex", 0 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithEmptyText_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithParagraphs("Item 1");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "paragraphIndex", 0 },
            { "text", "" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("text", ex.Message);
    }

    [Fact]
    public void Execute_WithInvalidParagraphIndex_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithParagraphs("Item 1");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "paragraphIndex", 99 },
            { "text", "Updated" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("out of range", ex.Message);
    }

    [Fact]
    public void Execute_WithNegativeParagraphIndex_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithParagraphs("Item 1");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "paragraphIndex", -1 },
            { "text", "Updated" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("out of range", ex.Message);
    }

    #endregion
}
