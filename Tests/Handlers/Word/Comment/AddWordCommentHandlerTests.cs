using Aspose.Words;
using AsposeMcpServer.Handlers.Word.Comment;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Word.Comment;

public class AddWordCommentHandlerTests : WordHandlerTestBase
{
    private readonly AddWordCommentHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Add()
    {
        Assert.Equal("add", _handler.Operation);
    }

    #endregion

    #region Helper Methods

    private static Document CreateDocumentWithParagraphs(int count)
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        for (var i = 0; i < count; i++) builder.Writeln($"Paragraph {i}");
        return doc;
    }

    #endregion

    #region Without Paragraph Index

    [Fact]
    public void Execute_WithoutParagraphIndex_AddsAtDocumentEnd()
    {
        var doc = CreateDocumentWithParagraphs(3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "Comment at end" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("Comment added successfully", result.Message);
    }

    #endregion

    #region Basic Add Operations

    [Fact]
    public void Execute_AddsComment()
    {
        var doc = CreateDocumentWithText("Sample text for comment");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "This is a comment" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("Comment added successfully", result.Message);
        AssertModified(context);
    }

    [Fact]
    public void Execute_ReturnsCommentText()
    {
        var doc = CreateDocumentWithText("Sample text");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "My comment content" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("My comment content", result.Message);
    }

    [Fact]
    public void Execute_ReturnsAuthor()
    {
        var doc = CreateDocumentWithText("Sample text");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "Comment text" },
            { "author", "John Doe" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("John Doe", result.Message);
    }

    [Fact]
    public void Execute_DefaultsAuthor()
    {
        var doc = CreateDocumentWithText("Sample text");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "Comment text" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("Author:", result.Message);
    }

    #endregion

    #region Paragraph Index

    [Fact]
    public void Execute_WithParagraphIndex_AddsToSpecificParagraph()
    {
        var doc = CreateDocumentWithParagraphs(3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "Comment on second paragraph" },
            { "paragraphIndex", 1 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("Comment added successfully", result.Message);
    }

    [Fact]
    public void Execute_WithParagraphIndexMinusOne_AddsToLastParagraph()
    {
        var doc = CreateDocumentWithParagraphs(3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "Comment on last paragraph" },
            { "paragraphIndex", -1 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("Comment added successfully", result.Message);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutText_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithText("Sample text");
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("text", ex.Message);
    }

    [Fact]
    public void Execute_WithEmptyText_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithText("Sample text");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("text", ex.Message);
    }

    [Fact]
    public void Execute_WithInvalidParagraphIndex_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithParagraphs(3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "Comment text" },
            { "paragraphIndex", 99 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("out of range", ex.Message);
    }

    #endregion

    #region Author Initial

    [Fact]
    public void Execute_WithAuthorInitial_UsesProvidedInitial()
    {
        var doc = CreateDocumentWithText("Sample text");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "Comment text" },
            { "author", "John Doe" },
            { "authorInitial", "JD" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("Comment added successfully", result.Message);
    }

    [Fact]
    public void Execute_WithShortAuthor_CalculatesInitial()
    {
        var doc = CreateDocumentWithText("Sample text");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "Comment text" },
            { "author", "A" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("Comment added successfully", result.Message);
    }

    #endregion

    #region Run Indices

    [Fact]
    public void Execute_WithStartRunIndex_AddsToSpecificRun()
    {
        var doc = CreateDocumentWithParagraphs(1);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "Comment on run" },
            { "paragraphIndex", 0 },
            { "startRunIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("Comment added successfully", result.Message);
    }

    [Fact]
    public void Execute_WithStartAndEndRunIndex_AddsToRunRange()
    {
        var doc = CreateDocumentWithParagraphs(1);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "Comment on run range" },
            { "paragraphIndex", 0 },
            { "startRunIndex", 0 },
            { "endRunIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("Comment added successfully", result.Message);
    }

    [Fact]
    public void Execute_WithInvalidStartRunIndex_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithParagraphs(1);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "Comment text" },
            { "paragraphIndex", 0 },
            { "startRunIndex", 99 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("out of range", ex.Message);
    }

    [Fact]
    public void Execute_WithInvalidStartEndRunIndices_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithParagraphs(1);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "Comment text" },
            { "paragraphIndex", 0 },
            { "startRunIndex", 99 },
            { "endRunIndex", 99 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("out of range", ex.Message);
    }

    #endregion
}
