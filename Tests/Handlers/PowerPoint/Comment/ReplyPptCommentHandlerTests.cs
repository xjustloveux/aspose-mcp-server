using System.Drawing;
using Aspose.Slides;
using AsposeMcpServer.Handlers.PowerPoint.Comment;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.Comment;

public class ReplyPptCommentHandlerTests : PptHandlerTestBase
{
    private readonly ReplyPptCommentHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Reply()
    {
        Assert.Equal("reply", _handler.Operation);
    }

    #endregion

    #region Helper Methods

    private static Presentation CreatePresentationWithParentComment(string text = "Parent comment",
        string authorName = "Original Author")
    {
        var pres = new Presentation();
        var author = pres.CommentAuthors.AddAuthor(authorName, "OA");
        author.Comments.AddComment(text, pres.Slides[0], new PointF(0, 0), DateTime.UtcNow);
        return pres;
    }

    #endregion

    #region Basic Reply Operations

    [Fact]
    public void Execute_ValidReply_AddsReply()
    {
        var pres = CreatePresentationWithParentComment();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "commentIndex", 0 },
            { "text", "Reply text" },
            { "author", "Replier" },
            { "slideIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        AssertModified(context);
    }

    [Fact]
    public void Execute_ReturnsCommentIndexInMessage()
    {
        var pres = CreatePresentationWithParentComment();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "commentIndex", 0 },
            { "text", "Reply" },
            { "author", "Author" },
            { "slideIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);
        Assert.Contains("comment 0", result.Message);
    }

    [Fact]
    public void Execute_ReturnsAuthorInMessage()
    {
        var pres = CreatePresentationWithParentComment();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "commentIndex", 0 },
            { "text", "Reply" },
            { "author", "Jane Doe" },
            { "slideIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);
        Assert.Contains("Jane Doe", result.Message);
    }

    [Fact]
    public void Execute_ReturnsSlideIndexInMessage()
    {
        var pres = CreatePresentationWithParentComment();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "commentIndex", 0 },
            { "text", "Reply" },
            { "author", "Author" },
            { "slideIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);
        Assert.Contains("slide 0", result.Message);
    }

    #endregion

    #region Reply Author Handling

    [Fact]
    public void Execute_SameAuthorAsParent_ReusesAuthor()
    {
        var pres = CreatePresentationWithParentComment("Parent", "Same Author");
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "commentIndex", 0 },
            { "text", "Reply" },
            { "author", "Same Author" },
            { "slideIndex", 0 }
        });

        _handler.Execute(context, parameters);

        Assert.Single(pres.CommentAuthors);
    }

    [Fact]
    public void Execute_DifferentAuthor_CreatesNewAuthor()
    {
        var pres = CreatePresentationWithParentComment("Parent", "Original");
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "commentIndex", 0 },
            { "text", "Reply" },
            { "author", "New Author" },
            { "slideIndex", 0 }
        });

        _handler.Execute(context, parameters);

        Assert.Equal(2, pres.CommentAuthors.Count);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_InvalidCommentIndex_ThrowsArgumentException()
    {
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "commentIndex", 0 },
            { "text", "Reply" },
            { "author", "Author" },
            { "slideIndex", 0 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_NegativeCommentIndex_ThrowsArgumentException()
    {
        var pres = CreatePresentationWithParentComment();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "commentIndex", -1 },
            { "text", "Reply" },
            { "author", "Author" },
            { "slideIndex", 0 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_OutOfRangeCommentIndex_ThrowsArgumentException()
    {
        var pres = CreatePresentationWithParentComment();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "commentIndex", 99 },
            { "text", "Reply" },
            { "author", "Author" },
            { "slideIndex", 0 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_MissingText_ThrowsArgumentException()
    {
        var pres = CreatePresentationWithParentComment();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "commentIndex", 0 },
            { "author", "Author" },
            { "slideIndex", 0 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_MissingAuthor_ThrowsArgumentException()
    {
        var pres = CreatePresentationWithParentComment();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "commentIndex", 0 },
            { "text", "Reply" },
            { "slideIndex", 0 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_MissingCommentIndex_ThrowsArgumentException()
    {
        var pres = CreatePresentationWithParentComment();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "Reply" },
            { "author", "Author" },
            { "slideIndex", 0 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion
}
