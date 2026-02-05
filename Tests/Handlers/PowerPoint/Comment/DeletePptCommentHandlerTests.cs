using System.Drawing;
using Aspose.Slides;
using AsposeMcpServer.Handlers.PowerPoint.Comment;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.Comment;

public class DeletePptCommentHandlerTests : PptHandlerTestBase
{
    private readonly DeletePptCommentHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Delete()
    {
        Assert.Equal("delete", _handler.Operation);
    }

    #endregion

    #region Helper Methods

    private static Presentation CreatePresentationWithComments(int count)
    {
        var pres = new Presentation();
        for (var i = 0; i < count; i++)
        {
            var author = pres.CommentAuthors.AddAuthor($"Author {i}", $"A{i}");
            author.Comments.AddComment($"Comment {i}", pres.Slides[0], new PointF(0, 0), DateTime.UtcNow);
        }

        return pres;
    }

    #endregion

    #region Basic Delete Operations

    [Fact]
    public void Execute_ValidIndex_DeletesComment()
    {
        var pres = CreatePresentationWithComments(1);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "commentIndex", 0 },
            { "slideIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        AssertModified(context);
    }

    [Fact]
    public void Execute_ReturnsCommentIndexInMessage()
    {
        var pres = CreatePresentationWithComments(1);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "commentIndex", 0 },
            { "slideIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);
        Assert.Contains("Comment 0", result.Message);
    }

    [Fact]
    public void Execute_ReturnsSlideIndexInMessage()
    {
        var pres = CreatePresentationWithComments(1);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "commentIndex", 0 },
            { "slideIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);
        Assert.Contains("slide 0", result.Message);
    }

    #endregion

    #region Various Comment Indices

    [Fact]
    public void Execute_DeletesSecondComment()
    {
        var pres = CreatePresentationWithComments(3);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "commentIndex", 1 },
            { "slideIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
    }

    [Fact]
    public void Execute_DeletesLastComment()
    {
        var pres = CreatePresentationWithComments(3);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "commentIndex", 2 },
            { "slideIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_InvalidIndex_ThrowsArgumentException()
    {
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "commentIndex", 0 },
            { "slideIndex", 0 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_NegativeIndex_ThrowsArgumentException()
    {
        var pres = CreatePresentationWithComments(1);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "commentIndex", -1 },
            { "slideIndex", 0 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_OutOfRangeIndex_ThrowsArgumentException()
    {
        var pres = CreatePresentationWithComments(2);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "commentIndex", 99 },
            { "slideIndex", 0 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_MissingCommentIndex_ThrowsArgumentException()
    {
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion
}
