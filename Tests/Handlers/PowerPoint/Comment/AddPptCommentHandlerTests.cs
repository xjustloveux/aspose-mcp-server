using AsposeMcpServer.Handlers.PowerPoint.Comment;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.Comment;

public class AddPptCommentHandlerTests : PptHandlerTestBase
{
    private readonly AddPptCommentHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Add()
    {
        Assert.Equal("add", _handler.Operation);
    }

    #endregion

    #region Various Slide Indices

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    public void Execute_AddsCommentAtVariousSlideIndices(int slideIndex)
    {
        var pres = CreatePresentationWithSlides(3);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", $"Comment on slide {slideIndex}" },
            { "author", "Author" },
            { "slideIndex", slideIndex }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
    }

    #endregion

    #region Basic Add Operations

    [Fact]
    public void Execute_AddsComment()
    {
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "Test comment" },
            { "author", "Test Author" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        Assert.Contains("Comment added", ((SuccessResult)res).Message);
        AssertModified(context);
    }

    [Fact]
    public void Execute_ReturnsAuthorInMessage()
    {
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
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
    public void Execute_ReturnsSlideIndexInMessage()
    {
        var pres = CreatePresentationWithSlides(3);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "Comment" },
            { "author", "Author" },
            { "slideIndex", 1 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);
        Assert.Contains("slide 1", result.Message);
    }

    [Fact]
    public void Execute_DefaultSlideIndex_AddsToFirstSlide()
    {
        var pres = CreatePresentationWithSlides(3);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "Default slide comment" },
            { "author", "Author" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        var comments = pres.Slides[0].GetSlideComments(null);
        Assert.True(comments.Length > 0);
    }

    #endregion

    #region Position Parameters

    [Fact]
    public void Execute_WithPosition_ShouldSetPosition()
    {
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "Positioned comment" },
            { "author", "Author" },
            { "x", 100f },
            { "y", 200f }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
    }

    [Fact]
    public void Execute_DefaultPosition_UsesZeroZero()
    {
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "Comment" },
            { "author", "Author" }
        });

        _handler.Execute(context, parameters);

        var comments = pres.Slides[0].GetSlideComments(null);
        Assert.Equal(0f, comments[0].Position.X);
        Assert.Equal(0f, comments[0].Position.Y);
    }

    #endregion

    #region Author Reuse

    [Fact]
    public void Execute_SameAuthor_ShouldReuse()
    {
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);

        var p1 = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "First" },
            { "author", "Same Author" }
        });
        _handler.Execute(context, p1);

        var p2 = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "Second" },
            { "author", "Same Author" }
        });
        _handler.Execute(context, p2);

        Assert.Single(pres.CommentAuthors);
    }

    [Fact]
    public void Execute_DifferentAuthors_CreatesMultiple()
    {
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);

        var p1 = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "First" },
            { "author", "Author A" }
        });
        _handler.Execute(context, p1);

        var p2 = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "Second" },
            { "author", "Author B" }
        });
        _handler.Execute(context, p2);

        Assert.Equal(2, pres.CommentAuthors.Count);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_MissingText_ThrowsArgumentException()
    {
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "author", "Author" }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_MissingAuthor_ThrowsArgumentException()
    {
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "Comment" }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_InvalidSlideIndex_ThrowsArgumentException()
    {
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "Comment" },
            { "author", "Author" },
            { "slideIndex", 99 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_NegativeSlideIndex_ThrowsArgumentException()
    {
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "Comment" },
            { "author", "Author" },
            { "slideIndex", -1 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion
}
