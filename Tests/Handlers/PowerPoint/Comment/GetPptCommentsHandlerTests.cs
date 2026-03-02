using System.Drawing;
using System.Runtime.Versioning;
using Aspose.Slides;
using AsposeMcpServer.Handlers.PowerPoint.Comment;
using AsposeMcpServer.Results.PowerPoint.Comment;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.Comment;

[SupportedOSPlatform("windows")]
public class GetPptCommentsHandlerTests : PptHandlerTestBase
{
    private readonly GetPptCommentsHandler _handler = new();

    #region Operation Property

    [SkippableFact]
    public void Operation_Returns_Get()
    {
        SkipIfNotWindows();
        Assert.Equal("get", _handler.Operation);
    }

    #endregion

    #region Modification State

    [SkippableFact]
    public void Execute_ShouldNotMarkModified()
    {
        SkipIfNotWindows();
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 }
        });

        _handler.Execute(context, parameters);

        AssertNotModified(context);
    }

    #endregion

    #region Default Slide Index

    [SkippableFact]
    public void Execute_DefaultSlideIndex_QueriesFirstSlide()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithComment();
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetCommentsPptResult>(res);
        Assert.Equal(0, result.SlideIndex);
        Assert.Equal(1, result.Count);
    }

    #endregion

    #region Helper Methods

    private static Presentation CreatePresentationWithComment(string text = "Test comment",
        string authorName = "Test Author")
    {
        var pres = new Presentation();
        var author = pres.CommentAuthors.AddAuthor(authorName, "TA");
        author.Comments.AddComment(text, pres.Slides[0], new PointF(0, 0), DateTime.UtcNow);
        return pres;
    }

    private static Presentation CreatePresentationWithMultipleComments(int count)
    {
        var pres = new Presentation();
        for (var i = 0; i < count; i++)
        {
            var author = pres.CommentAuthors.AddAuthor($"Author {i}", $"A{i}");
            author.Comments.AddComment($"Comment {i}", pres.Slides[0], new PointF(i * 10, i * 10), DateTime.UtcNow);
        }

        return pres;
    }

    #endregion

    #region Get Comments - Empty

    [SkippableFact]
    public void Execute_NoComments_ReturnsEmptyResult()
    {
        SkipIfNotWindows();
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetCommentsPptResult>(res);
        Assert.Equal(0, result.Count);
        Assert.Empty(result.Items);
    }

    [SkippableFact]
    public void Execute_NoComments_ReturnsMessage()
    {
        SkipIfNotWindows();
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetCommentsPptResult>(res);
        Assert.Contains("0 comment", result.Message);
    }

    #endregion

    #region Get Comments - With Data

    [SkippableFact]
    public void Execute_WithComment_ReturnsCount()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithComment();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetCommentsPptResult>(res);
        Assert.Equal(1, result.Count);
    }

    [SkippableFact]
    public void Execute_WithComment_ReturnsAuthor()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithComment("Hello");
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetCommentsPptResult>(res);
        Assert.Equal("Test Author", result.Items[0].Author);
    }

    [SkippableFact]
    public void Execute_WithComment_ReturnsText()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithComment("Hello World");
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetCommentsPptResult>(res);
        Assert.Equal("Hello World", result.Items[0].Text);
    }

    [SkippableFact]
    public void Execute_WithMultipleComments_ReturnsAll()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithMultipleComments(3);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetCommentsPptResult>(res);
        Assert.Equal(3, result.Count);
        Assert.Equal(3, result.Items.Count);
    }

    [SkippableFact]
    public void Execute_ReturnsSlideIndex()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithComment();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetCommentsPptResult>(res);
        Assert.Equal(0, result.SlideIndex);
    }

    [SkippableFact]
    public void Execute_ReturnsCommentIndex()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithMultipleComments(2);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetCommentsPptResult>(res);
        Assert.Equal(0, result.Items[0].Index);
        Assert.Equal(1, result.Items[1].Index);
    }

    #endregion
}
