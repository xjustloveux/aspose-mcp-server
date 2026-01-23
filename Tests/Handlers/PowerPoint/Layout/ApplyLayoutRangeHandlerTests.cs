using AsposeMcpServer.Handlers.PowerPoint.Layout;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.Layout;

public class ApplyLayoutRangeHandlerTests : PptHandlerTestBase
{
    private readonly ApplyLayoutRangeHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_ApplyLayoutRange()
    {
        Assert.Equal("apply_layout_range", _handler.Operation);
    }

    #endregion

    #region Basic Apply Operations

    [Fact]
    public void Execute_AppliesLayoutToRange()
    {
        var pres = CreatePresentationWithSlides(3);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndices", "[0, 1]" },
            { "layout", "Title" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("applied", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("2", result.Message);
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithBlankLayout_AppliesBlankLayout()
    {
        var pres = CreatePresentationWithSlides(2);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndices", "[0]" },
            { "layout", "Blank" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("blank", result.Message, StringComparison.OrdinalIgnoreCase);
        AssertModified(context);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutSlideIndices_ThrowsArgumentException()
    {
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "layout", "Title" }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithoutLayout_ThrowsArgumentException()
    {
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndices", "[0]" }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithEmptySlideIndices_ThrowsArgumentException()
    {
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndices", "[]" },
            { "layout", "Title" }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithInvalidSlideIndex_ThrowsArgumentException()
    {
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndices", "[99]" },
            { "layout", "Title" }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion
}
