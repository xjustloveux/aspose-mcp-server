using AsposeMcpServer.Handlers.PowerPoint.PageSetup;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.PageSetup;

public class SetSlideSizeHandlerTests : PptHandlerTestBase
{
    private readonly SetSlideSizeHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_SetSize()
    {
        Assert.Equal("set_size", _handler.Operation);
    }

    #endregion

    #region Basic Set Slide Size Operations

    [Fact]
    public void Execute_SetsDefaultSize()
    {
        var presentation = CreateEmptyPresentation();
        var context = CreateContext(presentation);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("slide size set", result.ToLower());
        AssertModified(context);
    }

    [Fact]
    public void Execute_Sets16x10Size()
    {
        var presentation = CreateEmptyPresentation();
        var context = CreateContext(presentation);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "preset", "OnScreen16x10" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("slide size set", result.ToLower());
    }

    [Fact]
    public void Execute_SetsA4Size()
    {
        var presentation = CreateEmptyPresentation();
        var context = CreateContext(presentation);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "preset", "A4" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("slide size set", result.ToLower());
    }

    [Fact]
    public void Execute_SetsCustomSize()
    {
        var presentation = CreateEmptyPresentation();
        var context = CreateContext(presentation);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "preset", "Custom" },
            { "width", 800.0 },
            { "height", 600.0 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("custom", result.ToLower());
    }

    [Fact]
    public void Execute_CustomWithoutWidth_ThrowsArgumentException()
    {
        var presentation = CreateEmptyPresentation();
        var context = CreateContext(presentation);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "preset", "Custom" },
            { "height", 600.0 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_CustomWithoutHeight_ThrowsArgumentException()
    {
        var presentation = CreateEmptyPresentation();
        var context = CreateContext(presentation);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "preset", "Custom" },
            { "width", 800.0 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion
}
