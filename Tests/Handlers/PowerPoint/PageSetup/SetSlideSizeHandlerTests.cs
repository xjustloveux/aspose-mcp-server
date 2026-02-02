using Aspose.Slides;
using AsposeMcpServer.Handlers.PowerPoint.PageSetup;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

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

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        if (!IsEvaluationMode()) Assert.Equal(SlideSizeType.OnScreen16x9, presentation.SlideSize.Type);
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

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        if (!IsEvaluationMode()) Assert.Equal(SlideSizeType.OnScreen16x10, presentation.SlideSize.Type);
        AssertModified(context);
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

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        if (!IsEvaluationMode()) Assert.Equal(SlideSizeType.A4Paper, presentation.SlideSize.Type);
        AssertModified(context);
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

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        if (!IsEvaluationMode())
        {
            Assert.Equal(SlideSizeType.Custom, presentation.SlideSize.Type);
            Assert.Equal(800f, presentation.SlideSize.Size.Width);
            Assert.Equal(600f, presentation.SlideSize.Size.Height);
        }

        AssertModified(context);
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

    #region Preset Variations

    [Fact]
    public void Execute_SetsWidescreenSize()
    {
        var presentation = CreateEmptyPresentation();
        var context = CreateContext(presentation);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "preset", "Widescreen" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        if (!IsEvaluationMode()) Assert.Equal(SlideSizeType.Widescreen, presentation.SlideSize.Type);
        AssertModified(context);
    }

    [Fact]
    public void Execute_SetsLetterSize()
    {
        var presentation = CreateEmptyPresentation();
        var context = CreateContext(presentation);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "preset", "Letter" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        if (!IsEvaluationMode()) Assert.Equal(SlideSizeType.LetterPaper, presentation.SlideSize.Type);
        AssertModified(context);
    }

    [Fact]
    public void Execute_SetsBannerSize()
    {
        var presentation = CreateEmptyPresentation();
        var context = CreateContext(presentation);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "preset", "Banner" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        if (!IsEvaluationMode()) Assert.Equal(SlideSizeType.Banner, presentation.SlideSize.Type);
        AssertModified(context);
    }

    #endregion

    #region Validation Errors

    [Fact]
    public void Execute_WithUnsupportedPreset_ThrowsArgumentException()
    {
        var presentation = CreateEmptyPresentation();
        var context = CreateContext(presentation);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "preset", "InvalidPreset" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("Unsupported preset", ex.Message);
    }

    [Fact]
    public void Execute_WithUnsupportedScaleType_ThrowsArgumentException()
    {
        var presentation = CreateEmptyPresentation();
        var context = CreateContext(presentation);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "preset", "OnScreen16x9" },
            { "scaleType", "InvalidScale" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("Unsupported scaleType", ex.Message);
    }

    [Fact]
    public void Execute_CustomWithZeroWidth_ThrowsArgumentException()
    {
        var presentation = CreateEmptyPresentation();
        var context = CreateContext(presentation);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "preset", "Custom" },
            { "width", 0.0 },
            { "height", 600.0 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("Width", ex.Message);
    }

    [Fact]
    public void Execute_CustomWithExcessiveWidth_ThrowsArgumentException()
    {
        var presentation = CreateEmptyPresentation();
        var context = CreateContext(presentation);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "preset", "Custom" },
            { "width", 6000.0 },
            { "height", 600.0 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("Width", ex.Message);
    }

    [Fact]
    public void Execute_CustomWithZeroHeight_ThrowsArgumentException()
    {
        var presentation = CreateEmptyPresentation();
        var context = CreateContext(presentation);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "preset", "Custom" },
            { "width", 800.0 },
            { "height", 0.0 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("Height", ex.Message);
    }

    [Fact]
    public void Execute_CustomWithExcessiveHeight_ThrowsArgumentException()
    {
        var presentation = CreateEmptyPresentation();
        var context = CreateContext(presentation);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "preset", "Custom" },
            { "width", 800.0 },
            { "height", 6000.0 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("Height", ex.Message);
    }

    #endregion

    #region Scale Type Variations

    [Fact]
    public void Execute_WithMaximizeScaleType_Succeeds()
    {
        var presentation = CreateEmptyPresentation();
        var context = CreateContext(presentation);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "preset", "A4" },
            { "scaleType", "Maximize" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
    }

    [Fact]
    public void Execute_WithDoNotScaleScaleType_Succeeds()
    {
        var presentation = CreateEmptyPresentation();
        var context = CreateContext(presentation);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "preset", "A4" },
            { "scaleType", "DoNotScale" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
    }

    #endregion
}
