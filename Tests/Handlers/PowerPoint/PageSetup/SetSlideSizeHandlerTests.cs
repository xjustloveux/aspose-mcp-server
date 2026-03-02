using System.Runtime.Versioning;
using Aspose.Slides;
using AsposeMcpServer.Handlers.PowerPoint.PageSetup;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.PageSetup;

[SupportedOSPlatform("windows")]
public class SetSlideSizeHandlerTests : PptHandlerTestBase
{
    private readonly SetSlideSizeHandler _handler = new();

    #region Operation Property

    [SkippableFact]
    public void Operation_Returns_SetSize()
    {
        SkipIfNotWindows();
        Assert.Equal("set_size", _handler.Operation);
    }

    #endregion

    #region Basic Set Slide Size Operations

    [SkippableFact]
    public void Execute_SetsDefaultSize()
    {
        SkipIfNotWindows();
        var presentation = CreateEmptyPresentation();
        var context = CreateContext(presentation);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        if (!IsEvaluationMode()) Assert.Equal(SlideSizeType.OnScreen16x9, presentation.SlideSize.Type);
        AssertModified(context);
    }

    [SkippableFact]
    public void Execute_Sets16x10Size()
    {
        SkipIfNotWindows();
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

    [SkippableFact]
    public void Execute_SetsA4Size()
    {
        SkipIfNotWindows();
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

    [SkippableFact]
    public void Execute_SetsCustomSize()
    {
        SkipIfNotWindows();
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

    [SkippableFact]
    public void Execute_CustomWithoutWidth_ThrowsArgumentException()
    {
        SkipIfNotWindows();
        var presentation = CreateEmptyPresentation();
        var context = CreateContext(presentation);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "preset", "Custom" },
            { "height", 600.0 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [SkippableFact]
    public void Execute_CustomWithoutHeight_ThrowsArgumentException()
    {
        SkipIfNotWindows();
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

    [SkippableFact]
    public void Execute_SetsWidescreenSize()
    {
        SkipIfNotWindows();
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

    [SkippableFact]
    public void Execute_SetsLetterSize()
    {
        SkipIfNotWindows();
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

    [SkippableFact]
    public void Execute_SetsBannerSize()
    {
        SkipIfNotWindows();
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

    [SkippableFact]
    public void Execute_WithUnsupportedPreset_ThrowsArgumentException()
    {
        SkipIfNotWindows();
        var presentation = CreateEmptyPresentation();
        var context = CreateContext(presentation);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "preset", "InvalidPreset" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("Unsupported preset", ex.Message);
    }

    [SkippableFact]
    public void Execute_WithUnsupportedScaleType_ThrowsArgumentException()
    {
        SkipIfNotWindows();
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

    [SkippableFact]
    public void Execute_CustomWithZeroWidth_ThrowsArgumentException()
    {
        SkipIfNotWindows();
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

    [SkippableFact]
    public void Execute_CustomWithExcessiveWidth_ThrowsArgumentException()
    {
        SkipIfNotWindows();
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

    [SkippableFact]
    public void Execute_CustomWithZeroHeight_ThrowsArgumentException()
    {
        SkipIfNotWindows();
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

    [SkippableFact]
    public void Execute_CustomWithExcessiveHeight_ThrowsArgumentException()
    {
        SkipIfNotWindows();
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

    [SkippableFact]
    public void Execute_WithMaximizeScaleType_Succeeds()
    {
        SkipIfNotWindows();
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

    [SkippableFact]
    public void Execute_WithDoNotScaleScaleType_Succeeds()
    {
        SkipIfNotWindows();
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
