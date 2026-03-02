using System.Runtime.Versioning;
using Aspose.Slides.SmartArt;
using AsposeMcpServer.Handlers.PowerPoint.SmartArt;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.SmartArt;

[SupportedOSPlatform("windows")]
public class AddSmartArtHandlerTests : PptHandlerTestBase
{
    private readonly AddSmartArtHandler _handler = new();

    #region Operation Property

    [SkippableFact]
    public void Operation_Returns_Add()
    {
        SkipIfNotWindows();
        Assert.Equal("add", _handler.Operation);
    }

    #endregion

    #region Basic Add Operations

    [SkippableFact]
    public void Execute_AddsSmartArt()
    {
        SkipIfNotWindows();
        var pres = CreateEmptyPresentation();
        var initialShapeCount = pres.Slides[0].Shapes.Count;
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "layout", "BasicProcess" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        Assert.True(pres.Slides[0].Shapes.Count > initialShapeCount);
        if (!IsEvaluationMode())
            Assert.Contains(pres.Slides[0].Shapes, s => s is ISmartArt);
        AssertModified(context);
    }

    [SkippableFact]
    public void Execute_WithCustomPosition_AddsSmartArtAtPosition()
    {
        SkipIfNotWindows();
        var pres = CreateEmptyPresentation();
        var initialShapeCount = pres.Slides[0].Shapes.Count;
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "layout", "Hierarchy" },
            { "x", 50f },
            { "y", 50f },
            { "width", 300f },
            { "height", 250f }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        Assert.True(pres.Slides[0].Shapes.Count > initialShapeCount);
        if (!IsEvaluationMode())
            Assert.Contains(pres.Slides[0].Shapes, s => s is ISmartArt);
        AssertModified(context);
    }

    #endregion

    #region Error Handling

    [SkippableFact]
    public void Execute_WithoutSlideIndex_ThrowsArgumentException()
    {
        SkipIfNotWindows();
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "layout", "BasicProcess" }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [SkippableFact]
    public void Execute_WithoutLayout_ThrowsArgumentException()
    {
        SkipIfNotWindows();
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [SkippableFact]
    public void Execute_WithInvalidLayout_ThrowsArgumentException()
    {
        SkipIfNotWindows();
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "layout", "InvalidLayout" }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [SkippableFact]
    public void Execute_WithInvalidSlideIndex_ThrowsArgumentException()
    {
        SkipIfNotWindows();
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 99 },
            { "layout", "BasicProcess" }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion
}
