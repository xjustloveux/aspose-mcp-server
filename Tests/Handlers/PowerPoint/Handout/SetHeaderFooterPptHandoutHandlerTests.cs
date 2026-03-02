using System.Runtime.Versioning;
using AsposeMcpServer.Handlers.PowerPoint.Handout;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.Handout;

[SupportedOSPlatform("windows")]
public class SetHeaderFooterPptHandoutHandlerTests : PptHandlerTestBase
{
    private readonly SetHeaderFooterPptHandoutHandler _handler = new();

    #region Operation Property

    [SkippableFact]
    public void Operation_Returns_SetHeaderFooter()
    {
        SkipIfNotWindows();
        Assert.Equal("set_header_footer", _handler.Operation);
    }

    #endregion

    #region Auto-Create Handout Master

    [SkippableFact]
    public void Execute_WithNoHandoutMaster_AutoCreatesAndSetsHeader()
    {
        SkipIfNotWindows();
        var presentation = CreateEmptyPresentation();
        var context = CreateContext(presentation);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "headerText", "Test Header" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(result);
        if (!IsEvaluationMode())
        {
            var handoutMaster = presentation.MasterHandoutSlideManager.MasterHandoutSlide;
            Assert.NotNull(handoutMaster);
            Assert.True(handoutMaster.HeaderFooterManager.IsHeaderVisible);
        }

        AssertModified(context);
    }

    [SkippableFact]
    public void Execute_WithNoHandoutMaster_AutoCreatesAndSetsFooter()
    {
        SkipIfNotWindows();
        var presentation = CreateEmptyPresentation();
        var context = CreateContext(presentation);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "footerText", "Test Footer" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(result);
        if (!IsEvaluationMode())
        {
            var handoutMaster = presentation.MasterHandoutSlideManager.MasterHandoutSlide;
            Assert.NotNull(handoutMaster);
            Assert.True(handoutMaster.HeaderFooterManager.IsFooterVisible);
        }

        AssertModified(context);
    }

    [SkippableFact]
    public void Execute_WithNoHandoutMaster_AutoCreatesAndSetsDate()
    {
        SkipIfNotWindows();
        var presentation = CreateEmptyPresentation();
        var context = CreateContext(presentation);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "dateText", "2026-01-11" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(result);
        if (!IsEvaluationMode())
        {
            var handoutMaster = presentation.MasterHandoutSlideManager.MasterHandoutSlide;
            Assert.NotNull(handoutMaster);
            Assert.True(handoutMaster.HeaderFooterManager.IsDateTimeVisible);
        }

        AssertModified(context);
    }

    [SkippableFact]
    public void Execute_WithNoHandoutMaster_AutoCreatesAndSetsAllSettings()
    {
        SkipIfNotWindows();
        var presentation = CreateEmptyPresentation();
        var context = CreateContext(presentation);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "headerText", "Header" },
            { "footerText", "Footer" },
            { "dateText", "Date" },
            { "showPageNumber", true }
        });

        var result = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(result);
        if (!IsEvaluationMode())
        {
            var handoutMaster = presentation.MasterHandoutSlideManager.MasterHandoutSlide;
            Assert.NotNull(handoutMaster);
            var manager = handoutMaster.HeaderFooterManager;
            Assert.True(manager.IsHeaderVisible);
            Assert.True(manager.IsFooterVisible);
            Assert.True(manager.IsDateTimeVisible);
            Assert.True(manager.IsSlideNumberVisible);
        }

        AssertModified(context);
    }

    #endregion
}
