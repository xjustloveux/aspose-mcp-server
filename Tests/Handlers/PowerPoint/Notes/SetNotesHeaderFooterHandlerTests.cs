using System.Runtime.Versioning;
using AsposeMcpServer.Handlers.PowerPoint.Notes;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.Notes;

[SupportedOSPlatform("windows")]
public class SetNotesHeaderFooterHandlerTests : PptHandlerTestBase
{
    private readonly SetNotesHeaderFooterHandler _handler = new();

    #region Operation Property

    [SkippableFact]
    public void Operation_Returns_SetHeaderFooter()
    {
        SkipIfNotWindows();
        Assert.Equal("set_header_footer", _handler.Operation);
    }

    #endregion

    #region Basic Set Notes Header Footer Operations

    [SkippableFact]
    public void Execute_SetsHeaderText()
    {
        SkipIfNotWindows();
        var presentation = CreateEmptyPresentation();
        var context = CreateContext(presentation);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "headerText", "Test Header" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        if (!IsEvaluationMode())
        {
            var notesMaster = presentation.MasterNotesSlideManager.MasterNotesSlide;
            Assert.NotNull(notesMaster);
            Assert.True(notesMaster.HeaderFooterManager.IsHeaderVisible);
        }

        AssertModified(context);
    }

    [SkippableFact]
    public void Execute_SetsFooterText()
    {
        SkipIfNotWindows();
        var presentation = CreateEmptyPresentation();
        var context = CreateContext(presentation);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "footerText", "Test Footer" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        if (!IsEvaluationMode())
        {
            var notesMaster = presentation.MasterNotesSlideManager.MasterNotesSlide;
            Assert.NotNull(notesMaster);
            Assert.True(notesMaster.HeaderFooterManager.IsFooterVisible);
        }

        AssertModified(context);
    }

    [SkippableFact]
    public void Execute_SetsDateText()
    {
        SkipIfNotWindows();
        var presentation = CreateEmptyPresentation();
        var context = CreateContext(presentation);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "dateText", "2026-01-11" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        if (!IsEvaluationMode())
        {
            var notesMaster = presentation.MasterNotesSlideManager.MasterNotesSlide;
            Assert.NotNull(notesMaster);
            Assert.True(notesMaster.HeaderFooterManager.IsDateTimeVisible);
        }

        AssertModified(context);
    }

    [SkippableFact]
    public void Execute_SetsPageNumberVisibility()
    {
        SkipIfNotWindows();
        var presentation = CreateEmptyPresentation();
        var context = CreateContext(presentation);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "showPageNumber", false }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        if (!IsEvaluationMode())
        {
            var notesMaster = presentation.MasterNotesSlideManager.MasterNotesSlide;
            Assert.NotNull(notesMaster);
            Assert.False(notesMaster.HeaderFooterManager.IsSlideNumberVisible);
        }

        AssertModified(context);
    }

    [SkippableFact]
    public void Execute_WithAllSettings()
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

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        if (!IsEvaluationMode())
        {
            var notesMaster = presentation.MasterNotesSlideManager.MasterNotesSlide;
            Assert.NotNull(notesMaster);
            var manager = notesMaster.HeaderFooterManager;
            Assert.True(manager.IsHeaderVisible);
            Assert.True(manager.IsFooterVisible);
            Assert.True(manager.IsDateTimeVisible);
            Assert.True(manager.IsSlideNumberVisible);
        }

        AssertModified(context);
    }

    #endregion
}
