using System.Runtime.Versioning;
using AsposeMcpServer.Handlers.PowerPoint.Security;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.Security;

[SupportedOSPlatform("windows")]
public class MarkFinalPptSecurityHandlerTests : PptHandlerTestBase
{
    private readonly MarkFinalPptSecurityHandler _handler = new();

    #region Operation Property

    [SkippableFact]
    public void Operation_Returns_MarkFinal()
    {
        SkipIfNotWindows();
        Assert.Equal("mark_final", _handler.Operation);
    }

    #endregion

    #region Mark As Final

    [SkippableFact]
    public void Execute_MarkAsFinal_SetsProperty()
    {
        SkipIfNotWindows();
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "markAsFinal", true }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        AssertModified(context);
        Assert.True((bool)pres.DocumentProperties["_MarkAsFinal"]);
    }

    [SkippableFact]
    public void Execute_MarkAsFinal_ReturnsMarkedMessage()
    {
        SkipIfNotWindows();
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "markAsFinal", true }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);
        Assert.Contains("marked as final", result.Message);
    }

    [SkippableFact]
    public void Execute_DefaultMarkAsFinal_MarksAsFinal()
    {
        SkipIfNotWindows();
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        Assert.True((bool)pres.DocumentProperties["_MarkAsFinal"]);
    }

    #endregion

    #region Unmark As Final

    [SkippableFact]
    public void Execute_UnmarkAsFinal_UnsetsProperty()
    {
        SkipIfNotWindows();
        var pres = CreateEmptyPresentation();
        pres.DocumentProperties["_MarkAsFinal"] = true;
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "markAsFinal", false }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        Assert.False((bool)pres.DocumentProperties["_MarkAsFinal"]);
    }

    [SkippableFact]
    public void Execute_UnmarkAsFinal_ReturnsUnmarkedMessage()
    {
        SkipIfNotWindows();
        var pres = CreateEmptyPresentation();
        pres.DocumentProperties["_MarkAsFinal"] = true;
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "markAsFinal", false }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);
        Assert.Contains("unmarked as final", result.Message);
    }

    #endregion
}
