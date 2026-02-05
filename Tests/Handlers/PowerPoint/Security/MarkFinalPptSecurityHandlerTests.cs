using AsposeMcpServer.Handlers.PowerPoint.Security;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.Security;

public class MarkFinalPptSecurityHandlerTests : PptHandlerTestBase
{
    private readonly MarkFinalPptSecurityHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_MarkFinal()
    {
        Assert.Equal("mark_final", _handler.Operation);
    }

    #endregion

    #region Mark As Final

    [Fact]
    public void Execute_MarkAsFinal_SetsProperty()
    {
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

    [Fact]
    public void Execute_MarkAsFinal_ReturnsMarkedMessage()
    {
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

    [Fact]
    public void Execute_DefaultMarkAsFinal_MarksAsFinal()
    {
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        Assert.True((bool)pres.DocumentProperties["_MarkAsFinal"]);
    }

    #endregion

    #region Unmark As Final

    [Fact]
    public void Execute_UnmarkAsFinal_UnsetsProperty()
    {
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

    [Fact]
    public void Execute_UnmarkAsFinal_ReturnsUnmarkedMessage()
    {
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
