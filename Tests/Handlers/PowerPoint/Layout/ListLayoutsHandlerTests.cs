using System.Runtime.Versioning;
using AsposeMcpServer.Handlers.PowerPoint.Layout;
using AsposeMcpServer.Results.PowerPoint.Layout;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.Layout;

[SupportedOSPlatform("windows")]
public class ListLayoutsHandlerTests : PptHandlerTestBase
{
    private readonly ListLayoutsHandler _handler = new();

    #region Operation Property

    [SkippableFact]
    public void Operation_Returns_GetLayouts()
    {
        SkipIfNotWindows();
        Assert.Equal("list", _handler.Operation);
    }

    #endregion

    #region Error Handling

    [SkippableFact]
    public void Execute_WithInvalidMasterIndex_ThrowsArgumentException()
    {
        SkipIfNotWindows();
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "masterIndex", 99 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion

    #region Basic Get Operations

    [SkippableFact]
    public void Execute_ReturnsLayouts()
    {
        SkipIfNotWindows();
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetLayoutsResult>(res);

        Assert.NotNull(result.Masters);
        Assert.True(result.Masters.Count > 0);
        Assert.NotNull(result.Masters[0].Layouts);
        AssertNotModified(context);
    }

    [SkippableFact]
    public void Execute_WithMasterIndex_ReturnsLayoutsForMaster()
    {
        SkipIfNotWindows();
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "masterIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetLayoutsResult>(res);

        Assert.Equal(0, result.MasterIndex);
        Assert.NotNull(result.Layouts);
        AssertNotModified(context);
    }

    [SkippableFact]
    public void Execute_ReturnsResultType()
    {
        SkipIfNotWindows();
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetLayoutsResult>(res);

        Assert.NotNull(result);
        Assert.IsType<GetLayoutsResult>(result);
    }

    #endregion
}
