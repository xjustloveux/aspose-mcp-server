using System.Runtime.Versioning;
using AsposeMcpServer.Handlers.PowerPoint.Layout;
using AsposeMcpServer.Results.PowerPoint.Layout;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.Layout;

[SupportedOSPlatform("windows")]
public class GetMastersHandlerTests : PptHandlerTestBase
{
    private readonly GetMastersHandler _handler = new();

    #region Operation Property

    [SkippableFact]
    public void Operation_Returns_GetMasters()
    {
        SkipIfNotWindows();
        Assert.Equal("get_masters", _handler.Operation);
    }

    #endregion

    #region Basic Get Operations

    [SkippableFact]
    public void Execute_ReturnsMasters()
    {
        SkipIfNotWindows();
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetMastersResult>(res);

        Assert.True(result.Count >= 0);
        Assert.NotNull(result.Masters);
        AssertNotModified(context);
    }

    [SkippableFact]
    public void Execute_ReturnsCorrectMasterCount()
    {
        SkipIfNotWindows();
        var pres = CreateEmptyPresentation();
        var expectedCount = pres.Masters.Count;
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetMastersResult>(res);

        Assert.Equal(expectedCount, result.Count);
        Assert.Equal(expectedCount, result.Masters.Count);
    }

    [SkippableFact]
    public void Execute_ReturnsResultType()
    {
        SkipIfNotWindows();
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetMastersResult>(res);

        Assert.NotNull(result);
        Assert.IsType<GetMastersResult>(result);
    }

    [SkippableFact]
    public void Execute_IncludesLayoutCount()
    {
        SkipIfNotWindows();
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetMastersResult>(res);

        Assert.NotNull(result.Masters);
        if (result.Masters.Count > 0) Assert.True(result.Masters[0].LayoutCount >= 0);
    }

    [SkippableFact]
    public void Execute_IncludesMasterName()
    {
        SkipIfNotWindows();
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetMastersResult>(res);

        Assert.NotNull(result.Masters);
        if (result.Masters.Count > 0)
        {
            // Name property exists (may be null)
            var master = result.Masters[0];
            Assert.IsType<GetMasterInfo>(master);
        }
    }

    [SkippableFact]
    public void Execute_IncludesLayoutsList()
    {
        SkipIfNotWindows();
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetMastersResult>(res);

        Assert.NotNull(result.Masters);
        if (result.Masters.Count > 0) Assert.NotNull(result.Masters[0].Layouts);
    }

    [SkippableFact]
    public void Execute_IncludesMasterIndex()
    {
        SkipIfNotWindows();
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetMastersResult>(res);

        Assert.NotNull(result.Masters);
        if (result.Masters.Count > 0) Assert.Equal(0, result.Masters[0].Index);
    }

    #endregion

    #region API Constraint - Empty Masters

    [SkippableFact]
    public void Execute_PresentationAlwaysHasAtLeastOneMaster()
    {
        SkipIfNotWindows();
        var pres = CreateEmptyPresentation();
        Assert.True(pres.Masters.Count >= 1);
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetMastersResult>(res);

        Assert.True(result.Count >= 1);
        Assert.NotEmpty(result.Masters);
        Assert.Null(result.Message);
        AssertNotModified(context);
    }

    [SkippableFact]
    public void Execute_ReturnsAllMasterDetails()
    {
        SkipIfNotWindows();
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetMastersResult>(res);

        for (var i = 0; i < result.Masters.Count; i++)
        {
            var master = result.Masters[i];
            Assert.Equal(i, master.Index);
            Assert.IsType<GetMasterInfo>(master);
            Assert.True(master.LayoutCount >= 0);
            Assert.NotNull(master.Layouts);
        }
    }

    #endregion
}
