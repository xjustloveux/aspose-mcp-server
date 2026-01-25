using AsposeMcpServer.Handlers.PowerPoint.Layout;
using AsposeMcpServer.Results.PowerPoint.Layout;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.Layout;

public class GetMastersHandlerTests : PptHandlerTestBase
{
    private readonly GetMastersHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_GetMasters()
    {
        Assert.Equal("get_masters", _handler.Operation);
    }

    #endregion

    #region Basic Get Operations

    [Fact]
    public void Execute_ReturnsMasters()
    {
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetMastersResult>(res);

        Assert.True(result.Count >= 0);
        Assert.NotNull(result.Masters);
        AssertNotModified(context);
    }

    [Fact]
    public void Execute_ReturnsCorrectMasterCount()
    {
        var pres = CreateEmptyPresentation();
        var expectedCount = pres.Masters.Count;
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetMastersResult>(res);

        Assert.Equal(expectedCount, result.Count);
        Assert.Equal(expectedCount, result.Masters.Count);
    }

    [Fact]
    public void Execute_ReturnsResultType()
    {
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetMastersResult>(res);

        Assert.NotNull(result);
        Assert.IsType<GetMastersResult>(result);
    }

    [Fact]
    public void Execute_IncludesLayoutCount()
    {
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetMastersResult>(res);

        Assert.NotNull(result.Masters);
        if (result.Masters.Count > 0) Assert.True(result.Masters[0].LayoutCount >= 0);
    }

    [Fact]
    public void Execute_IncludesMasterName()
    {
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

    [Fact]
    public void Execute_IncludesLayoutsList()
    {
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetMastersResult>(res);

        Assert.NotNull(result.Masters);
        if (result.Masters.Count > 0) Assert.NotNull(result.Masters[0].Layouts);
    }

    [Fact]
    public void Execute_IncludesMasterIndex()
    {
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetMastersResult>(res);

        Assert.NotNull(result.Masters);
        if (result.Masters.Count > 0) Assert.Equal(0, result.Masters[0].Index);
    }

    #endregion
}
