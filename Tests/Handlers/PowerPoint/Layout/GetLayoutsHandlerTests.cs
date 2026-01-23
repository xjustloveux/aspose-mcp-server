using AsposeMcpServer.Handlers.PowerPoint.Layout;
using AsposeMcpServer.Results.PowerPoint.Layout;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.Layout;

public class GetLayoutsHandlerTests : PptHandlerTestBase
{
    private readonly GetLayoutsHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_GetLayouts()
    {
        Assert.Equal("get_layouts", _handler.Operation);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithInvalidMasterIndex_ThrowsArgumentException()
    {
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

    [Fact]
    public void Execute_ReturnsLayouts()
    {
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

    [Fact]
    public void Execute_WithMasterIndex_ReturnsLayoutsForMaster()
    {
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

    [Fact]
    public void Execute_ReturnsResultType()
    {
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
