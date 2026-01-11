using AsposeMcpServer.Handlers.PowerPoint.Layout;
using AsposeMcpServer.Tests.Helpers;

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

        var result = _handler.Execute(context, parameters);

        Assert.Contains("masters", result.ToLower());
        Assert.Contains("layouts", result.ToLower());
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

        var result = _handler.Execute(context, parameters);

        Assert.Contains("masterIndex", result);
        Assert.Contains("layouts", result.ToLower());
        AssertNotModified(context);
    }

    [Fact]
    public void Execute_ReturnsJsonFormat()
    {
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("{", result);
        Assert.Contains("}", result);
    }

    #endregion
}
