using AsposeMcpServer.Handlers.PowerPoint.Layout;
using AsposeMcpServer.Tests.Helpers;

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

        var result = _handler.Execute(context, parameters);

        Assert.Contains("count", result.ToLower());
        Assert.Contains("masters", result.ToLower());
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

    [Fact]
    public void Execute_IncludesLayoutCount()
    {
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("layoutCount", result);
    }

    #endregion
}
