using AsposeMcpServer.Handlers.PowerPoint.Layout;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.Layout;

public class ApplyMasterHandlerTests : PptHandlerTestBase
{
    private readonly ApplyMasterHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_ApplyMaster()
    {
        Assert.Equal("apply_master", _handler.Operation);
    }

    #endregion

    #region Basic Apply Operations

    [Fact]
    public void Execute_AppliesMaster()
    {
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "masterIndex", 0 },
            { "layoutIndex", 0 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("applied", result.ToLower());
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithSlideIndices_AppliesToSpecificSlides()
    {
        var pres = CreatePresentationWithSlides(3);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "masterIndex", 0 },
            { "layoutIndex", 0 },
            { "slideIndices", "[0, 1]" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("2", result);
        Assert.Contains("applied", result.ToLower());
        AssertModified(context);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutMasterIndex_ThrowsArgumentException()
    {
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "layoutIndex", 0 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithoutLayoutIndex_ThrowsArgumentException()
    {
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "masterIndex", 0 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithInvalidMasterIndex_ThrowsArgumentException()
    {
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "masterIndex", 99 },
            { "layoutIndex", 0 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithInvalidLayoutIndex_ThrowsArgumentException()
    {
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "masterIndex", 0 },
            { "layoutIndex", 999 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion
}
