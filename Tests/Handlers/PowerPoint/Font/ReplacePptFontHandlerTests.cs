using System.Runtime.Versioning;
using AsposeMcpServer.Handlers.PowerPoint.Font;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.Font;

[SupportedOSPlatform("windows")]
public class ReplacePptFontHandlerTests : PptHandlerTestBase
{
    private readonly ReplacePptFontHandler _handler = new();

    #region Operation Property

    [SkippableFact]
    public void Operation_Returns_Replace()
    {
        SkipIfNotWindows();
        Assert.Equal("replace", _handler.Operation);
    }

    #endregion

    #region Replace Font

    [SkippableFact]
    public void Execute_ReplacesFont_ReturnsSuccessResult()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithText("Hello World");
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sourceFont", "Calibri" },
            { "targetFont", "Arial" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);
        Assert.Contains("Calibri", result.Message);
        Assert.Contains("Arial", result.Message);
    }

    [SkippableFact]
    public void Execute_ReplacesFont_MarksContextModified()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithText("Hello World");
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sourceFont", "Calibri" },
            { "targetFont", "Arial" }
        });

        _handler.Execute(context, parameters);

        AssertModified(context);
    }

    [SkippableFact]
    public void Execute_WithMissingSourceFont_ThrowsArgumentException()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithText("Hello World");
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "targetFont", "Arial" }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [SkippableFact]
    public void Execute_WithMissingTargetFont_ThrowsArgumentException()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithText("Hello World");
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sourceFont", "Calibri" }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion
}
