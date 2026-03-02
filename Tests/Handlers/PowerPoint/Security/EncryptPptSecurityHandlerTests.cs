using System.Runtime.Versioning;
using AsposeMcpServer.Handlers.PowerPoint.Security;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.Security;

[SupportedOSPlatform("windows")]
public class EncryptPptSecurityHandlerTests : PptHandlerTestBase
{
    private readonly EncryptPptSecurityHandler _handler = new();

    #region Operation Property

    [SkippableFact]
    public void Operation_Returns_Encrypt()
    {
        SkipIfNotWindows();
        Assert.Equal("encrypt", _handler.Operation);
    }

    #endregion

    #region Error Handling

    [SkippableFact]
    public void Execute_MissingPassword_ThrowsArgumentException()
    {
        SkipIfNotWindows();
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion

    #region Basic Encrypt Operations

    [SkippableFact]
    public void Execute_EncryptsPresentation()
    {
        SkipIfNotWindows();
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "password", "test123" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        Assert.True(pres.ProtectionManager.IsEncrypted);
        AssertModified(context);
    }

    [SkippableFact]
    public void Execute_ReturnsSuccessMessage()
    {
        SkipIfNotWindows();
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "password", "secret" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);
        Assert.Contains("encrypted", result.Message);
    }

    #endregion
}
