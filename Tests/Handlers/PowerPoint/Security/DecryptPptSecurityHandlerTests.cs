using System.Runtime.Versioning;
using AsposeMcpServer.Handlers.PowerPoint.Security;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.Security;

[SupportedOSPlatform("windows")]
public class DecryptPptSecurityHandlerTests : PptHandlerTestBase
{
    private readonly DecryptPptSecurityHandler _handler = new();

    #region Operation Property

    [SkippableFact]
    public void Operation_Returns_Decrypt()
    {
        SkipIfNotWindows();
        Assert.Equal("decrypt", _handler.Operation);
    }

    #endregion

    #region Basic Decrypt Operations

    [SkippableFact]
    public void Execute_DecryptsEncryptedPresentation()
    {
        SkipIfNotWindows();
        var pres = CreateEmptyPresentation();
        pres.ProtectionManager.Encrypt("test");
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        AssertModified(context);
    }

    [SkippableFact]
    public void Execute_ReturnsSuccessMessage()
    {
        SkipIfNotWindows();
        var pres = CreateEmptyPresentation();
        pres.ProtectionManager.Encrypt("test");
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);
        Assert.Contains("decrypted", result.Message);
    }

    [SkippableFact]
    public void Execute_NotEncrypted_StillSucceeds()
    {
        SkipIfNotWindows();
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
    }

    #endregion
}
