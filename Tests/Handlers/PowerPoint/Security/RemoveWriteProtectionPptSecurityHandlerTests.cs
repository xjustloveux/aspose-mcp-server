using System.Runtime.Versioning;
using AsposeMcpServer.Handlers.PowerPoint.Security;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.Security;

[SupportedOSPlatform("windows")]
public class RemoveWriteProtectionPptSecurityHandlerTests : PptHandlerTestBase
{
    private readonly RemoveWriteProtectionPptSecurityHandler _handler = new();

    #region Operation Property

    [SkippableFact]
    public void Operation_Returns_RemoveWriteProtection()
    {
        SkipIfNotWindows();
        Assert.Equal("remove_write_protect", _handler.Operation);
    }

    #endregion

    #region Basic Remove Write Protection Operations

    [SkippableFact]
    public void Execute_RemovesWriteProtection()
    {
        SkipIfNotWindows();
        var pres = CreateEmptyPresentation();
        pres.ProtectionManager.SetWriteProtection("pass");
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
        pres.ProtectionManager.SetWriteProtection("pass");
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);
        Assert.Contains("Write protection removed", result.Message);
    }

    [SkippableFact]
    public void Execute_NotProtected_StillSucceeds()
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
