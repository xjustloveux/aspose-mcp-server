using AsposeMcpServer.Handlers.PowerPoint.Security;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.Security;

public class DecryptPptSecurityHandlerTests : PptHandlerTestBase
{
    private readonly DecryptPptSecurityHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Decrypt()
    {
        Assert.Equal("decrypt", _handler.Operation);
    }

    #endregion

    #region Basic Decrypt Operations

    [Fact]
    public void Execute_DecryptsEncryptedPresentation()
    {
        var pres = CreateEmptyPresentation();
        pres.ProtectionManager.Encrypt("test");
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        AssertModified(context);
    }

    [Fact]
    public void Execute_ReturnsSuccessMessage()
    {
        var pres = CreateEmptyPresentation();
        pres.ProtectionManager.Encrypt("test");
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);
        Assert.Contains("decrypted", result.Message);
    }

    [Fact]
    public void Execute_NotEncrypted_StillSucceeds()
    {
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
    }

    #endregion
}
