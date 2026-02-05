using AsposeMcpServer.Handlers.PowerPoint.Security;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.Security;

public class EncryptPptSecurityHandlerTests : PptHandlerTestBase
{
    private readonly EncryptPptSecurityHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Encrypt()
    {
        Assert.Equal("encrypt", _handler.Operation);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_MissingPassword_ThrowsArgumentException()
    {
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion

    #region Basic Encrypt Operations

    [Fact]
    public void Execute_EncryptsPresentation()
    {
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

    [Fact]
    public void Execute_ReturnsSuccessMessage()
    {
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
