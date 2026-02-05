using AsposeMcpServer.Handlers.PowerPoint.Security;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.Security;

public class SetWriteProtectionPptSecurityHandlerTests : PptHandlerTestBase
{
    private readonly SetWriteProtectionPptSecurityHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_SetWriteProtection()
    {
        Assert.Equal("set_write_protection", _handler.Operation);
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

    #region Basic Set Write Protection Operations

    [Fact]
    public void Execute_SetsWriteProtection()
    {
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "password", "edit_pass" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        Assert.True(pres.ProtectionManager.IsWriteProtected);
        AssertModified(context);
    }

    [Fact]
    public void Execute_ReturnsSuccessMessage()
    {
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "password", "pass" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);
        Assert.Contains("Write protection set", result.Message);
    }

    #endregion
}
