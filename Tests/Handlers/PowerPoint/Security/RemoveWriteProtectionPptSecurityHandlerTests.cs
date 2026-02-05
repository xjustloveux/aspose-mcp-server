using AsposeMcpServer.Handlers.PowerPoint.Security;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.Security;

public class RemoveWriteProtectionPptSecurityHandlerTests : PptHandlerTestBase
{
    private readonly RemoveWriteProtectionPptSecurityHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_RemoveWriteProtection()
    {
        Assert.Equal("remove_write_protection", _handler.Operation);
    }

    #endregion

    #region Basic Remove Write Protection Operations

    [Fact]
    public void Execute_RemovesWriteProtection()
    {
        var pres = CreateEmptyPresentation();
        pres.ProtectionManager.SetWriteProtection("pass");
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
        pres.ProtectionManager.SetWriteProtection("pass");
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);
        Assert.Contains("Write protection removed", result.Message);
    }

    [Fact]
    public void Execute_NotProtected_StillSucceeds()
    {
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
    }

    #endregion
}
