using Aspose.Slides;
using AsposeMcpServer.Handlers.PowerPoint.Security;
using AsposeMcpServer.Results.PowerPoint.Security;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.Security;

public class GetStatusPptSecurityHandlerTests : PptHandlerTestBase
{
    private readonly GetStatusPptSecurityHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_GetStatus()
    {
        Assert.Equal("get_status", _handler.Operation);
    }

    #endregion

    #region Helper Methods

    /// <summary>
    ///     Creates a presentation with the _MarkAsFinal custom property initialized.
    ///     The GetStatusPptSecurityHandler reads this property, which must exist in DocumentProperties.
    /// </summary>
    private static Presentation CreatePresentationForSecurityStatus()
    {
        var pres = new Presentation();
        pres.DocumentProperties["_MarkAsFinal"] = false;
        return pres;
    }

    #endregion

    #region Modification State

    [Fact]
    public void Execute_ShouldNotMarkModified()
    {
        var pres = CreatePresentationForSecurityStatus();
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        _handler.Execute(context, parameters);

        AssertNotModified(context);
    }

    #endregion

    #region Default Status

    [Fact]
    public void Execute_Default_ReturnsSecurityStatus()
    {
        var pres = CreatePresentationForSecurityStatus();
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SecurityStatusPptResult>(res);
        Assert.NotNull(result);
    }

    [Fact]
    public void Execute_Default_NotEncrypted()
    {
        var pres = CreatePresentationForSecurityStatus();
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SecurityStatusPptResult>(res);
        Assert.False(result.IsEncrypted);
    }

    [Fact]
    public void Execute_Default_NotWriteProtected()
    {
        var pres = CreatePresentationForSecurityStatus();
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SecurityStatusPptResult>(res);
        Assert.False(result.IsWriteProtected);
    }

    [Fact]
    public void Execute_Default_NotMarkedFinal()
    {
        var pres = CreatePresentationForSecurityStatus();
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SecurityStatusPptResult>(res);
        Assert.False(result.IsMarkedFinal);
    }

    [Fact]
    public void Execute_Default_NotReadOnlyRecommended()
    {
        var pres = CreatePresentationForSecurityStatus();
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SecurityStatusPptResult>(res);
        Assert.False(result.IsReadOnlyRecommended);
    }

    [Fact]
    public void Execute_ReturnsMessage()
    {
        var pres = CreatePresentationForSecurityStatus();
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SecurityStatusPptResult>(res);
        Assert.Contains("Encrypted:", result.Message);
        Assert.Contains("WriteProtected:", result.Message);
    }

    #endregion

    #region With Protection Applied

    [Fact]
    public void Execute_WithWriteProtection_ReflectsStatus()
    {
        var pres = CreatePresentationForSecurityStatus();
        pres.ProtectionManager.SetWriteProtection("pass");
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SecurityStatusPptResult>(res);
        Assert.True(result.IsWriteProtected);
    }

    [Fact]
    public void Execute_WithEncryption_ReflectsStatus()
    {
        var pres = CreatePresentationForSecurityStatus();
        pres.ProtectionManager.Encrypt("secret");
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SecurityStatusPptResult>(res);
        Assert.True(result.IsEncrypted);
    }

    [Fact]
    public void Execute_WithMarkFinal_ReflectsStatus()
    {
        var pres = CreatePresentationForSecurityStatus();
        pres.DocumentProperties["_MarkAsFinal"] = true;
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SecurityStatusPptResult>(res);
        Assert.True(result.IsMarkedFinal);
    }

    #endregion
}
