using System.Runtime.Versioning;
using Aspose.Slides;
using AsposeMcpServer.Handlers.PowerPoint.Security;
using AsposeMcpServer.Results.PowerPoint.Security;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.Security;

[SupportedOSPlatform("windows")]
public class GetStatusPptSecurityHandlerTests : PptHandlerTestBase
{
    private readonly GetStatusPptSecurityHandler _handler = new();

    #region Operation Property

    [SkippableFact]
    public void Operation_Returns_GetStatus()
    {
        SkipIfNotWindows();
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

    [SkippableFact]
    public void Execute_ShouldNotMarkModified()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationForSecurityStatus();
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        _handler.Execute(context, parameters);

        AssertNotModified(context);
    }

    #endregion

    #region Default Status

    [SkippableFact]
    public void Execute_Default_ReturnsSecurityStatus()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationForSecurityStatus();
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SecurityStatusPptResult>(res);
        Assert.NotNull(result);
    }

    [SkippableFact]
    public void Execute_Default_NotEncrypted()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationForSecurityStatus();
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SecurityStatusPptResult>(res);
        Assert.False(result.IsEncrypted);
    }

    [SkippableFact]
    public void Execute_Default_NotWriteProtected()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationForSecurityStatus();
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SecurityStatusPptResult>(res);
        Assert.False(result.IsWriteProtected);
    }

    [SkippableFact]
    public void Execute_Default_NotMarkedFinal()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationForSecurityStatus();
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SecurityStatusPptResult>(res);
        Assert.False(result.IsMarkedFinal);
    }

    [SkippableFact]
    public void Execute_Default_NotReadOnlyRecommended()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationForSecurityStatus();
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SecurityStatusPptResult>(res);
        Assert.False(result.IsReadOnlyRecommended);
    }

    [SkippableFact]
    public void Execute_ReturnsMessage()
    {
        SkipIfNotWindows();
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

    [SkippableFact]
    public void Execute_WithWriteProtection_ReflectsStatus()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationForSecurityStatus();
        pres.ProtectionManager.SetWriteProtection("pass");
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SecurityStatusPptResult>(res);
        Assert.True(result.IsWriteProtected);
    }

    [SkippableFact]
    public void Execute_WithEncryption_ReflectsStatus()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationForSecurityStatus();
        pres.ProtectionManager.Encrypt("secret");
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SecurityStatusPptResult>(res);
        Assert.True(result.IsEncrypted);
    }

    [SkippableFact]
    public void Execute_WithMarkFinal_ReflectsStatus()
    {
        SkipIfNotWindows();
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
