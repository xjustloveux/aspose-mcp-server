using Aspose.Slides;
using Aspose.Slides.Export;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Results.PowerPoint.Security;
using AsposeMcpServer.Tests.Infrastructure;
using AsposeMcpServer.Tools.PowerPoint;

namespace AsposeMcpServer.Tests.Tools.PowerPoint;

/// <summary>
///     Integration tests for PptSecurityTool.
///     Focuses on session management, file I/O, and operation routing.
///     Detailed parameter validation and business logic tests are in Handler tests.
/// </summary>
public class PptSecurityToolTests : PptTestBase
{
    private readonly PptSecurityTool _tool;

    public PptSecurityToolTests()
    {
        _tool = new PptSecurityTool(SessionManager);
    }

    /// <summary>
    ///     Creates a presentation with the _MarkAsFinal property initialized for security status tests.
    /// </summary>
    private string CreateSecurityPresentation(string fileName)
    {
        var filePath = CreateTestFilePath(fileName);
        using var presentation = new Presentation();
        presentation.DocumentProperties["_MarkAsFinal"] = false;
        presentation.Save(filePath, SaveFormat.Pptx);
        return filePath;
    }

    private string CreatePresentationWithWriteProtection(string fileName, string password = "pass")
    {
        var filePath = CreateTestFilePath(fileName);
        using var presentation = new Presentation();
        presentation.DocumentProperties["_MarkAsFinal"] = false;
        presentation.ProtectionManager.SetWriteProtection(password);
        presentation.Save(filePath, SaveFormat.Pptx);
        return filePath;
    }

    #region File I/O Smoke Tests

    [Fact]
    public void Encrypt_ShouldEncrypt()
    {
        var pptPath = CreatePresentation("test_encrypt.pptx");
        var outputPath = CreateTestFilePath("test_encrypt_output.pptx");
        var result = _tool.Execute("encrypt", pptPath, password: "secret", outputPath: outputPath);
        var data = GetResultData<SuccessResult>(result);
        Assert.Contains("encrypted", data.Message);
    }

    [Fact]
    public void Decrypt_ShouldDecrypt()
    {
        var pptPath = CreatePresentation("test_decrypt.pptx");
        var outputPath = CreateTestFilePath("test_decrypt_output.pptx");
        var result = _tool.Execute("decrypt", pptPath, outputPath: outputPath);
        var data = GetResultData<SuccessResult>(result);
        Assert.Contains("decrypted", data.Message);
    }

    [Fact]
    public void SetWriteProtection_ShouldSetProtection()
    {
        var pptPath = CreatePresentation("test_set_wp.pptx");
        var outputPath = CreateTestFilePath("test_set_wp_output.pptx");
        var result = _tool.Execute("set_write_protection", pptPath, password: "edit_pass", outputPath: outputPath);
        var data = GetResultData<SuccessResult>(result);
        Assert.Contains("Write protection set", data.Message);
        using var presentation = new Presentation(outputPath);
        Assert.True(presentation.ProtectionManager.IsWriteProtected);
    }

    [Fact]
    public void RemoveWriteProtection_ShouldRemoveProtection()
    {
        var pptPath = CreatePresentationWithWriteProtection("test_remove_wp.pptx");
        var outputPath = CreateTestFilePath("test_remove_wp_output.pptx");
        var result = _tool.Execute("remove_write_protection", pptPath, outputPath: outputPath);
        var data = GetResultData<SuccessResult>(result);
        Assert.Contains("Write protection removed", data.Message);
    }

    [Fact]
    public void MarkFinal_ShouldMarkAsFinal()
    {
        var pptPath = CreateSecurityPresentation("test_mark_final.pptx");
        var outputPath = CreateTestFilePath("test_mark_final_output.pptx");
        var result = _tool.Execute("mark_final", pptPath, markAsFinal: true, outputPath: outputPath);
        var data = GetResultData<SuccessResult>(result);
        Assert.Contains("marked as final", data.Message);
    }

    [Fact]
    public void GetStatus_ShouldReturnStatus()
    {
        var pptPath = CreateSecurityPresentation("test_get_status.pptx");
        var result = _tool.Execute("get_status", pptPath);
        var data = GetResultData<SecurityStatusPptResult>(result);
        Assert.False(data.IsEncrypted);
        Assert.False(data.IsWriteProtected);
        Assert.False(data.IsMarkedFinal);
    }

    [Fact]
    public void GetStatus_WithWriteProtection_ShouldReflect()
    {
        var pptPath = CreatePresentationWithWriteProtection("test_get_status_wp.pptx");
        var result = _tool.Execute("get_status", pptPath);
        var data = GetResultData<SecurityStatusPptResult>(result);
        Assert.True(data.IsWriteProtected);
    }

    #endregion

    #region Operation Routing

    [Theory]
    [InlineData("ENCRYPT")]
    [InlineData("Encrypt")]
    [InlineData("encrypt")]
    public void Operation_ShouldBeCaseInsensitive(string operation)
    {
        var pptPath = CreatePresentation($"test_case_{operation}.pptx");
        var outputPath = CreateTestFilePath($"test_case_{operation}_output.pptx");
        var result = _tool.Execute(operation, pptPath, password: "pass", outputPath: outputPath);
        var data = GetResultData<SuccessResult>(result);
        Assert.Contains("encrypted", data.Message);
    }

    [Fact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        var pptPath = CreatePresentation("test_unknown_op.pptx");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("unknown", pptPath));
        Assert.Contains("Unknown operation", ex.Message);
    }

    #endregion

    #region Session Management

    [Fact]
    public void GetStatus_WithSessionId_ShouldReturnStatusFromMemory()
    {
        var pptPath = CreateSecurityPresentation("test_session_status.pptx");
        var sessionId = OpenSession(pptPath);
        var result = _tool.Execute("get_status", sessionId: sessionId);
        var data = GetResultData<SecurityStatusPptResult>(result);
        Assert.NotNull(data);
        Assert.False(data.IsEncrypted);
        var output = GetResultOutput<SecurityStatusPptResult>(result);
        Assert.True(output.IsSession);
    }

    [Fact]
    public void Encrypt_WithSessionId_ShouldEncryptInMemory()
    {
        var pptPath = CreatePresentation("test_session_encrypt.pptx");
        var sessionId = OpenSession(pptPath);
        var result = _tool.Execute("encrypt", sessionId: sessionId, password: "secret");
        var data = GetResultData<SuccessResult>(result);
        Assert.Contains("encrypted", data.Message);
        var output = GetResultOutput<SuccessResult>(result);
        Assert.True(output.IsSession);
    }

    [Fact]
    public void Decrypt_WithSessionId_ShouldDecryptInMemory()
    {
        var pptPath = CreatePresentation("test_session_decrypt.pptx");
        var sessionId = OpenSession(pptPath);
        var result = _tool.Execute("decrypt", sessionId: sessionId);
        var data = GetResultData<SuccessResult>(result);
        Assert.Contains("decrypted", data.Message);
        var output = GetResultOutput<SuccessResult>(result);
        Assert.True(output.IsSession);
    }

    [Fact]
    public void SetWriteProtection_WithSessionId_ShouldSetInMemory()
    {
        var pptPath = CreatePresentation("test_session_set_wp.pptx");
        var sessionId = OpenSession(pptPath);
        var result = _tool.Execute("set_write_protection", sessionId: sessionId, password: "edit_pass");
        var data = GetResultData<SuccessResult>(result);
        Assert.Contains("Write protection set", data.Message);
        var ppt = SessionManager.GetDocument<Presentation>(sessionId);
        Assert.True(ppt.ProtectionManager.IsWriteProtected);
        var output = GetResultOutput<SuccessResult>(result);
        Assert.True(output.IsSession);
    }

    [Fact]
    public void RemoveWriteProtection_WithSessionId_ShouldRemoveInMemory()
    {
        var pptPath = CreatePresentationWithWriteProtection("test_session_remove_wp.pptx");
        var sessionId = OpenSession(pptPath);
        var result = _tool.Execute("remove_write_protection", sessionId: sessionId);
        var data = GetResultData<SuccessResult>(result);
        Assert.Contains("Write protection removed", data.Message);
        var output = GetResultOutput<SuccessResult>(result);
        Assert.True(output.IsSession);
    }

    [Fact]
    public void MarkFinal_WithSessionId_ShouldMarkInMemory()
    {
        var pptPath = CreateSecurityPresentation("test_session_mark_final.pptx");
        var sessionId = OpenSession(pptPath);
        var result = _tool.Execute("mark_final", sessionId: sessionId, markAsFinal: true);
        var data = GetResultData<SuccessResult>(result);
        Assert.Contains("marked as final", data.Message);
        var output = GetResultOutput<SuccessResult>(result);
        Assert.True(output.IsSession);
    }

    [Fact]
    public void Execute_WithInvalidSessionId_ShouldThrowKeyNotFoundException()
    {
        Assert.Throws<KeyNotFoundException>(() => _tool.Execute("get_status", sessionId: "invalid_session"));
    }

    [Fact]
    public void Execute_WithBothPathAndSessionId_ShouldPreferSessionId()
    {
        var pptPath1 = CreatePresentationWithWriteProtection("test_path_security.pptx");
        var pptPath2 = CreateSecurityPresentation("test_session_security.pptx");
        var sessionId = OpenSession(pptPath2);
        var result = _tool.Execute("get_status", pptPath1, sessionId);
        var data = GetResultData<SecurityStatusPptResult>(result);
        Assert.False(data.IsWriteProtected);
        var output = GetResultOutput<SecurityStatusPptResult>(result);
        Assert.True(output.IsSession);
    }

    #endregion
}
