using Aspose.Words;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;
using AsposeMcpServer.Tools.Word;

namespace AsposeMcpServer.Tests.Tools.Word;

/// <summary>
///     Integration tests for WordProtectionTool.
///     Focuses on session management, file I/O, and operation routing.
///     Detailed parameter validation and business logic tests are in Handler tests.
/// </summary>
public class WordProtectionToolTests : WordTestBase
{
    private readonly WordProtectionTool _tool;

    public WordProtectionToolTests()
    {
        _tool = new WordProtectionTool(SessionManager);
    }

    #region File I/O Smoke Tests

    [Fact]
    public void ProtectDocument_ShouldProtectDocument()
    {
        var docPath = CreateWordDocument("test_protect.docx");
        var outputPath = CreateTestFilePath("test_protect_output.docx");
        var result = _tool.Execute("protect", docPath, outputPath: outputPath,
            password: "test123", protectionType: "ReadOnly");
        var doc = new Document(outputPath);
        Assert.Equal(ProtectionType.ReadOnly, doc.ProtectionType);
        var data = GetResultData<SuccessResult>(result);
        Assert.Contains("ReadOnly", data.Message);
    }

    [Fact]
    public void UnprotectDocument_ShouldUnprotectDocument()
    {
        var docPath = CreateWordDocument("test_unprotect.docx");
        var doc = new Document(docPath);
        doc.Protect(ProtectionType.ReadOnly, "test123");
        doc.Save(docPath);

        var outputPath = CreateTestFilePath("test_unprotect_output.docx");
        var result = _tool.Execute("unprotect", docPath, outputPath: outputPath, password: "test123");
        var resultDoc = new Document(outputPath);
        Assert.Equal(ProtectionType.NoProtection, resultDoc.ProtectionType);
        var data = GetResultData<SuccessResult>(result);
        Assert.StartsWith("Protection removed successfully", data.Message);
    }

    [Fact]
    public void UnprotectDocument_WhenNotProtected_ShouldReturnMessage()
    {
        var docPath = CreateWordDocument("test_unprotect_notprotected.docx");
        var result = _tool.Execute("unprotect", docPath);
        var data = GetResultData<SuccessResult>(result);
        Assert.Contains("not protected", data.Message);
    }

    #endregion

    #region Operation Routing

    [Theory]
    [InlineData("PROTECT")]
    [InlineData("PrOtEcT")]
    [InlineData("protect")]
    public void Operation_ShouldBeCaseInsensitive(string operation)
    {
        var docPath = CreateWordDocument($"test_case_{operation}.docx");
        var outputPath = CreateTestFilePath($"test_case_{operation}_output.docx");
        var result = _tool.Execute(operation, docPath, outputPath: outputPath,
            password: "test123", protectionType: "ReadOnly");
        var data = GetResultData<SuccessResult>(result);
        Assert.Contains("protected", data.Message, StringComparison.OrdinalIgnoreCase);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocumentWithContent("test_unknown_op.docx", "Test content");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("unknown_operation", docPath));
        Assert.Contains("Unknown operation", ex.Message);
    }

    #endregion

    #region Session Management

    [Fact]
    public void ProtectDocument_WithSessionId_ShouldProtectInMemory()
    {
        var docPath = CreateWordDocument("test_session_protect.docx");
        var sessionId = OpenSession(docPath);
        var result = _tool.Execute("protect", sessionId: sessionId,
            password: "session123", protectionType: "ReadOnly");
        var data = GetResultData<SuccessResult>(result);
        Assert.Contains("ReadOnly", data.Message);
        var sessionDoc = SessionManager.GetDocument<Document>(sessionId);
        Assert.Equal(ProtectionType.ReadOnly, sessionDoc.ProtectionType);
    }

    [Fact]
    public void UnprotectDocument_WithSessionId_ShouldUnprotectInMemory()
    {
        var docPath = CreateWordDocument("test_session_unprotect.docx");
        var doc = new Document(docPath);
        doc.Protect(ProtectionType.ReadOnly, "test123");
        doc.Save(docPath);

        var sessionId = OpenSession(docPath);
        var result = _tool.Execute("unprotect", sessionId: sessionId, password: "test123");
        var data = GetResultData<SuccessResult>(result);
        Assert.StartsWith("Protection removed", data.Message);
        var sessionDoc = SessionManager.GetDocument<Document>(sessionId);
        Assert.Equal(ProtectionType.NoProtection, sessionDoc.ProtectionType);
    }

    [Fact]
    public void Execute_WithInvalidSessionId_ShouldThrowKeyNotFoundException()
    {
        Assert.Throws<KeyNotFoundException>(() =>
            _tool.Execute("protect", sessionId: "invalid_session_id", password: "test123"));
    }

    [Fact]
    public void Execute_WithBothPathAndSessionId_ShouldPreferSessionId()
    {
        var docPath1 = CreateWordDocument("test_path_protect.docx");
        var docPath2 = CreateWordDocument("test_session_protect2.docx");
        var sessionId = OpenSession(docPath2);
        _tool.Execute("protect", docPath1, sessionId,
            password: "both123", protectionType: "AllowOnlyRevisions");
        var sessionDoc = SessionManager.GetDocument<Document>(sessionId);
        Assert.Equal(ProtectionType.AllowOnlyRevisions, sessionDoc.ProtectionType);
        var fileDoc = new Document(docPath1);
        Assert.Equal(ProtectionType.NoProtection, fileDoc.ProtectionType);
    }

    #endregion
}
