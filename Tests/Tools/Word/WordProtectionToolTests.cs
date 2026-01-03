using Aspose.Words;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Word;

namespace AsposeMcpServer.Tests.Tools.Word;

public class WordProtectionToolTests : WordTestBase
{
    private readonly WordProtectionTool _tool;

    public WordProtectionToolTests()
    {
        _tool = new WordProtectionTool(SessionManager);
    }

    #region General Tests

    [Fact]
    public void ProtectDocument_WithReadOnly_ShouldProtectDocument()
    {
        var docPath = CreateWordDocument("test_protect_readonly.docx");
        var outputPath = CreateTestFilePath("test_protect_readonly_output.docx");
        var result = _tool.Execute("protect", docPath, outputPath: outputPath,
            password: "test123", protectionType: "ReadOnly");
        var doc = new Document(outputPath);
        Assert.Equal(ProtectionType.ReadOnly, doc.ProtectionType);
        Assert.Contains("ReadOnly", result);
    }

    [Fact]
    public void ProtectDocument_WithAllowOnlyComments_ShouldProtectDocument()
    {
        var docPath = CreateWordDocument("test_protect_comments.docx");
        var outputPath = CreateTestFilePath("test_protect_comments_output.docx");
        var result = _tool.Execute("protect", docPath, outputPath: outputPath,
            password: "test123", protectionType: "AllowOnlyComments");
        var doc = new Document(outputPath);
        Assert.Equal(ProtectionType.AllowOnlyComments, doc.ProtectionType);
        Assert.Contains("AllowOnlyComments", result);
    }

    [Fact]
    public void ProtectDocument_WithAllowOnlyFormFields_ShouldProtectDocument()
    {
        var docPath = CreateWordDocument("test_protect_formfields.docx");
        var outputPath = CreateTestFilePath("test_protect_formfields_output.docx");
        var result = _tool.Execute("protect", docPath, outputPath: outputPath,
            password: "test123", protectionType: "AllowOnlyFormFields");
        var doc = new Document(outputPath);
        Assert.Equal(ProtectionType.AllowOnlyFormFields, doc.ProtectionType);
        Assert.Contains("AllowOnlyFormFields", result);
    }

    [Fact]
    public void ProtectDocument_WithAllowOnlyRevisions_ShouldProtectDocument()
    {
        var docPath = CreateWordDocument("test_protect_revisions.docx");
        var outputPath = CreateTestFilePath("test_protect_revisions_output.docx");
        var result = _tool.Execute("protect", docPath, outputPath: outputPath,
            password: "test123", protectionType: "AllowOnlyRevisions");
        var doc = new Document(outputPath);
        Assert.Equal(ProtectionType.AllowOnlyRevisions, doc.ProtectionType);
        Assert.Contains("AllowOnlyRevisions", result);
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
        Assert.Contains("Protection removed successfully", result);
        Assert.Contains("ReadOnly", result); // Should show previous protection type
    }

    [Fact]
    public void UnprotectDocument_WhenNotProtected_ShouldReturnMessage()
    {
        var docPath = CreateWordDocument("test_unprotect_notprotected.docx");
        var result = _tool.Execute("unprotect", docPath);
        Assert.Contains("not protected", result);
    }

    [Fact]
    public void UnprotectDocument_WhenNotProtected_WithDifferentOutputPath_ShouldSave()
    {
        var docPath = CreateWordDocument("test_unprotect_notprotected_save.docx");
        var outputPath = CreateTestFilePath("test_unprotect_notprotected_save_output.docx");
        var result = _tool.Execute("unprotect", docPath, outputPath: outputPath);
        Assert.Contains("not protected", result);
        Assert.Contains("saved to", result);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void UnprotectDocument_WithWrongPassword_ShouldThrowException()
    {
        var docPath = CreateWordDocument("test_unprotect_wrongpwd.docx");
        var doc = new Document(docPath);
        doc.Protect(ProtectionType.ReadOnly, "correctpassword");
        doc.Save(docPath);
        var exception = Assert.Throws<InvalidOperationException>(() =>
            _tool.Execute("unprotect", docPath, password: "wrongpassword"));
        Assert.Contains("password", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void ProtectDocument_DefaultProtectionType_ShouldUseReadOnly()
    {
        var docPath = CreateWordDocument("test_protect_default.docx");
        var outputPath = CreateTestFilePath("test_protect_default_output.docx");
        _tool.Execute("protect", docPath, outputPath: outputPath, password: "test123");
        var doc = new Document(outputPath);
        Assert.Equal(ProtectionType.ReadOnly, doc.ProtectionType);
    }

    [Fact]
    public void ProtectDocument_WithCaseInsensitiveProtectionType_ShouldWork()
    {
        var docPath = CreateWordDocument("test_protect_lowercase.docx");
        var outputPath = CreateTestFilePath("test_protect_lowercase_output.docx");
        _tool.Execute("protect", docPath, outputPath: outputPath,
            password: "test123", protectionType: "readonly");
        var doc = new Document(outputPath);
        Assert.Equal(ProtectionType.ReadOnly, doc.ProtectionType);
    }

    #endregion

    #region Exception Tests

    [Fact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocumentWithContent("test_unknown_op.docx", "Test content");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("unknown_operation", docPath));

        Assert.Contains("Unknown operation", ex.Message);
        Assert.Contains("unknown_operation", ex.Message);
    }

    [Fact]
    public void ProtectDocument_WithEmptyPassword_ShouldThrowException()
    {
        var docPath = CreateWordDocument("test_protect_emptypwd.docx");
        var exception = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("protect", docPath, password: "", protectionType: "ReadOnly"));
        Assert.Contains("Password is required", exception.Message);
    }

    [Fact]
    public void ProtectDocument_WithoutPassword_ShouldThrowException()
    {
        var docPath = CreateWordDocument("test_protect_nopwd.docx");
        var exception = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("protect", docPath, protectionType: "ReadOnly"));
        Assert.Contains("Password is required", exception.Message);
    }

    [Fact]
    public void ProtectDocument_WithNonExistentFile_ShouldThrowException()
    {
        var nonExistentPath = CreateTestFilePath("non_existent_file.docx");
        var exception = Assert.Throws<InvalidOperationException>(() =>
            _tool.Execute("protect", nonExistentPath, password: "test123"));
        Assert.Contains("not found", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void ProtectDocument_WithInvalidProtectionType_ShouldUseDefaultReadOnly()
    {
        var docPath = CreateWordDocument("test_protect_invalid_type.docx");
        var outputPath = CreateTestFilePath("test_protect_invalid_type_output.docx");

        // Act - Invalid protection type should fall back to ReadOnly
        var result = _tool.Execute("protect", docPath, outputPath: outputPath,
            password: "test123", protectionType: "InvalidType");

        // Assert - Should succeed with default ReadOnly protection
        Assert.Contains("protected", result, StringComparison.OrdinalIgnoreCase);
        Assert.True(File.Exists(outputPath));
    }

    #endregion

    #region Session ID Tests

    [Fact]
    public void ProtectDocument_WithSessionId_ShouldProtectInMemory()
    {
        var docPath = CreateWordDocument("test_session_protect.docx");
        var sessionId = OpenSession(docPath);
        var result = _tool.Execute("protect", sessionId: sessionId,
            password: "session123", protectionType: "ReadOnly");
        Assert.Contains("ReadOnly", result);

        // Verify in-memory document is protected
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
        Assert.Contains("Protection removed", result);

        // Verify in-memory document is unprotected
        var sessionDoc = SessionManager.GetDocument<Document>(sessionId);
        Assert.Equal(ProtectionType.NoProtection, sessionDoc.ProtectionType);
    }

    [Fact]
    public void ProtectDocument_WithSessionId_VerifyInMemoryProtection()
    {
        var docPath = CreateWordDocument("test_session_verify_protect.docx");
        var sessionId = OpenSession(docPath);

        // First protect
        _tool.Execute("protect", sessionId: sessionId,
            password: "verify123", protectionType: "AllowOnlyComments");

        // Act - verify protection type via in-memory document
        var sessionDoc = SessionManager.GetDocument<Document>(sessionId);
        Assert.Equal(ProtectionType.AllowOnlyComments, sessionDoc.ProtectionType);
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
        var docPath2 = CreateWordDocument("test_session_protect.docx");

        var sessionId = OpenSession(docPath2);

        // Act - provide both path and sessionId
        _tool.Execute("protect", docPath1, sessionId,
            password: "both123", protectionType: "AllowOnlyRevisions");

        // Assert - session document should be protected
        var sessionDoc = SessionManager.GetDocument<Document>(sessionId);
        Assert.Equal(ProtectionType.AllowOnlyRevisions, sessionDoc.ProtectionType);

        // Original file should not be protected
        var fileDoc = new Document(docPath1);
        Assert.Equal(ProtectionType.NoProtection, fileDoc.ProtectionType);
    }

    [Fact]
    public void UnprotectDocument_WithSessionId_WhenNotProtected_ShouldReturnMessage()
    {
        var docPath = CreateWordDocument("test_session_unprotect_none.docx");
        var sessionId = OpenSession(docPath);
        var result = _tool.Execute("unprotect", sessionId: sessionId);
        Assert.Contains("not protected", result);
    }

    #endregion
}