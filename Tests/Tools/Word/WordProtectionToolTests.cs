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

    #region General

    [Theory]
    [InlineData("ReadOnly", ProtectionType.ReadOnly)]
    [InlineData("AllowOnlyComments", ProtectionType.AllowOnlyComments)]
    [InlineData("AllowOnlyFormFields", ProtectionType.AllowOnlyFormFields)]
    [InlineData("AllowOnlyRevisions", ProtectionType.AllowOnlyRevisions)]
    public void ProtectDocument_WithDifferentProtectionTypes_ShouldProtectDocument(string protectionType,
        ProtectionType expectedType)
    {
        var docPath = CreateWordDocument($"test_protect_{protectionType.ToLower()}.docx");
        var outputPath = CreateTestFilePath($"test_protect_{protectionType.ToLower()}_output.docx");
        var result = _tool.Execute("protect", docPath, outputPath: outputPath,
            password: "test123", protectionType: protectionType);
        var doc = new Document(outputPath);
        Assert.Equal(expectedType, doc.ProtectionType);
        Assert.Contains(protectionType, result);
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
    public void ProtectDocument_WithInvalidProtectionType_ShouldUseDefaultReadOnly()
    {
        var docPath = CreateWordDocument("test_protect_invalid_type.docx");
        var outputPath = CreateTestFilePath("test_protect_invalid_type_output.docx");

        var result = _tool.Execute("protect", docPath, outputPath: outputPath,
            password: "test123", protectionType: "InvalidType");

        Assert.Contains("protected", result, StringComparison.OrdinalIgnoreCase);
        Assert.True(File.Exists(outputPath));
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
        Assert.StartsWith("Protection removed successfully", result);
        Assert.Contains("ReadOnly", result); // Check protection type was ReadOnly
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

    [Theory]
    [InlineData("PROTECT")]
    [InlineData("PrOtEcT")]
    [InlineData("protect")]
    public void Execute_ProtectOperationIsCaseInsensitive(string operation)
    {
        var docPath = CreateWordDocument($"test_case_{operation}.docx");
        var outputPath = CreateTestFilePath($"test_case_{operation}_output.docx");

        var result = _tool.Execute(operation, docPath, outputPath: outputPath,
            password: "test123", protectionType: "ReadOnly");

        Assert.Contains("protected", result, StringComparison.OrdinalIgnoreCase);
        Assert.True(File.Exists(outputPath));
    }

    [Theory]
    [InlineData("UNPROTECT")]
    [InlineData("UnPrOtEcT")]
    [InlineData("unprotect")]
    public void Execute_UnprotectOperationIsCaseInsensitive(string operation)
    {
        var docPath = CreateWordDocument($"test_unprotect_case_{operation}.docx");
        var doc = new Document(docPath);
        doc.Protect(ProtectionType.ReadOnly, "test123");
        doc.Save(docPath);

        var outputPath = CreateTestFilePath($"test_unprotect_case_{operation}_output.docx");
        var result = _tool.Execute(operation, docPath, outputPath: outputPath, password: "test123");

        Assert.StartsWith("Protection removed", result);
    }

    [Theory]
    [InlineData("READONLY", ProtectionType.ReadOnly)]
    [InlineData("readonly", ProtectionType.ReadOnly)]
    [InlineData("allowonlycomments", ProtectionType.AllowOnlyComments)]
    [InlineData("ALLOWONLYFORMFIELDS", ProtectionType.AllowOnlyFormFields)]
    [InlineData("AllowOnlyRevisions", ProtectionType.AllowOnlyRevisions)]
    public void ProtectDocument_ProtectionTypeIsCaseInsensitive(string protectionType, ProtectionType expectedType)
    {
        var docPath = CreateWordDocument($"test_ptype_{protectionType.ToLower()}.docx");
        var outputPath = CreateTestFilePath($"test_ptype_{protectionType.ToLower()}_output.docx");

        _tool.Execute("protect", docPath, outputPath: outputPath,
            password: "test123", protectionType: protectionType);

        var doc = new Document(outputPath);
        Assert.Equal(expectedType, doc.ProtectionType);
    }

    #endregion

    #region Exception

    [Fact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocumentWithContent("test_unknown_op.docx", "Test content");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("unknown_operation", docPath));

        Assert.Contains("Unknown operation", ex.Message);
        Assert.Contains("unknown_operation", ex.Message);
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    [InlineData(null)]
    public void ProtectDocument_WithInvalidPassword_ShouldThrowException(string? password)
    {
        var docPath = CreateWordDocument($"test_protect_pwd_{password?.Length ?? 0}.docx");
        var exception = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("protect", docPath, password: password, protectionType: "ReadOnly"));
        Assert.Contains("Password is required", exception.Message);
    }

    [Fact]
    public void ProtectDocument_WithNonExistentFile_ShouldThrowException()
    {
        var nonExistentPath = CreateTestFilePath("non_existent_file.docx");
        Assert.ThrowsAny<Exception>(() =>
            _tool.Execute("protect", nonExistentPath, password: "test123"));
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

    [Theory]
    [InlineData("protect")]
    [InlineData("unprotect")]
    public void Execute_WithoutPathAndSessionId_ShouldThrowException(string operation)
    {
        var exception = Assert.Throws<ArgumentException>(() =>
            _tool.Execute(operation, password: "test123", protectionType: "ReadOnly"));
        Assert.Contains("Either sessionId or path must be provided", exception.Message);
    }

    #endregion

    #region Session

    [Theory]
    [InlineData("ReadOnly", ProtectionType.ReadOnly)]
    [InlineData("AllowOnlyComments", ProtectionType.AllowOnlyComments)]
    [InlineData("AllowOnlyFormFields", ProtectionType.AllowOnlyFormFields)]
    [InlineData("AllowOnlyRevisions", ProtectionType.AllowOnlyRevisions)]
    public void ProtectDocument_WithSessionId_ShouldProtectInMemory(string protectionType, ProtectionType expectedType)
    {
        var docPath = CreateWordDocument($"test_session_protect_{protectionType.ToLower()}.docx");
        var sessionId = OpenSession(docPath);
        var result = _tool.Execute("protect", sessionId: sessionId,
            password: "session123", protectionType: protectionType);
        Assert.Contains(protectionType, result);

        var sessionDoc = SessionManager.GetDocument<Document>(sessionId);
        Assert.Equal(expectedType, sessionDoc.ProtectionType);
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
        Assert.StartsWith("Protection removed", result);

        var sessionDoc = SessionManager.GetDocument<Document>(sessionId);
        Assert.Equal(ProtectionType.NoProtection, sessionDoc.ProtectionType);
    }

    [Fact]
    public void UnprotectDocument_WithSessionId_WhenNotProtected_ShouldReturnMessage()
    {
        var docPath = CreateWordDocument("test_session_unprotect_none.docx");
        var sessionId = OpenSession(docPath);
        var result = _tool.Execute("unprotect", sessionId: sessionId);
        Assert.Contains("not protected", result);
    }

    [Theory]
    [InlineData("protect")]
    [InlineData("unprotect")]
    public void Execute_WithInvalidSessionId_ShouldThrowKeyNotFoundException(string operation)
    {
        Assert.Throws<KeyNotFoundException>(() =>
            _tool.Execute(operation, sessionId: "invalid_session_id", password: "test123"));
    }

    [Fact]
    public void Execute_WithBothPathAndSessionId_ShouldPreferSessionId()
    {
        var docPath1 = CreateWordDocument("test_path_protect.docx");
        var docPath2 = CreateWordDocument("test_session_protect.docx");

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