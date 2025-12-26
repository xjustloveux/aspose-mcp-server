using Aspose.Words;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Word;

namespace AsposeMcpServer.Tests.Word;

public class WordProtectionToolTests : WordTestBase
{
    private readonly WordProtectionTool _tool = new();

    [Fact]
    public async Task ProtectDocument_WithReadOnly_ShouldProtectDocument()
    {
        // Arrange
        var docPath = CreateWordDocument("test_protect_readonly.docx");
        var outputPath = CreateTestFilePath("test_protect_readonly_output.docx");
        var arguments = CreateArguments("protect", docPath, outputPath);
        arguments["password"] = "test123";
        arguments["protectionType"] = "ReadOnly";

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        var doc = new Document(outputPath);
        Assert.Equal(ProtectionType.ReadOnly, doc.ProtectionType);
        Assert.Contains("ReadOnly", result);
    }

    [Fact]
    public async Task ProtectDocument_WithAllowOnlyComments_ShouldProtectDocument()
    {
        // Arrange
        var docPath = CreateWordDocument("test_protect_comments.docx");
        var outputPath = CreateTestFilePath("test_protect_comments_output.docx");
        var arguments = CreateArguments("protect", docPath, outputPath);
        arguments["password"] = "test123";
        arguments["protectionType"] = "AllowOnlyComments";

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        var doc = new Document(outputPath);
        Assert.Equal(ProtectionType.AllowOnlyComments, doc.ProtectionType);
        Assert.Contains("AllowOnlyComments", result);
    }

    [Fact]
    public async Task ProtectDocument_WithAllowOnlyFormFields_ShouldProtectDocument()
    {
        // Arrange
        var docPath = CreateWordDocument("test_protect_formfields.docx");
        var outputPath = CreateTestFilePath("test_protect_formfields_output.docx");
        var arguments = CreateArguments("protect", docPath, outputPath);
        arguments["password"] = "test123";
        arguments["protectionType"] = "AllowOnlyFormFields";

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        var doc = new Document(outputPath);
        Assert.Equal(ProtectionType.AllowOnlyFormFields, doc.ProtectionType);
        Assert.Contains("AllowOnlyFormFields", result);
    }

    [Fact]
    public async Task ProtectDocument_WithAllowOnlyRevisions_ShouldProtectDocument()
    {
        // Arrange
        var docPath = CreateWordDocument("test_protect_revisions.docx");
        var outputPath = CreateTestFilePath("test_protect_revisions_output.docx");
        var arguments = CreateArguments("protect", docPath, outputPath);
        arguments["password"] = "test123";
        arguments["protectionType"] = "AllowOnlyRevisions";

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        var doc = new Document(outputPath);
        Assert.Equal(ProtectionType.AllowOnlyRevisions, doc.ProtectionType);
        Assert.Contains("AllowOnlyRevisions", result);
    }

    [Fact]
    public async Task UnprotectDocument_ShouldUnprotectDocument()
    {
        // Arrange
        var docPath = CreateWordDocument("test_unprotect.docx");
        var doc = new Document(docPath);
        doc.Protect(ProtectionType.ReadOnly, "test123");
        doc.Save(docPath);

        var outputPath = CreateTestFilePath("test_unprotect_output.docx");
        var arguments = CreateArguments("unprotect", docPath, outputPath);
        arguments["password"] = "test123";

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        var resultDoc = new Document(outputPath);
        Assert.Equal(ProtectionType.NoProtection, resultDoc.ProtectionType);
        Assert.Contains("Protection removed successfully", result);
        Assert.Contains("ReadOnly", result); // Should show previous protection type
    }

    [Fact]
    public async Task UnprotectDocument_WhenNotProtected_ShouldReturnMessage()
    {
        // Arrange
        var docPath = CreateWordDocument("test_unprotect_notprotected.docx");
        var arguments = CreateArguments("unprotect", docPath);

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("not protected", result);
    }

    [Fact]
    public async Task UnprotectDocument_WhenNotProtected_WithDifferentOutputPath_ShouldSave()
    {
        // Arrange
        var docPath = CreateWordDocument("test_unprotect_notprotected_save.docx");
        var outputPath = CreateTestFilePath("test_unprotect_notprotected_save_output.docx");
        var arguments = CreateArguments("unprotect", docPath, outputPath);

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("not protected", result);
        Assert.Contains("saved to", result);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public async Task UnprotectDocument_WithWrongPassword_ShouldThrowException()
    {
        // Arrange
        var docPath = CreateWordDocument("test_unprotect_wrongpwd.docx");
        var doc = new Document(docPath);
        doc.Protect(ProtectionType.ReadOnly, "correctpassword");
        doc.Save(docPath);

        var arguments = CreateArguments("unprotect", docPath);
        arguments["password"] = "wrongpassword";

        // Act & Assert
        var exception = await Assert.ThrowsAsync<InvalidOperationException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("password", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task ProtectDocument_DefaultProtectionType_ShouldUseReadOnly()
    {
        // Arrange
        var docPath = CreateWordDocument("test_protect_default.docx");
        var outputPath = CreateTestFilePath("test_protect_default_output.docx");
        var arguments = CreateArguments("protect", docPath, outputPath);
        arguments["password"] = "test123";
        // Not specifying protectionType - should default to ReadOnly

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var doc = new Document(outputPath);
        Assert.Equal(ProtectionType.ReadOnly, doc.ProtectionType);
    }

    [Fact]
    public async Task ProtectDocument_WithEmptyPassword_ShouldThrowException()
    {
        // Arrange
        var docPath = CreateWordDocument("test_protect_emptypwd.docx");
        var arguments = CreateArguments("protect", docPath);
        arguments["password"] = "";
        arguments["protectionType"] = "ReadOnly";

        // Act & Assert
        var exception = await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("Password is required", exception.Message);
    }

    [Fact]
    public async Task ProtectDocument_WithoutPassword_ShouldThrowException()
    {
        // Arrange
        var docPath = CreateWordDocument("test_protect_nopwd.docx");
        var arguments = CreateArguments("protect", docPath);
        arguments["protectionType"] = "ReadOnly";
        // No password provided

        // Act & Assert
        var exception = await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("Password is required", exception.Message);
    }

    [Fact]
    public async Task ProtectDocument_WithNonExistentFile_ShouldThrowException()
    {
        // Arrange
        var nonExistentPath = CreateTestFilePath("non_existent_file.docx");
        var arguments = CreateArguments("protect", nonExistentPath);
        arguments["password"] = "test123";

        // Act & Assert
        var exception = await Assert.ThrowsAsync<InvalidOperationException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("not found", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task ProtectDocument_WithCaseInsensitiveProtectionType_ShouldWork()
    {
        // Arrange
        var docPath = CreateWordDocument("test_protect_lowercase.docx");
        var outputPath = CreateTestFilePath("test_protect_lowercase_output.docx");
        var arguments = CreateArguments("protect", docPath, outputPath);
        arguments["password"] = "test123";
        arguments["protectionType"] = "readonly"; // lowercase

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var doc = new Document(outputPath);
        Assert.Equal(ProtectionType.ReadOnly, doc.ProtectionType);
    }
}