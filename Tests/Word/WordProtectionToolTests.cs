using Aspose.Words;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Word;

namespace AsposeMcpServer.Tests.Word;

public class WordProtectionToolTests : WordTestBase
{
    private readonly WordProtectionTool _tool = new();

    [Fact]
    public async Task ProtectDocument_ShouldProtectDocument()
    {
        // Arrange
        var docPath = CreateWordDocument("test_protect.docx");
        var outputPath = CreateTestFilePath("test_protect_output.docx");
        var arguments = CreateArguments("protect", docPath, outputPath);
        arguments["password"] = "test123";
        arguments["protectionType"] = "ReadOnly";

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var doc = new Document(outputPath);
        Assert.True(doc.ProtectionType != ProtectionType.NoProtection, "Document should be protected");
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
        await _tool.ExecuteAsync(arguments);

        // Assert
        var resultDoc = new Document(outputPath);
        Assert.Equal(ProtectionType.NoProtection, resultDoc.ProtectionType);
    }
}