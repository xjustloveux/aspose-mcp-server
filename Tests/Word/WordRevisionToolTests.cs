using Aspose.Words;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Word;

namespace AsposeMcpServer.Tests.Word;

public class WordRevisionToolTests : WordTestBase
{
    private readonly WordRevisionTool _tool = new();

    [Fact]
    public async Task GetRevisions_ShouldReturnRevisions()
    {
        // Arrange
        var docPath = CreateWordDocument("test_get_revisions.docx");
        var doc = new Document(docPath);
        doc.StartTrackRevisions("Test Author");
        var builder = new DocumentBuilder(doc);
        builder.Writeln("Original text");
        builder.Writeln("Modified text");
        doc.StopTrackRevisions();
        doc.Save(docPath);

        var arguments = CreateArguments("get_revisions", docPath);

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.NotNull(result);
        Assert.NotEmpty(result);
        Assert.Contains("\"revisions\"", result); // JSON format
        Assert.Contains("\"index\"", result); // Has index property
    }

    [Fact]
    public async Task GetRevisions_WithNoRevisions_ShouldReturnZeroCount()
    {
        // Arrange
        var docPath = CreateWordDocumentWithContent("test_no_revisions.docx", "Plain text");
        var arguments = CreateArguments("get_revisions", docPath);

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("\"count\": 0", result); // JSON format
    }

    [Fact]
    public async Task AcceptAllRevisions_ShouldAcceptAll()
    {
        // Arrange
        var docPath = CreateWordDocument("test_accept_all_revisions.docx");
        var doc = new Document(docPath);
        doc.StartTrackRevisions("Test Author");
        var builder = new DocumentBuilder(doc);
        builder.Writeln("Text with revision");
        doc.StopTrackRevisions();
        doc.Save(docPath);

        var revisionsBefore = doc.Revisions.Count;
        Assert.True(revisionsBefore > 0, "Document should have revisions before accepting");

        var outputPath = CreateTestFilePath("test_accept_all_revisions_output.docx");
        var arguments = CreateArguments("accept_all", docPath, outputPath);

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        var resultDoc = new Document(outputPath);
        Assert.Equal(0, resultDoc.Revisions.Count);
        Assert.Contains("Accepted", result);
    }

    [Fact]
    public async Task RejectAllRevisions_ShouldRejectAll()
    {
        // Arrange
        var docPath = CreateWordDocument("test_reject_all_revisions.docx");
        var doc = new Document(docPath);
        doc.StartTrackRevisions("Test Author");
        var builder = new DocumentBuilder(doc);
        builder.Writeln("Text with revision");
        doc.StopTrackRevisions();
        doc.Save(docPath);

        var outputPath = CreateTestFilePath("test_reject_all_revisions_output.docx");
        var arguments = CreateArguments("reject_all", docPath, outputPath);

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        var resultDoc = new Document(outputPath);
        Assert.Equal(0, resultDoc.Revisions.Count);
        Assert.Contains("Rejected", result);
    }

    [Fact]
    public async Task ManageRevision_Accept_ShouldAcceptSpecificRevision()
    {
        // Arrange
        var docPath = CreateWordDocument("test_manage_accept.docx");
        var doc = new Document(docPath);
        doc.StartTrackRevisions("Test Author");
        var builder = new DocumentBuilder(doc);
        builder.Writeln("First revision");
        builder.Writeln("Second revision");
        doc.StopTrackRevisions();
        doc.Save(docPath);

        var outputPath = CreateTestFilePath("test_manage_accept_output.docx");
        var arguments = CreateArguments("manage", docPath, outputPath);
        arguments["revisionIndex"] = 0;
        arguments["action"] = "accept";

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("[0]", result);
        Assert.Contains("accepted", result);
    }

    [Fact]
    public async Task ManageRevision_Reject_ShouldRejectSpecificRevision()
    {
        // Arrange
        var docPath = CreateWordDocument("test_manage_reject.docx");
        var doc = new Document(docPath);
        doc.StartTrackRevisions("Test Author");
        var builder = new DocumentBuilder(doc);
        builder.Writeln("Revision to reject");
        doc.StopTrackRevisions();
        doc.Save(docPath);

        var outputPath = CreateTestFilePath("test_manage_reject_output.docx");
        var arguments = CreateArguments("manage", docPath, outputPath);
        arguments["revisionIndex"] = 0;
        arguments["action"] = "reject";

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("[0]", result);
        Assert.Contains("rejected", result);
    }

    [Fact]
    public async Task ManageRevision_WithInvalidIndex_ShouldThrowException()
    {
        // Arrange
        var docPath = CreateWordDocument("test_manage_invalid_index.docx");
        var doc = new Document(docPath);
        doc.StartTrackRevisions("Test Author");
        var builder = new DocumentBuilder(doc);
        builder.Writeln("Single revision");
        doc.StopTrackRevisions();
        doc.Save(docPath);

        var arguments = CreateArguments("manage", docPath);
        arguments["revisionIndex"] = 99;
        arguments["action"] = "accept";

        // Act & Assert
        var exception = await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("revisionIndex must be between", exception.Message);
    }

    [Fact]
    public async Task ManageRevision_WithNoRevisions_ShouldReturnMessage()
    {
        // Arrange
        var docPath = CreateWordDocumentWithContent("test_manage_no_rev.docx", "Plain text");
        var arguments = CreateArguments("manage", docPath);
        arguments["revisionIndex"] = 0;
        arguments["action"] = "accept";

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("no revisions", result);
    }

    [Fact]
    public async Task CompareDocuments_ShouldCreateComparisonDocument()
    {
        // Arrange
        var originalPath = CreateWordDocumentWithContent("test_compare_original.docx", "Original content");
        var revisedPath = CreateWordDocumentWithContent("test_compare_revised.docx", "Revised content");
        var outputPath = CreateTestFilePath("test_compare_output.docx");

        var arguments = CreateArguments("compare", originalPath, outputPath);
        arguments["originalPath"] = originalPath;
        arguments["revisedPath"] = revisedPath;

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.True(File.Exists(outputPath));
        Assert.Contains("Comparison completed", result);
        Assert.Contains("difference(s) found", result);
    }

    [Fact]
    public async Task CompareDocuments_WithIgnoreFormatting_ShouldWork()
    {
        // Arrange
        var originalPath = CreateWordDocumentWithContent("test_compare_fmt_orig.docx", "Same content");
        var revisedPath = CreateWordDocumentWithContent("test_compare_fmt_rev.docx", "Same content");
        var outputPath = CreateTestFilePath("test_compare_fmt_output.docx");

        var arguments = CreateArguments("compare", originalPath, outputPath);
        arguments["originalPath"] = originalPath;
        arguments["revisedPath"] = revisedPath;
        arguments["ignoreFormatting"] = true;
        arguments["ignoreComments"] = true;

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.True(File.Exists(outputPath));
        Assert.Contains("Comparison completed", result);
    }
}