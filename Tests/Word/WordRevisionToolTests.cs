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
        Assert.Contains("Revision", result, StringComparison.OrdinalIgnoreCase);
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
        await _tool.ExecuteAsync(arguments);

        // Assert
        var resultDoc = new Document(outputPath);
        var revisionsAfter = resultDoc.Revisions.Count;
        Assert.True(revisionsAfter < revisionsBefore,
            $"Revisions should be accepted. Before: {revisionsBefore}, After: {revisionsAfter}");
    }
}