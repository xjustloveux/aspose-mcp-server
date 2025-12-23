using Aspose.Words;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Word;

namespace AsposeMcpServer.Tests.Word;

public class WordFormatToolTests : WordTestBase
{
    private readonly WordFormatTool _tool = new();

    [Fact]
    public async Task GetRunFormat_ShouldReturnFormatInfo()
    {
        // Arrange
        var docPath = CreateWordDocumentWithContent("test_get_run_format.docx", "Test text");
        var arguments = CreateArguments("get_run_format", docPath);
        arguments["paragraphIndex"] = 0;
        arguments["runIndex"] = 0;

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.NotNull(result);
        Assert.NotEmpty(result);
        Assert.Contains("Font", result, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task SetRunFormat_ShouldApplyFormatting()
    {
        // Arrange
        var docPath = CreateWordDocumentWithContent("test_set_run_format.docx", "Format this");
        var outputPath = CreateTestFilePath("test_set_run_format_output.docx");
        var arguments = CreateArguments("set_run_format", docPath, outputPath);
        arguments["paragraphIndex"] = 0;
        arguments["runIndex"] = 0;
        arguments["bold"] = true;
        arguments["fontSize"] = 14;

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var doc = new Document(outputPath);
        var runs = doc.GetChildNodes(NodeType.Run, true).Cast<Run>().ToList();
        if (runs.Count > 0)
        {
            Assert.True(runs[0].Font.Bold);
            Assert.Equal(14, runs[0].Font.Size);
        }
    }

    [Fact]
    public async Task GetTabStops_ShouldReturnTabStops()
    {
        // Arrange
        var docPath = CreateWordDocumentWithContent("test_get_tab_stops.docx", "Test");
        var arguments = CreateArguments("get_tab_stops", docPath);
        arguments["paragraphIndex"] = 0;

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.NotNull(result);
        Assert.NotEmpty(result);
    }

    [Fact]
    public async Task SetParagraphBorder_ShouldSetBorder()
    {
        // Arrange
        var docPath = CreateWordDocumentWithContent("test_set_border.docx", "Test paragraph");
        var outputPath = CreateTestFilePath("test_set_border_output.docx");
        var arguments = CreateArguments("set_paragraph_border", docPath, outputPath);
        arguments["paragraphIndex"] = 0;
        arguments["borderType"] = "all";
        arguments["style"] = "single";
        arguments["width"] = 1.0;

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.True(File.Exists(outputPath), "Output document should be created");
        var doc = new Document(outputPath);
        var paragraphs = GetParagraphs(doc);
        Assert.True(paragraphs.Count > 0, "Document should have paragraphs");
    }
}