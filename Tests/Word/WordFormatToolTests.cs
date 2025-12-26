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

    [Fact]
    public async Task AddTabStop_ShouldAddTabStop()
    {
        // Arrange
        var docPath = CreateWordDocumentWithContent("test_add_tab_stop.docx", "Test text with tab");
        var outputPath = CreateTestFilePath("test_add_tab_stop_output.docx");
        var arguments = CreateArguments("add_tab_stop", docPath, outputPath);
        arguments["paragraphIndex"] = 0;
        arguments["tabPosition"] = 72.0; // 1 inch in points
        arguments["alignment"] = "left";
        arguments["leader"] = "none";

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.True(File.Exists(outputPath), "Output document should be created");
        var doc = new Document(outputPath);
        var paragraphs = GetParagraphs(doc);
        Assert.True(paragraphs.Count > 0, "Document should have paragraphs");
        var tabStops = paragraphs[0].ParagraphFormat.TabStops;
        Assert.True(tabStops.Count > 0, "Paragraph should have at least one tab stop");
    }

    [Fact]
    public async Task ClearTabStops_ShouldClearTabStops()
    {
        // Arrange
        var docPath = CreateWordDocumentWithContent("test_clear_tab_stops.docx", "Test text");
        var doc = new Document(docPath);
        var paragraphs = GetParagraphs(doc);
        if (paragraphs.Count > 0)
        {
            paragraphs[0].ParagraphFormat.TabStops.Add(72.0, TabAlignment.Left, TabLeader.None);
            doc.Save(docPath);
        }

        var outputPath = CreateTestFilePath("test_clear_tab_stops_output.docx");
        var arguments = CreateArguments("clear_tab_stops", docPath, outputPath);
        arguments["paragraphIndex"] = 0;

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.True(File.Exists(outputPath), "Output document should be created");
        var resultDoc = new Document(outputPath);
        var resultParagraphs = GetParagraphs(resultDoc);
        Assert.True(resultParagraphs.Count > 0, "Document should have paragraphs");
        var tabStops = resultParagraphs[0].ParagraphFormat.TabStops;
        Assert.Equal(0, tabStops.Count);
    }
}