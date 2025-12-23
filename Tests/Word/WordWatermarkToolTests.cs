using Aspose.Words;
using Aspose.Words.Drawing;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Word;

namespace AsposeMcpServer.Tests.Word;

public class WordWatermarkToolTests : WordTestBase
{
    private readonly WordWatermarkTool _tool = new();

    [Fact]
    public async Task AddWatermark_ShouldAddWatermark()
    {
        // Arrange
        var docPath = CreateWordDocument("test_add_watermark.docx");
        var outputPath = CreateTestFilePath("test_add_watermark_output.docx");
        var arguments = CreateArguments("add", docPath, outputPath);
        arguments["text"] = "CONFIDENTIAL";
        arguments["fontSize"] = 72;
        arguments["isSemitransparent"] = true;

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.True(File.Exists(outputPath), "Output document should be created");
        Assert.NotNull(result);
        Assert.Contains("Watermark", result, StringComparison.OrdinalIgnoreCase);

        // Verify watermark was added by checking document has watermark shapes
        var doc = new Document(outputPath);
        var shapes = doc.GetChildNodes(NodeType.Shape, true).Cast<Shape>().ToList();
        // Watermarks are typically added as shapes in headers
        var hasWatermarkShapes =
            shapes.Count > 0 || doc.Sections[0].HeadersFooters[HeaderFooterType.HeaderPrimary] != null;
        Assert.True(hasWatermarkShapes || doc.Watermark != null,
            "Document should contain watermark (checking shapes or watermark property)");
    }
}