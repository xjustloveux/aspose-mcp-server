using Aspose.Words;
using AsposeMcpServer.Results;
using AsposeMcpServer.Results.Word.Render;
using AsposeMcpServer.Tests.Infrastructure;
using AsposeMcpServer.Tools.Word;

namespace AsposeMcpServer.Tests.Tools.Word;

/// <summary>
///     Integration tests for WordRenderTool.
/// </summary>
public class WordRenderToolTests : WordTestBase
{
    private readonly WordRenderTool _tool;

    public WordRenderToolTests()
    {
        _tool = new WordRenderTool(SessionManager);
    }

    #region File I/O Smoke Tests

    [Fact]
    public void RenderPage_ShouldRenderAndPersist()
    {
        var docPath = CreateWordDocumentWithContent("test_render_tool.docx", "Render test content");
        var outputPath = CreateTestFilePath("test_render_output.png");
        var result = _tool.Execute("render_page", docPath, outputPath, 1);
        var data = GetResultData<RenderResult>(result);
        Assert.Contains("Page 1", data.Message);
        Assert.Single(data.OutputPaths);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void RenderThumbnail_ShouldRenderAndPersist()
    {
        var docPath = CreateWordDocumentWithContent("test_thumb_tool.docx", "Thumbnail test content");
        var outputPath = CreateTestFilePath("test_thumb_output.png");
        var result = _tool.Execute("render_thumbnail", docPath, outputPath);
        var data = GetResultData<RenderResult>(result);
        Assert.Contains("Thumbnail", data.Message);
        Assert.Single(data.OutputPaths);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void RenderPage_WithJpeg_ShouldRenderJpeg()
    {
        var docPath = CreateWordDocumentWithContent("test_render_jpeg.docx", "JPEG render");
        var outputPath = CreateTestFilePath("test_render_output.jpeg");
        var result = _tool.Execute("render_page", docPath, outputPath, 1, "jpeg");
        var data = GetResultData<RenderResult>(result);
        Assert.Equal("jpeg", data.Format);
    }

    [Fact]
    public void RenderPage_AllPages_ShouldRenderMultiple()
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Write("Page 1");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Write("Page 2");
        var docPath = CreateTestFilePath("test_render_multi.docx");
        doc.Save(docPath);

        var outputPath = CreateTestFilePath("test_render_all.png");
        var result = _tool.Execute("render_page", docPath, outputPath);
        var data = GetResultData<RenderResult>(result);
        Assert.True(data.OutputPaths.Count >= 2);
    }

    [Fact]
    public void RenderPage_WithCustomDpi_ShouldSucceed()
    {
        var docPath = CreateWordDocumentWithContent("test_render_dpi.docx", "DPI test");
        var outputPath = CreateTestFilePath("test_render_dpi.png");
        var result = _tool.Execute("render_page", docPath, outputPath, 1, dpi: 300);
        var data = GetResultData<RenderResult>(result);
        Assert.NotNull(data);
    }

    #endregion

    #region Operation Routing

    [Theory]
    [InlineData("RENDER_PAGE")]
    [InlineData("Render_Page")]
    [InlineData("render_page")]
    public void Operation_ShouldBeCaseInsensitive(string operation)
    {
        var docPath = CreateWordDocumentWithContent($"test_render_case_{operation}.docx", "Test");
        var outputPath = CreateTestFilePath($"test_render_case_{operation}_output.png");
        var result = _tool.Execute(operation, docPath, outputPath, 1);
        Assert.IsType<FinalizedResult<RenderResult>>(result);
    }

    [Fact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocumentWithContent("test_render_unknown.docx", "Test");
        var outputPath = CreateTestFilePath("test_render_unknown_output.png");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("unknown_operation", docPath, outputPath));
        Assert.Contains("Unknown operation", ex.Message);
    }

    [Fact]
    public void Execute_WithNoPath_ShouldThrowException()
    {
        Assert.ThrowsAny<Exception>(() =>
            _tool.Execute("render_page", outputPath: "output.png"));
    }

    #endregion
}
