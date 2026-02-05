using Aspose.Words;
using AsposeMcpServer.Handlers.Word.Render;
using AsposeMcpServer.Results.Word.Render;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Word.Render;

/// <summary>
///     Tests for RenderPageWordHandler.
/// </summary>
public class RenderPageWordHandlerTests : WordHandlerTestBase
{
    private readonly RenderPageWordHandler _handler = new();

    [Fact]
    public void Operation_ShouldBeRenderPage()
    {
        Assert.Equal("render_page", _handler.Operation);
    }

    [Fact]
    public void Execute_WithSinglePage_ShouldRenderToImage()
    {
        var doc = CreateDocumentWithText("Test page content");
        var docPath = Path.Combine(TestDir, "test_render_single.docx");
        doc.Save(docPath);

        var outputPath = Path.Combine(TestDir, "test_render_page1.png");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", docPath },
            { "outputPath", outputPath },
            { "pageIndex", 1 },
            { "format", "png" },
            { "dpi", 150 }
        });

        var result = _handler.Execute(context, parameters);

        var renderResult = Assert.IsType<RenderResult>(result);
        Assert.Contains("Page 1", renderResult.Message);
        Assert.Contains("PNG", renderResult.Message);
        Assert.Single(renderResult.OutputPaths);
        Assert.True(System.IO.File.Exists(outputPath));
    }

    [Fact]
    public void Execute_WithAllPages_ShouldRenderMultipleImages()
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Write("Page 1 content");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Write("Page 2 content");

        var docPath = Path.Combine(TestDir, "test_render_all.docx");
        doc.Save(docPath);

        var outputPath = Path.Combine(TestDir, "test_render_all.png");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", docPath },
            { "outputPath", outputPath },
            { "format", "png" },
            { "dpi", 150 }
        });

        var result = _handler.Execute(context, parameters);

        var renderResult = Assert.IsType<RenderResult>(result);
        Assert.Contains("page(s) rendered", renderResult.Message);
        Assert.True(renderResult.OutputPaths.Count >= 2);
    }

    [Fact]
    public void Execute_WithJpegFormat_ShouldRenderJpeg()
    {
        var doc = CreateDocumentWithText("JPEG test");
        var docPath = Path.Combine(TestDir, "test_render_jpeg.docx");
        doc.Save(docPath);

        var outputPath = Path.Combine(TestDir, "test_render_page.jpeg");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", docPath },
            { "outputPath", outputPath },
            { "pageIndex", 1 },
            { "format", "jpeg" },
            { "dpi", 150 }
        });

        var result = _handler.Execute(context, parameters);

        var renderResult = Assert.IsType<RenderResult>(result);
        Assert.Contains("JPEG", renderResult.Message);
        Assert.Equal("jpeg", renderResult.Format);
    }

    [Fact]
    public void Execute_WithInvalidPageIndex_ShouldThrowArgumentException()
    {
        var doc = CreateDocumentWithText("Single page");
        var docPath = Path.Combine(TestDir, "test_render_invalid.docx");
        doc.Save(docPath);

        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", docPath },
            { "outputPath", Path.Combine(TestDir, "output.png") },
            { "pageIndex", 99 },
            { "format", "png" },
            { "dpi", 150 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithZeroPageIndex_ShouldThrowArgumentException()
    {
        var doc = CreateDocumentWithText("Single page");
        var docPath = Path.Combine(TestDir, "test_render_zero.docx");
        doc.Save(docPath);

        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", docPath },
            { "outputPath", Path.Combine(TestDir, "output.png") },
            { "pageIndex", 0 },
            { "format", "png" },
            { "dpi", 150 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithUnknownFormat_ShouldThrowArgumentException()
    {
        var doc = CreateDocumentWithText("Test");
        var docPath = Path.Combine(TestDir, "test_render_fmt.docx");
        doc.Save(docPath);

        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", docPath },
            { "outputPath", Path.Combine(TestDir, "output.xyz") },
            { "pageIndex", 1 },
            { "format", "xyz" },
            { "dpi", 150 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithMissingPath_ShouldThrowArgumentException()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "outputPath", "output.png" },
            { "format", "png" },
            { "dpi", 150 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Theory]
    [InlineData("bmp")]
    [InlineData("tiff")]
    [InlineData("svg")]
    public void Execute_WithVariousFormats_ShouldSucceed(string format)
    {
        var doc = CreateDocumentWithText($"Test {format}");
        var docPath = Path.Combine(TestDir, $"test_render_{format}.docx");
        doc.Save(docPath);

        var outputPath = Path.Combine(TestDir, $"test_render.{format}");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", docPath },
            { "outputPath", outputPath },
            { "pageIndex", 1 },
            { "format", format },
            { "dpi", 150 }
        });

        var result = _handler.Execute(context, parameters);

        var renderResult = Assert.IsType<RenderResult>(result);
        Assert.Equal(format, renderResult.Format);
    }
}
