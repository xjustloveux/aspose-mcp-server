using AsposeMcpServer.Handlers.Word.Render;
using AsposeMcpServer.Results.Word.Render;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Word.Render;

/// <summary>
///     Tests for RenderThumbnailWordHandler.
/// </summary>
public class RenderThumbnailWordHandlerTests : WordHandlerTestBase
{
    private readonly RenderThumbnailWordHandler _handler = new();

    [Fact]
    public void Operation_ShouldBeRenderThumbnail()
    {
        Assert.Equal("render_thumbnail", _handler.Operation);
    }

    [Fact]
    public void Execute_WithDefaultScale_ShouldRenderThumbnail()
    {
        var doc = CreateDocumentWithText("Thumbnail test content");
        var docPath = Path.Combine(TestDir, "test_thumb.docx");
        doc.Save(docPath);

        var outputPath = Path.Combine(TestDir, "test_thumb.png");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", docPath },
            { "outputPath", outputPath },
            { "format", "png" },
            { "scale", 0.25 }
        });

        var result = _handler.Execute(context, parameters);

        var renderResult = Assert.IsType<RenderResult>(result);
        Assert.Contains("25%", renderResult.Message);
        Assert.Contains("PNG", renderResult.Message);
        Assert.Single(renderResult.OutputPaths);
        Assert.True(System.IO.File.Exists(outputPath));
    }

    [Fact]
    public void Execute_WithJpegFormat_ShouldRenderJpeg()
    {
        var doc = CreateDocumentWithText("JPEG thumbnail");
        var docPath = Path.Combine(TestDir, "test_thumb_jpeg.docx");
        doc.Save(docPath);

        var outputPath = Path.Combine(TestDir, "test_thumb.jpeg");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", docPath },
            { "outputPath", outputPath },
            { "format", "jpeg" },
            { "scale", 0.5 }
        });

        var result = _handler.Execute(context, parameters);

        var renderResult = Assert.IsType<RenderResult>(result);
        Assert.Contains("50%", renderResult.Message);
        Assert.Equal("jpeg", renderResult.Format);
    }

    [Fact]
    public void Execute_WithCustomScale_ShouldUseScale()
    {
        var doc = CreateDocumentWithText("Custom scale");
        var docPath = Path.Combine(TestDir, "test_thumb_scale.docx");
        doc.Save(docPath);

        var outputPath = Path.Combine(TestDir, "test_thumb_custom.png");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", docPath },
            { "outputPath", outputPath },
            { "format", "png" },
            { "scale", 1.0 }
        });

        var result = _handler.Execute(context, parameters);

        var renderResult = Assert.IsType<RenderResult>(result);
        Assert.Contains("100%", renderResult.Message);
    }

    [Fact]
    public void Execute_WithZeroScale_ShouldThrowArgumentException()
    {
        var doc = CreateDocumentWithText("Zero scale");
        var docPath = Path.Combine(TestDir, "test_thumb_zero.docx");
        doc.Save(docPath);

        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", docPath },
            { "outputPath", Path.Combine(TestDir, "output.png") },
            { "format", "png" },
            { "scale", 0.0 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithNegativeScale_ShouldThrowArgumentException()
    {
        var doc = CreateDocumentWithText("Negative scale");
        var docPath = Path.Combine(TestDir, "test_thumb_neg.docx");
        doc.Save(docPath);

        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", docPath },
            { "outputPath", Path.Combine(TestDir, "output.png") },
            { "format", "png" },
            { "scale", -0.5 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithScaleGreaterThanOne_ShouldThrowArgumentException()
    {
        var doc = CreateDocumentWithText("Large scale");
        var docPath = Path.Combine(TestDir, "test_thumb_large.docx");
        doc.Save(docPath);

        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", docPath },
            { "outputPath", Path.Combine(TestDir, "output.png") },
            { "format", "png" },
            { "scale", 1.5 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithUnsupportedFormat_ShouldThrowArgumentException()
    {
        var doc = CreateDocumentWithText("Bad format");
        var docPath = Path.Combine(TestDir, "test_thumb_fmt.docx");
        doc.Save(docPath);

        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", docPath },
            { "outputPath", Path.Combine(TestDir, "output.xyz") },
            { "format", "bmp" },
            { "scale", 0.25 }
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
            { "scale", 0.25 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }
}
