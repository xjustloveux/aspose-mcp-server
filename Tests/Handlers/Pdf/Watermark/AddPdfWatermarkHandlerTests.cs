using Aspose.Pdf;
using AsposeMcpServer.Handlers.Pdf.Watermark;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Pdf.Watermark;

public class AddPdfWatermarkHandlerTests : PdfHandlerTestBase
{
    private readonly AddPdfWatermarkHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Add()
    {
        Assert.Equal("add", _handler.Operation);
    }

    #endregion

    #region Helper Methods

    private static Document CreatePdfWithPages(int pageCount)
    {
        var document = new Document();
        for (var i = 0; i < pageCount; i++) document.Pages.Add();
        return document;
    }

    #endregion

    #region Basic Add Watermark Operations

    [SkippableFact]
    public void Execute_AddsWatermark()
    {
        SkipInEvaluationMode(AsposeLibraryType.Pdf);
        var document = CreateEmptyDocument();
        var context = CreateContext(document);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "CONFIDENTIAL" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        var page = document.Pages[1];
        Assert.True(page.Artifacts.Count > 0, "Page should have at least one artifact after adding watermark");
        var watermark = page.Artifacts.OfType<WatermarkArtifact>().First();
        Assert.NotNull(watermark);
        AssertModified(context);
    }

    [SkippableFact]
    public void Execute_AddsWatermarkWithCustomSettings()
    {
        SkipInEvaluationMode(AsposeLibraryType.Pdf);
        var document = CreateEmptyDocument();
        var context = CreateContext(document);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "DRAFT" },
            { "opacity", 0.5 },
            { "fontSize", 48.0 },
            { "rotation", 30.0 },
            { "color", "Red" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        var page = document.Pages[1];
        var watermark = page.Artifacts.OfType<WatermarkArtifact>().First();
        Assert.Equal(0.5, watermark.Opacity);
        Assert.Equal(30.0, watermark.Rotation);
        Assert.True(page.Artifacts.Count > 0, "Page should have at least one artifact after adding watermark");
    }

    [SkippableFact]
    public void Execute_AddsWatermarkWithAlignment()
    {
        SkipInEvaluationMode(AsposeLibraryType.Pdf);
        var document = CreateEmptyDocument();
        var context = CreateContext(document);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "SECRET" },
            { "horizontalAlignment", "Left" },
            { "verticalAlignment", "Top" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        var page = document.Pages[1];
        var watermark = page.Artifacts.OfType<WatermarkArtifact>().First();
        Assert.Equal(HorizontalAlignment.Left, watermark.ArtifactHorizontalAlignment);
        Assert.Equal(VerticalAlignment.Top, watermark.ArtifactVerticalAlignment);
    }

    [SkippableFact]
    public void Execute_AddsWatermarkAsBackground()
    {
        SkipInEvaluationMode(AsposeLibraryType.Pdf);
        var document = CreateEmptyDocument();
        var context = CreateContext(document);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "BACKGROUND" },
            { "isBackground", true }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        var page = document.Pages[1];
        var watermark = page.Artifacts.OfType<WatermarkArtifact>().First();
        Assert.True(watermark.IsBackground, "Watermark should be set as background");
    }

    [SkippableFact]
    public void Execute_AddsWatermarkToMultiplePages()
    {
        SkipInEvaluationMode(AsposeLibraryType.Pdf);
        var document = CreatePdfWithPages(3);
        var context = CreateContext(document);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "WATERMARK" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        for (var i = 1; i <= 3; i++)
        {
            var page = document.Pages[i];
            Assert.True(page.Artifacts.OfType<WatermarkArtifact>().Any(),
                $"Page {i} should have a watermark artifact");
        }
    }

    [SkippableFact]
    public void Execute_AddsWatermarkToPageRange()
    {
        SkipInEvaluationMode(AsposeLibraryType.Pdf);
        var document = CreatePdfWithPages(5);
        var context = CreateContext(document);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "WATERMARK" },
            { "pageRange", "1-3" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        for (var i = 1; i <= 3; i++)
        {
            var page = document.Pages[i];
            Assert.True(page.Artifacts.OfType<WatermarkArtifact>().Any(),
                $"Page {i} should have a watermark artifact");
        }

        for (var i = 4; i <= 5; i++)
        {
            var page = document.Pages[i];
            Assert.False(page.Artifacts.OfType<WatermarkArtifact>().Any(),
                $"Page {i} should not have a watermark artifact");
        }
    }

    [Fact]
    public void Execute_WithNoText_ThrowsArgumentException()
    {
        var document = CreateEmptyDocument();
        var context = CreateContext(document);
        var parameters = CreateEmptyParameters();

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithEmptyText_ThrowsArgumentException()
    {
        var document = CreateEmptyDocument();
        var context = CreateContext(document);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "" }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion
}
