using Aspose.Slides;
using Aspose.Slides.Export;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.PowerPoint;

namespace AsposeMcpServer.Tests.Tools.PowerPoint;

public class PptDataOperationsToolTests : TestBase
{
    private readonly PptDataOperationsTool _tool;

    public PptDataOperationsToolTests()
    {
        _tool = new PptDataOperationsTool(SessionManager);
    }

    private string CreateTestPresentation(string fileName)
    {
        var filePath = CreateTestFilePath(fileName);
        using var presentation = new Presentation();
        presentation.Slides.AddEmptySlide(presentation.LayoutSlides[0]);
        presentation.Save(filePath, SaveFormat.Pptx);
        return filePath;
    }

    #region General Tests

    [Fact]
    public void GetStatistics_ShouldReturnStatistics()
    {
        var pptPath = CreateTestPresentation("test_get_statistics.pptx");
        var result = _tool.Execute("get_statistics", pptPath);
        Assert.NotNull(result);
        Assert.NotEmpty(result);
        Assert.Contains("totalSlides", result);
    }

    [Fact]
    public void GetContent_ShouldReturnContent()
    {
        var pptPath = CreateTestPresentation("test_get_content.pptx");
        var result = _tool.Execute("get_content", pptPath);
        Assert.NotNull(result);
        Assert.NotEmpty(result);
        // Content may be empty for blank slides, just verify result is returned
        Assert.True(result.Length > 0, "Result should not be empty");
    }

    [Fact]
    public void GetSlideDetails_ShouldReturnSlideDetails()
    {
        var pptPath = CreateTestPresentation("test_get_slide_details.pptx");
        var result = _tool.Execute("get_slide_details", pptPath, slideIndex: 0);
        Assert.NotNull(result);
        Assert.NotEmpty(result);
        Assert.Contains("slideIndex", result);
        Assert.Contains("slideSize", result);
    }

    [Fact]
    public void GetStatistics_ShouldIncludeHiddenSlidesCount()
    {
        var pptPath = CreateTestFilePath("test_hidden_slides.pptx");
        using (var ppt = new Presentation())
        {
            ppt.Slides.AddEmptySlide(ppt.LayoutSlides[0]);
            ppt.Slides[0].Hidden = true;
            ppt.Save(pptPath, SaveFormat.Pptx);
        }

        var result = _tool.Execute("get_statistics", pptPath);

        Assert.Contains("totalHiddenSlides", result);
    }

    [Fact]
    public void GetContent_ShouldIncludeHiddenFlag()
    {
        var pptPath = CreateTestFilePath("test_content_hidden.pptx");
        using (var ppt = new Presentation())
        {
            ppt.Slides[0].Hidden = true;
            ppt.Save(pptPath, SaveFormat.Pptx);
        }

        var result = _tool.Execute("get_content", pptPath);

        Assert.Contains("hidden", result);
    }

    [Fact]
    public void GetSlideDetails_WithThumbnail_ShouldReturnBase64Image()
    {
        var pptPath = CreateTestPresentation("test_thumbnail.pptx");
        var result = _tool.Execute("get_slide_details", pptPath, slideIndex: 0, includeThumbnail: true);
        Assert.NotNull(result);
        Assert.Contains("thumbnail", result);
        // Base64 PNG starts with iVBORw0KGgo (PNG signature)
        Assert.Contains("iVBORw0KGgo", result);
    }

    [Fact]
    public void GetSlideDetails_WithoutThumbnail_ShouldNotIncludeThumbnailData()
    {
        var pptPath = CreateTestPresentation("test_no_thumbnail.pptx");
        var result = _tool.Execute("get_slide_details", pptPath, slideIndex: 0, includeThumbnail: false);
        Assert.NotNull(result);
        Assert.Contains("thumbnail", result);
        Assert.Contains("null", result); // thumbnail should be null
    }

    #endregion

    #region Exception Tests

    [Fact]
    public void Execute_UnknownOperation_ShouldThrowArgumentException()
    {
        var pptPath = CreateTestPresentation("test_unknown_op.pptx");

        Assert.Throws<ArgumentException>(() => _tool.Execute("unknown", pptPath));
    }

    [Fact]
    public void GetSlideDetails_InvalidSlideIndex_ShouldThrowArgumentException()
    {
        var pptPath = CreateTestPresentation("test_invalid_slide.pptx");

        Assert.Throws<ArgumentException>(() => _tool.Execute("get_slide_details", pptPath, slideIndex: 99));
    }

    #endregion

    #region Session ID Tests

    [Fact]
    public void GetStatistics_WithSessionId_ShouldReturnStatistics()
    {
        var pptPath = CreateTestPresentation("test_session_get_statistics.pptx");
        var sessionId = OpenSession(pptPath);
        var result = _tool.Execute("get_statistics", sessionId: sessionId);
        Assert.NotNull(result);
        Assert.Contains("totalSlides", result);
    }

    [Fact]
    public void GetContent_WithSessionId_ShouldReturnContent()
    {
        var pptPath = CreateTestPresentation("test_session_get_content.pptx");
        var sessionId = OpenSession(pptPath);
        var result = _tool.Execute("get_content", sessionId: sessionId);
        Assert.NotNull(result);
        Assert.True(result.Length > 0, "Result should not be empty");
    }

    [Fact]
    public void GetSlideDetails_WithSessionId_ShouldReturnSlideDetails()
    {
        var pptPath = CreateTestPresentation("test_session_get_slide_details.pptx");
        var sessionId = OpenSession(pptPath);
        var result = _tool.Execute("get_slide_details", sessionId: sessionId, slideIndex: 0);
        Assert.NotNull(result);
        Assert.Contains("slideIndex", result);
        Assert.Contains("slideSize", result);
    }

    [Fact]
    public void GetStatistics_WithSessionId_ShouldReflectInMemoryChanges()
    {
        var pptPath = CreateTestPresentation("test_session_statistics_changes.pptx");
        var sessionId = OpenSession(pptPath);

        var ppt = SessionManager.GetDocument<Presentation>(sessionId);
        var initialSlideCount = ppt.Slides.Count;

        // Add a slide in memory
        ppt.Slides.AddEmptySlide(ppt.LayoutSlides[0]);
        var result = _tool.Execute("get_statistics", sessionId: sessionId);
        Assert.Contains("totalSlides", result);
        Assert.Contains((initialSlideCount + 1).ToString(), result);
    }

    #endregion
}