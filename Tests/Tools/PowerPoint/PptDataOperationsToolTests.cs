using System.Runtime.Versioning;
using Aspose.Slides;
using Aspose.Slides.Export;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.PowerPoint;

namespace AsposeMcpServer.Tests.Tools.PowerPoint;

[SupportedOSPlatform("windows")]
public class PptDataOperationsToolTests : TestBase
{
    private readonly PptDataOperationsTool _tool;

    public PptDataOperationsToolTests()
    {
        _tool = new PptDataOperationsTool(SessionManager);
    }

    private string CreateTestPresentation(string fileName, int slideCount = 2)
    {
        var filePath = CreateTestFilePath(fileName);
        using var presentation = new Presentation();
        for (var i = 1; i < slideCount; i++)
            presentation.Slides.AddEmptySlide(presentation.LayoutSlides[0]);
        presentation.Save(filePath, SaveFormat.Pptx);
        return filePath;
    }

    private string CreatePresentationWithHiddenSlide(string fileName)
    {
        var filePath = CreateTestFilePath(fileName);
        using var ppt = new Presentation();
        ppt.Slides.AddEmptySlide(ppt.LayoutSlides[0]);
        ppt.Slides[0].Hidden = true;
        ppt.Save(filePath, SaveFormat.Pptx);
        return filePath;
    }

    #region General

    [Fact]
    public void GetStatistics_ShouldReturnStatistics()
    {
        var pptPath = CreateTestPresentation("test_get_statistics.pptx");
        var result = _tool.Execute("get_statistics", pptPath);
        Assert.Contains("\"totalSlides\":", result);
        Assert.Contains("\"totalShapes\":", result);
        Assert.Contains("\"slideSize\":", result);
    }

    [Fact]
    public void GetStatistics_WithHiddenSlides_ShouldIncludeHiddenCount()
    {
        var pptPath = CreatePresentationWithHiddenSlide("test_hidden_slides.pptx");
        var result = _tool.Execute("get_statistics", pptPath);
        Assert.Contains("\"totalHiddenSlides\":", result);
    }

    [Fact]
    public void GetStatistics_ShouldIncludeAllCounters()
    {
        var pptPath = CreateTestPresentation("test_statistics_counters.pptx");
        var result = _tool.Execute("get_statistics", pptPath);
        Assert.Contains("\"totalSlides\":", result);
        Assert.Contains("\"totalLayouts\":", result);
        Assert.Contains("\"totalMasters\":", result);
        Assert.Contains("\"totalImages\":", result);
        Assert.Contains("\"totalTables\":", result);
        Assert.Contains("\"totalCharts\":", result);
        Assert.Contains("\"totalAnimations\":", result);
        Assert.Contains("\"totalHyperlinks\":", result);
    }

    [Fact]
    public void GetContent_ShouldReturnContent()
    {
        var pptPath = CreateTestPresentation("test_get_content.pptx");
        var result = _tool.Execute("get_content", pptPath);
        Assert.Contains("\"totalSlides\":", result);
        Assert.Contains("\"slides\":", result);
    }

    [Fact]
    public void GetContent_WithHiddenSlide_ShouldIncludeHiddenFlag()
    {
        var pptPath = CreatePresentationWithHiddenSlide("test_content_hidden.pptx");
        var result = _tool.Execute("get_content", pptPath);
        Assert.Contains("\"hidden\":", result);
        Assert.Contains("true", result);
    }

    [Fact]
    public void GetContent_ShouldIncludeTextContent()
    {
        var pptPath = CreateTestPresentation("test_content_text.pptx");
        var result = _tool.Execute("get_content", pptPath);
        Assert.Contains("\"textContent\":", result);
        Assert.Contains("\"index\":", result);
    }

    [Fact]
    public void GetSlideDetails_ShouldReturnSlideDetails()
    {
        var pptPath = CreateTestPresentation("test_get_slide_details.pptx");
        var result = _tool.Execute("get_slide_details", pptPath, slideIndex: 0);
        Assert.Contains("\"slideIndex\":", result);
        Assert.Contains("\"slideSize\":", result);
        Assert.Contains("\"shapesCount\":", result);
        Assert.Contains("\"layout\":", result);
    }

    [Fact]
    public void GetSlideDetails_ShouldIncludeTransitionInfo()
    {
        var pptPath = CreateTestPresentation("test_slide_transition.pptx");
        var result = _tool.Execute("get_slide_details", pptPath, slideIndex: 0);
        Assert.Contains("\"transition\":", result);
    }

    [Fact]
    public void GetSlideDetails_ShouldIncludeAnimationsInfo()
    {
        var pptPath = CreateTestPresentation("test_slide_animations.pptx");
        var result = _tool.Execute("get_slide_details", pptPath, slideIndex: 0);
        Assert.Contains("\"animationsCount\":", result);
        Assert.Contains("\"animations\":", result);
    }

    [Fact]
    public void GetSlideDetails_ShouldIncludeBackgroundInfo()
    {
        var pptPath = CreateTestPresentation("test_slide_background.pptx");
        var result = _tool.Execute("get_slide_details", pptPath, slideIndex: 0);
        Assert.Contains("\"background\":", result);
    }

    [Fact]
    public void GetSlideDetails_WithThumbnail_ShouldReturnBase64Image()
    {
        var pptPath = CreateTestPresentation("test_thumbnail.pptx");
        var result = _tool.Execute("get_slide_details", pptPath, slideIndex: 0, includeThumbnail: true);
        Assert.Contains("\"thumbnail\":", result);
        Assert.Contains("iVBORw0KGgo", result);
    }

    [Fact]
    public void GetSlideDetails_WithoutThumbnail_ShouldReturnNullThumbnail()
    {
        var pptPath = CreateTestPresentation("test_no_thumbnail.pptx");
        var result = _tool.Execute("get_slide_details", pptPath, slideIndex: 0, includeThumbnail: false);
        Assert.Contains("\"thumbnail\":", result);
        Assert.Contains("null", result);
    }

    [Fact]
    public void GetSlideDetails_ForSecondSlide_ShouldReturnCorrectIndex()
    {
        var pptPath = CreateTestPresentation("test_second_slide.pptx", 3);
        var result = _tool.Execute("get_slide_details", pptPath, slideIndex: 1);
        Assert.Contains("\"slideIndex\": 1", result);
    }

    [Theory]
    [InlineData("GET_STATISTICS")]
    [InlineData("Get_Statistics")]
    [InlineData("get_statistics")]
    public void Operation_ShouldBeCaseInsensitive_GetStatistics(string operation)
    {
        var pptPath = CreateTestPresentation($"test_case_stats_{operation.Replace("_", "")}.pptx");
        var result = _tool.Execute(operation, pptPath);
        Assert.Contains("\"totalSlides\":", result);
    }

    [Theory]
    [InlineData("GET_CONTENT")]
    [InlineData("Get_Content")]
    [InlineData("get_content")]
    public void Operation_ShouldBeCaseInsensitive_GetContent(string operation)
    {
        var pptPath = CreateTestPresentation($"test_case_content_{operation.Replace("_", "")}.pptx");
        var result = _tool.Execute(operation, pptPath);
        Assert.Contains("\"slides\":", result);
    }

    [Theory]
    [InlineData("GET_SLIDE_DETAILS")]
    [InlineData("Get_Slide_Details")]
    [InlineData("get_slide_details")]
    public void Operation_ShouldBeCaseInsensitive_GetSlideDetails(string operation)
    {
        var pptPath = CreateTestPresentation($"test_case_details_{operation.Replace("_", "")}.pptx");
        var result = _tool.Execute(operation, pptPath, slideIndex: 0);
        Assert.Contains("\"slideIndex\":", result);
    }

    #endregion

    #region Exception

    [Fact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        var pptPath = CreateTestPresentation("test_unknown_op.pptx");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("unknown", pptPath));
        Assert.Contains("Unknown operation", ex.Message);
    }

    [Fact]
    public void GetSlideDetails_WithoutSlideIndex_ShouldThrowArgumentException()
    {
        var pptPath = CreateTestPresentation("test_no_slide_index.pptx");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("get_slide_details", pptPath));
        Assert.Contains("slideIndex is required", ex.Message);
    }

    [Fact]
    public void GetSlideDetails_WithInvalidSlideIndex_ShouldThrowArgumentException()
    {
        var pptPath = CreateTestPresentation("test_invalid_slide.pptx");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("get_slide_details", pptPath, slideIndex: 99));
        Assert.Contains("slide", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void GetSlideDetails_WithNegativeSlideIndex_ShouldThrowArgumentException()
    {
        var pptPath = CreateTestPresentation("test_negative_slide.pptx");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("get_slide_details", pptPath, slideIndex: -1));
        Assert.Contains("slide", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    #endregion

    #region Session

    [Fact]
    public void GetStatistics_WithSessionId_ShouldReturnStatistics()
    {
        var pptPath = CreateTestPresentation("test_session_get_statistics.pptx");
        var sessionId = OpenSession(pptPath);
        var result = _tool.Execute("get_statistics", sessionId: sessionId);
        Assert.Contains("\"totalSlides\":", result);
    }

    [Fact]
    public void GetContent_WithSessionId_ShouldReturnContent()
    {
        var pptPath = CreateTestPresentation("test_session_get_content.pptx");
        var sessionId = OpenSession(pptPath);
        var result = _tool.Execute("get_content", sessionId: sessionId);
        Assert.Contains("\"slides\":", result);
    }

    [Fact]
    public void GetSlideDetails_WithSessionId_ShouldReturnSlideDetails()
    {
        var pptPath = CreateTestPresentation("test_session_get_slide_details.pptx");
        var sessionId = OpenSession(pptPath);
        var result = _tool.Execute("get_slide_details", sessionId: sessionId, slideIndex: 0);
        Assert.Contains("\"slideIndex\":", result);
        Assert.Contains("\"slideSize\":", result);
    }

    [Fact]
    public void GetSlideDetails_WithSessionId_WithThumbnail_ShouldWork()
    {
        var pptPath = CreateTestPresentation("test_session_thumbnail.pptx");
        var sessionId = OpenSession(pptPath);
        var result = _tool.Execute("get_slide_details", sessionId: sessionId, slideIndex: 0, includeThumbnail: true);
        Assert.Contains("\"thumbnail\":", result);
        Assert.Contains("iVBORw0KGgo", result);
    }

    [Fact]
    public void GetStatistics_WithSessionId_ShouldReflectInMemoryChanges()
    {
        var pptPath = CreateTestPresentation("test_session_statistics_changes.pptx");
        var sessionId = OpenSession(pptPath);
        var ppt = SessionManager.GetDocument<Presentation>(sessionId);
        var initialSlideCount = ppt.Slides.Count;
        ppt.Slides.AddEmptySlide(ppt.LayoutSlides[0]);
        var result = _tool.Execute("get_statistics", sessionId: sessionId);
        Assert.Contains("\"totalSlides\":", result);
        Assert.Contains((initialSlideCount + 1).ToString(), result);
    }

    [Fact]
    public void Execute_WithInvalidSessionId_ShouldThrowKeyNotFoundException()
    {
        Assert.Throws<KeyNotFoundException>(() => _tool.Execute("get_statistics", sessionId: "invalid_session"));
    }

    [Fact]
    public void Execute_WithBothPathAndSessionId_ShouldPreferSessionId()
    {
        var pptPath1 = CreateTestPresentation("test_path_data.pptx", 1);
        var pptPath2 = CreateTestPresentation("test_session_data.pptx", 5);
        var sessionId = OpenSession(pptPath2);
        var result = _tool.Execute("get_statistics", pptPath1, sessionId);
        Assert.Contains("\"totalSlides\": 5", result);
    }

    #endregion
}