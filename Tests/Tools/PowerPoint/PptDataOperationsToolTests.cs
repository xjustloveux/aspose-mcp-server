using System.Runtime.Versioning;
using AsposeMcpServer.Results.PowerPoint.DataOperations;
using AsposeMcpServer.Tests.Infrastructure;
using AsposeMcpServer.Tools.PowerPoint;

namespace AsposeMcpServer.Tests.Tools.PowerPoint;

/// <summary>
///     Integration tests for PptDataOperationsTool.
///     Focuses on session management, file I/O, and operation routing.
///     Detailed parameter validation and business logic tests are in Handler tests.
/// </summary>
[SupportedOSPlatform("windows")]
public class PptDataOperationsToolTests : PptTestBase
{
    private readonly PptDataOperationsTool _tool;

    public PptDataOperationsToolTests()
    {
        _tool = new PptDataOperationsTool(SessionManager);
    }

    #region File I/O Smoke Tests

    [Fact]
    public void GetStatistics_ShouldReturnStatistics()
    {
        var pptPath = CreatePresentation("test_get_statistics.pptx");
        var result = _tool.Execute("get_statistics", pptPath);
        var data = GetResultData<GetStatisticsResult>(result);
        Assert.True(data.TotalSlides >= 0);
        Assert.True(data.TotalShapes >= 0);
        Assert.NotNull(data.SlideSize);
    }

    [Fact]
    public void GetContent_ShouldReturnContent()
    {
        var pptPath = CreatePresentation("test_get_content.pptx");
        var result = _tool.Execute("get_content", pptPath);
        var data = GetResultData<GetContentPptResult>(result);
        Assert.True(data.TotalSlides >= 0);
        Assert.NotNull(data.Slides);
    }

    [Fact]
    public void GetSlideDetails_ShouldReturnSlideDetails()
    {
        var pptPath = CreatePresentation("test_get_slide_details.pptx");
        var result = _tool.Execute("get_slide_details", pptPath, slideIndex: 0);
        var data = GetResultData<GetSlideDetailsResult>(result);
        Assert.Equal(0, data.SlideIndex);
        Assert.NotNull(data.SlideSize);
        Assert.True(data.ShapesCount >= 0);
    }

    #endregion

    #region Operation Routing

    [Theory]
    [InlineData("GET_STATISTICS")]
    [InlineData("Get_Statistics")]
    [InlineData("get_statistics")]
    public void Operation_ShouldBeCaseInsensitive(string operation)
    {
        var pptPath = CreatePresentation($"test_case_stats_{operation.Replace("_", "")}.pptx");
        var result = _tool.Execute(operation, pptPath);
        var data = GetResultData<GetStatisticsResult>(result);
        Assert.True(data.TotalSlides >= 0);
    }

    [Fact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        var pptPath = CreatePresentation("test_unknown_op.pptx");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("unknown", pptPath));
        Assert.Contains("Unknown operation", ex.Message);
    }

    #endregion

    #region Session Management

    [Fact]
    public void GetStatistics_WithSessionId_ShouldReturnStatistics()
    {
        var pptPath = CreatePresentation("test_session_get_statistics.pptx");
        var sessionId = OpenSession(pptPath);
        var result = _tool.Execute("get_statistics", sessionId: sessionId);
        var data = GetResultData<GetStatisticsResult>(result);
        Assert.True(data.TotalSlides >= 0);
    }

    [Fact]
    public void GetContent_WithSessionId_ShouldReturnContent()
    {
        var pptPath = CreatePresentation("test_session_get_content.pptx");
        var sessionId = OpenSession(pptPath);
        var result = _tool.Execute("get_content", sessionId: sessionId);
        var data = GetResultData<GetContentPptResult>(result);
        Assert.NotNull(data.Slides);
    }

    [Fact]
    public void GetSlideDetails_WithSessionId_ShouldReturnSlideDetails()
    {
        var pptPath = CreatePresentation("test_session_get_slide_details.pptx");
        var sessionId = OpenSession(pptPath);
        var result = _tool.Execute("get_slide_details", sessionId: sessionId, slideIndex: 0);
        var data = GetResultData<GetSlideDetailsResult>(result);
        Assert.Equal(0, data.SlideIndex);
        Assert.NotNull(data.SlideSize);
    }

    [Fact]
    public void Execute_WithInvalidSessionId_ShouldThrowKeyNotFoundException()
    {
        Assert.Throws<KeyNotFoundException>(() => _tool.Execute("get_statistics", sessionId: "invalid_session"));
    }

    [Fact]
    public void Execute_WithBothPathAndSessionId_ShouldPreferSessionId()
    {
        var pptPath1 = CreatePresentation("test_path_data.pptx");
        var pptPath2 = CreatePresentation("test_session_data.pptx", 5);
        var sessionId = OpenSession(pptPath2);
        var result = _tool.Execute("get_statistics", pptPath1, sessionId);
        var data = GetResultData<GetStatisticsResult>(result);
        Assert.Equal(5, data.TotalSlides);
    }

    #endregion
}
