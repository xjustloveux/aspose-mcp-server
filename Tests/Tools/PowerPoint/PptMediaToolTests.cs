using Aspose.Slides;
using Aspose.Slides.Export;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;
using AsposeMcpServer.Tools.PowerPoint;

namespace AsposeMcpServer.Tests.Tools.PowerPoint;

/// <summary>
///     Integration tests for PptMediaTool.
///     Focuses on session management, file I/O, and operation routing.
///     Detailed parameter validation and business logic tests are in Handler tests.
/// </summary>
public class PptMediaToolTests : PptTestBase
{
    private readonly PptMediaTool _tool;

    public PptMediaToolTests()
    {
        _tool = new PptMediaTool(SessionManager);
    }

    private string CreateFakeAudioFile(string fileName)
    {
        var audioPath = CreateTestFilePath(fileName);
        File.WriteAllBytes(audioPath, "ID3"u8.ToArray());
        return audioPath;
    }

    private string CreateFakeVideoFile(string fileName)
    {
        var videoPath = CreateTestFilePath(fileName);
        File.WriteAllBytes(videoPath,
        [
            0x00, 0x00, 0x00, 0x14, 0x66, 0x74, 0x79, 0x70, 0x69, 0x73, 0x6F, 0x6D, 0x00, 0x00, 0x00, 0x00, 0x69, 0x73,
            0x6F, 0x6D
        ]);
        return videoPath;
    }

    private string CreatePresentationWithAudio(string fileName)
    {
        var filePath = CreateTestFilePath(fileName);
        var audioPath = CreateFakeAudioFile($"audio_for_{Path.GetFileNameWithoutExtension(fileName)}.mp3");
        using var presentation = new Presentation();
        using var audioStream = File.OpenRead(audioPath);
        presentation.Slides[0].Shapes.AddAudioFrameEmbedded(100, 100, 80, 80, audioStream);
        presentation.Save(filePath, SaveFormat.Pptx);
        return filePath;
    }

    private string CreatePresentationWithVideo(string fileName)
    {
        var filePath = CreateTestFilePath(fileName);
        var videoPath = CreateFakeVideoFile($"video_for_{Path.GetFileNameWithoutExtension(fileName)}.mp4");
        using var presentation = new Presentation();
        using var videoStream = File.OpenRead(videoPath);
        var videoObj = presentation.Videos.AddVideo(videoStream, LoadingStreamBehavior.KeepLocked);
        presentation.Slides[0].Shapes.AddVideoFrame(100, 100, 320, 240, videoObj);
        presentation.Save(filePath, SaveFormat.Pptx);
        return filePath;
    }

    #region File I/O Smoke Tests

    [Fact]
    public void AddAudio_ShouldEmbedAudio()
    {
        var pptPath = CreatePresentation("test_add_audio.pptx");
        var audioPath = CreateFakeAudioFile("test_audio.mp3");
        var outputPath = CreateTestFilePath("test_add_audio_output.pptx");
        var result = _tool.Execute("add_audio", pptPath, slideIndex: 0, audioPath: audioPath, x: 100, y: 100,
            outputPath: outputPath);
        var data = GetResultData<SuccessResult>(result);
        Assert.StartsWith("Audio embedded into slide", data.Message);
        using var presentation = new Presentation(outputPath);
        Assert.NotEmpty(presentation.Slides[0].Shapes.OfType<IAudioFrame>());
    }

    [Fact]
    public void DeleteAudio_ShouldRemoveAudioFrame()
    {
        var pptPath = CreatePresentationWithAudio("test_delete_audio.pptx");
        var outputPath = CreateTestFilePath("test_delete_audio_output.pptx");
        var result = _tool.Execute("delete_audio", pptPath, slideIndex: 0, shapeIndex: 0, outputPath: outputPath);
        var data = GetResultData<SuccessResult>(result);
        Assert.StartsWith("Audio deleted from slide", data.Message);
        using var presentation = new Presentation(outputPath);
        Assert.Empty(presentation.Slides[0].Shapes.OfType<IAudioFrame>());
    }

    [Fact]
    public void AddVideo_ShouldEmbedVideo()
    {
        var pptPath = CreatePresentation("test_add_video.pptx");
        var videoPath = CreateFakeVideoFile("test_video.mp4");
        var outputPath = CreateTestFilePath("test_add_video_output.pptx");
        var result = _tool.Execute("add_video", pptPath, slideIndex: 0, videoPath: videoPath, x: 100, y: 100,
            width: 320, height: 240, outputPath: outputPath);
        var data = GetResultData<SuccessResult>(result);
        Assert.StartsWith("Video embedded into slide", data.Message);
        using var presentation = new Presentation(outputPath);
        Assert.NotEmpty(presentation.Slides[0].Shapes.OfType<IVideoFrame>());
    }

    [Fact]
    public void DeleteVideo_ShouldRemoveVideoFrame()
    {
        var pptPath = CreatePresentationWithVideo("test_delete_video.pptx");
        var outputPath = CreateTestFilePath("test_delete_video_output.pptx");
        var result = _tool.Execute("delete_video", pptPath, slideIndex: 0, shapeIndex: 0, outputPath: outputPath);
        var data = GetResultData<SuccessResult>(result);
        Assert.StartsWith("Video deleted from slide", data.Message);
        using var presentation = new Presentation(outputPath);
        Assert.Empty(presentation.Slides[0].Shapes.OfType<IVideoFrame>());
    }

    [Fact]
    public void SetPlayback_ForAudio_ShouldUpdateSettings()
    {
        var pptPath = CreatePresentationWithAudio("test_playback_audio.pptx");
        var outputPath = CreateTestFilePath("test_playback_audio_output.pptx");
        var result = _tool.Execute("set_playback", pptPath, slideIndex: 0, shapeIndex: 0, playMode: "onclick",
            loop: true, volume: "loud", outputPath: outputPath);
        var data = GetResultData<SuccessResult>(result);
        Assert.StartsWith("Playback settings updated", data.Message);
    }

    #endregion

    #region Operation Routing

    [Theory]
    [InlineData("ADD_AUDIO")]
    [InlineData("Add_Audio")]
    [InlineData("add_audio")]
    public void Operation_ShouldBeCaseInsensitive(string operation)
    {
        var pptPath = CreatePresentation($"test_case_add_audio_{operation.Replace("_", "")}.pptx");
        var audioPath = CreateFakeAudioFile($"test_case_audio_{operation.Replace("_", "")}.mp3");
        var outputPath = CreateTestFilePath($"test_case_add_audio_{operation.Replace("_", "")}_output.pptx");
        var result = _tool.Execute(operation, pptPath, slideIndex: 0, audioPath: audioPath, outputPath: outputPath);
        var data = GetResultData<SuccessResult>(result);
        Assert.StartsWith("Audio embedded into slide", data.Message);
    }

    [Fact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        var pptPath = CreatePresentation("test_unknown_op.pptx");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("unknown", pptPath, slideIndex: 0));
        Assert.Contains("Unknown operation", ex.Message);
    }

    #endregion

    #region Session Management

    [Fact]
    public void AddAudio_WithSessionId_ShouldAddInMemory()
    {
        var pptPath = CreatePresentation("test_session_add_audio.pptx");
        var audioPath = CreateFakeAudioFile("session_audio.mp3");
        var sessionId = OpenSession(pptPath);
        var ppt = SessionManager.GetDocument<Presentation>(sessionId);
        var initialCount = ppt.Slides[0].Shapes.OfType<IAudioFrame>().Count();
        var result = _tool.Execute("add_audio", sessionId: sessionId, slideIndex: 0, audioPath: audioPath, x: 100,
            y: 100);
        var data = GetResultData<SuccessResult>(result);
        Assert.StartsWith("Audio embedded into slide", data.Message);
        Assert.True(ppt.Slides[0].Shapes.OfType<IAudioFrame>().Count() > initialCount);
        var output = GetResultOutput<SuccessResult>(result);
        Assert.True(output.IsSession);
    }

    [Fact]
    public void AddVideo_WithSessionId_ShouldAddInMemory()
    {
        var pptPath = CreatePresentation("test_session_add_video.pptx");
        var videoPath = CreateFakeVideoFile("session_video.mp4");
        var sessionId = OpenSession(pptPath);
        var ppt = SessionManager.GetDocument<Presentation>(sessionId);
        var initialCount = ppt.Slides[0].Shapes.OfType<IVideoFrame>().Count();
        var result = _tool.Execute("add_video", sessionId: sessionId, slideIndex: 0, videoPath: videoPath, x: 100,
            y: 100, width: 320, height: 240);
        var data = GetResultData<SuccessResult>(result);
        Assert.StartsWith("Video embedded into slide", data.Message);
        Assert.True(ppt.Slides[0].Shapes.OfType<IVideoFrame>().Count() > initialCount);
        var output = GetResultOutput<SuccessResult>(result);
        Assert.True(output.IsSession);
    }

    [Fact]
    public void DeleteAudio_WithSessionId_ShouldDeleteInMemory()
    {
        var pptPath = CreatePresentationWithAudio("test_session_delete_audio.pptx");
        var sessionId = OpenSession(pptPath);
        var ppt = SessionManager.GetDocument<Presentation>(sessionId);
        Assert.NotEmpty(ppt.Slides[0].Shapes.OfType<IAudioFrame>());
        var result = _tool.Execute("delete_audio", sessionId: sessionId, slideIndex: 0, shapeIndex: 0);
        var data = GetResultData<SuccessResult>(result);
        Assert.StartsWith("Audio deleted from slide", data.Message);
        Assert.Empty(ppt.Slides[0].Shapes.OfType<IAudioFrame>());
        var output = GetResultOutput<SuccessResult>(result);
        Assert.True(output.IsSession);
    }

    [Fact]
    public void Execute_WithInvalidSessionId_ShouldThrowKeyNotFoundException()
    {
        Assert.Throws<KeyNotFoundException>(() =>
            _tool.Execute("add_audio", sessionId: "invalid_session", slideIndex: 0));
    }

    [Fact]
    public void Execute_WithBothPathAndSessionId_ShouldPreferSessionId()
    {
        var pptPath1 = CreatePresentation("test_path_media.pptx");
        var pptPath2 = CreatePresentation("test_session_media.pptx");
        var audioPath = CreateFakeAudioFile("preference_audio.mp3");
        var sessionId = OpenSession(pptPath2);
        var ppt = SessionManager.GetDocument<Presentation>(sessionId);
        var result = _tool.Execute("add_audio", pptPath1, sessionId, slideIndex: 0, audioPath: audioPath, x: 100,
            y: 100);
        Assert.NotEmpty(ppt.Slides[0].Shapes.OfType<IAudioFrame>());
        var output = GetResultOutput<SuccessResult>(result);
        Assert.True(output.IsSession);
    }

    #endregion
}
