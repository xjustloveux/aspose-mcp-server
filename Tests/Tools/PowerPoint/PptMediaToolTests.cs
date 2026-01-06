using Aspose.Slides;
using Aspose.Slides.Export;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.PowerPoint;

namespace AsposeMcpServer.Tests.Tools.PowerPoint;

public class PptMediaToolTests : TestBase
{
    private readonly PptMediaTool _tool;

    public PptMediaToolTests()
    {
        _tool = new PptMediaTool(SessionManager);
    }

    private string CreateTestPresentation(string fileName)
    {
        var filePath = CreateTestFilePath(fileName);
        using var presentation = new Presentation();
        presentation.Slides.AddEmptySlide(presentation.LayoutSlides[0]);
        presentation.Save(filePath, SaveFormat.Pptx);
        return filePath;
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
        File.WriteAllBytes(videoPath, [0x00, 0x00, 0x00, 0x1C, 0x66, 0x74, 0x79, 0x70]);
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

    private string CreatePresentationWithShape(string fileName)
    {
        var filePath = CreateTestFilePath(fileName);
        using var presentation = new Presentation();
        presentation.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
        presentation.Save(filePath, SaveFormat.Pptx);
        return filePath;
    }

    #region General

    [Fact]
    public void AddAudio_ShouldEmbedAudio()
    {
        var pptPath = CreateTestPresentation("test_add_audio.pptx");
        var audioPath = CreateFakeAudioFile("test_audio.mp3");
        var outputPath = CreateTestFilePath("test_add_audio_output.pptx");
        var result = _tool.Execute("add_audio", pptPath, slideIndex: 0, audioPath: audioPath, x: 100, y: 100,
            outputPath: outputPath);
        Assert.StartsWith("Audio embedded into slide", result);
        using var presentation = new Presentation(outputPath);
        Assert.NotEmpty(presentation.Slides[0].Shapes.OfType<IAudioFrame>());
    }

    [Fact]
    public void AddAudio_WithCustomSize_ShouldSetDimensions()
    {
        var pptPath = CreateTestPresentation("test_add_audio_size.pptx");
        var audioPath = CreateFakeAudioFile("test_audio_size.mp3");
        var outputPath = CreateTestFilePath("test_add_audio_size_output.pptx");
        var result = _tool.Execute("add_audio", pptPath, slideIndex: 0, audioPath: audioPath, x: 50, y: 50, width: 100,
            height: 100, outputPath: outputPath);
        Assert.StartsWith("Audio embedded into slide", result);
        using var presentation = new Presentation(outputPath);
        var audioFrame = presentation.Slides[0].Shapes.OfType<IAudioFrame>().First();
        Assert.Equal(100, audioFrame.Width, 1);
        Assert.Equal(100, audioFrame.Height, 1);
    }

    [Fact]
    public void DeleteAudio_ShouldRemoveAudioFrame()
    {
        var pptPath = CreatePresentationWithAudio("test_delete_audio.pptx");
        var outputPath = CreateTestFilePath("test_delete_audio_output.pptx");
        var result = _tool.Execute("delete_audio", pptPath, slideIndex: 0, shapeIndex: 0, outputPath: outputPath);
        Assert.StartsWith("Audio deleted from slide", result);
        using var presentation = new Presentation(outputPath);
        Assert.Empty(presentation.Slides[0].Shapes.OfType<IAudioFrame>());
    }

    [Fact]
    public void AddVideo_ShouldEmbedVideo()
    {
        var pptPath = CreateTestPresentation("test_add_video.pptx");
        var videoPath = CreateFakeVideoFile("test_video.mp4");
        var outputPath = CreateTestFilePath("test_add_video_output.pptx");
        var result = _tool.Execute("add_video", pptPath, slideIndex: 0, videoPath: videoPath, x: 100, y: 100,
            width: 320, height: 240, outputPath: outputPath);
        Assert.StartsWith("Video embedded into slide", result);
        using var presentation = new Presentation(outputPath);
        Assert.NotEmpty(presentation.Slides[0].Shapes.OfType<IVideoFrame>());
    }

    [Fact]
    public void AddVideo_WithDefaultSize_ShouldUseDefaults()
    {
        var pptPath = CreateTestPresentation("test_add_video_default.pptx");
        var videoPath = CreateFakeVideoFile("test_video_default.mp4");
        var outputPath = CreateTestFilePath("test_add_video_default_output.pptx");
        var result = _tool.Execute("add_video", pptPath, slideIndex: 0, videoPath: videoPath, outputPath: outputPath);
        Assert.StartsWith("Video embedded into slide", result);
        using var presentation = new Presentation(outputPath);
        var videoFrame = presentation.Slides[0].Shapes.OfType<IVideoFrame>().First();
        Assert.Equal(320, videoFrame.Width, 1);
        Assert.Equal(240, videoFrame.Height, 1);
    }

    [Fact]
    public void DeleteVideo_ShouldRemoveVideoFrame()
    {
        var pptPath = CreatePresentationWithVideo("test_delete_video.pptx");
        var outputPath = CreateTestFilePath("test_delete_video_output.pptx");
        var result = _tool.Execute("delete_video", pptPath, slideIndex: 0, shapeIndex: 0, outputPath: outputPath);
        Assert.StartsWith("Video deleted from slide", result);
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
        Assert.StartsWith("Playback settings updated", result);
        Assert.Contains("playMode=onclick", result);
        Assert.Contains("volume=loud", result);
        Assert.Contains("loop=true", result);
    }

    [Fact]
    public void SetPlayback_ForVideo_ShouldUpdateSettings()
    {
        var pptPath = CreatePresentationWithVideo("test_playback_video.pptx");
        var outputPath = CreateTestFilePath("test_playback_video_output.pptx");
        var result = _tool.Execute("set_playback", pptPath, slideIndex: 0, shapeIndex: 0, playMode: "auto", loop: true,
            rewind: true, volume: "mute", outputPath: outputPath);
        Assert.StartsWith("Playback settings updated", result);
        Assert.Contains("rewind=true", result);
        Assert.Contains("loop=true", result);
    }

    [Theory]
    [InlineData("ADD_AUDIO")]
    [InlineData("Add_Audio")]
    [InlineData("add_audio")]
    public void Operation_ShouldBeCaseInsensitive_AddAudio(string operation)
    {
        var pptPath = CreateTestPresentation($"test_case_add_audio_{operation.Replace("_", "")}.pptx");
        var audioPath = CreateFakeAudioFile($"test_case_audio_{operation.Replace("_", "")}.mp3");
        var outputPath = CreateTestFilePath($"test_case_add_audio_{operation.Replace("_", "")}_output.pptx");
        var result = _tool.Execute(operation, pptPath, slideIndex: 0, audioPath: audioPath, outputPath: outputPath);
        Assert.StartsWith("Audio embedded into slide", result);
    }

    [Theory]
    [InlineData("DELETE_AUDIO")]
    [InlineData("Delete_Audio")]
    [InlineData("delete_audio")]
    public void Operation_ShouldBeCaseInsensitive_DeleteAudio(string operation)
    {
        var pptPath = CreatePresentationWithAudio($"test_case_delete_audio_{operation.Replace("_", "")}.pptx");
        var outputPath = CreateTestFilePath($"test_case_delete_audio_{operation.Replace("_", "")}_output.pptx");
        var result = _tool.Execute(operation, pptPath, slideIndex: 0, shapeIndex: 0, outputPath: outputPath);
        Assert.StartsWith("Audio deleted from slide", result);
    }

    [Theory]
    [InlineData("ADD_VIDEO")]
    [InlineData("Add_Video")]
    [InlineData("add_video")]
    public void Operation_ShouldBeCaseInsensitive_AddVideo(string operation)
    {
        var pptPath = CreateTestPresentation($"test_case_add_video_{operation.Replace("_", "")}.pptx");
        var videoPath = CreateFakeVideoFile($"test_case_video_{operation.Replace("_", "")}.mp4");
        var outputPath = CreateTestFilePath($"test_case_add_video_{operation.Replace("_", "")}_output.pptx");
        var result = _tool.Execute(operation, pptPath, slideIndex: 0, videoPath: videoPath, outputPath: outputPath);
        Assert.StartsWith("Video embedded into slide", result);
    }

    [Theory]
    [InlineData("DELETE_VIDEO")]
    [InlineData("Delete_Video")]
    [InlineData("delete_video")]
    public void Operation_ShouldBeCaseInsensitive_DeleteVideo(string operation)
    {
        var pptPath = CreatePresentationWithVideo($"test_case_delete_video_{operation.Replace("_", "")}.pptx");
        var outputPath = CreateTestFilePath($"test_case_delete_video_{operation.Replace("_", "")}_output.pptx");
        var result = _tool.Execute(operation, pptPath, slideIndex: 0, shapeIndex: 0, outputPath: outputPath);
        Assert.StartsWith("Video deleted from slide", result);
    }

    [Theory]
    [InlineData("SET_PLAYBACK")]
    [InlineData("Set_Playback")]
    [InlineData("set_playback")]
    public void Operation_ShouldBeCaseInsensitive_SetPlayback(string operation)
    {
        var pptPath = CreatePresentationWithAudio($"test_case_playback_{operation.Replace("_", "")}.pptx");
        var outputPath = CreateTestFilePath($"test_case_playback_{operation.Replace("_", "")}_output.pptx");
        var result = _tool.Execute(operation, pptPath, slideIndex: 0, shapeIndex: 0, outputPath: outputPath);
        Assert.StartsWith("Playback settings updated", result);
    }

    #endregion

    #region Exception

    [Fact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        var pptPath = CreateTestPresentation("test_unknown_op.pptx");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("unknown", pptPath, slideIndex: 0));
        Assert.Contains("Unknown operation", ex.Message);
    }

    [Fact]
    public void AddAudio_WithoutAudioPath_ShouldThrowArgumentException()
    {
        var pptPath = CreateTestPresentation("test_add_audio_no_path.pptx");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("add_audio", pptPath, slideIndex: 0));
        Assert.Contains("audioPath is required", ex.Message);
    }

    [Fact]
    public void AddAudio_WithNonExistentFile_ShouldThrowFileNotFoundException()
    {
        var pptPath = CreateTestPresentation("test_add_audio_notfound.pptx");
        var ex = Assert.Throws<FileNotFoundException>(() =>
            _tool.Execute("add_audio", pptPath, slideIndex: 0, audioPath: "nonexistent.mp3"));
        Assert.Contains("Audio file not found", ex.Message);
    }

    [Fact]
    public void AddAudio_WithInvalidSlideIndex_ShouldThrowArgumentException()
    {
        var pptPath = CreateTestPresentation("test_add_audio_invalid_slide.pptx");
        var audioPath = CreateFakeAudioFile("test_invalid_slide.mp3");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("add_audio", pptPath, slideIndex: 99, audioPath: audioPath));
        Assert.Contains("slide", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void DeleteAudio_WithoutShapeIndex_ShouldThrowArgumentException()
    {
        var pptPath = CreatePresentationWithAudio("test_delete_audio_no_index.pptx");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("delete_audio", pptPath, slideIndex: 0));
        Assert.Contains("shapeIndex is required", ex.Message);
    }

    [Fact]
    public void DeleteAudio_WithNonAudioShape_ShouldThrowArgumentException()
    {
        var pptPath = CreatePresentationWithShape("test_delete_audio_wrong_type.pptx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("delete_audio", pptPath, slideIndex: 0, shapeIndex: 0));
        Assert.Contains("not an audio frame", ex.Message);
    }

    [Fact]
    public void AddVideo_WithoutVideoPath_ShouldThrowArgumentException()
    {
        var pptPath = CreateTestPresentation("test_add_video_no_path.pptx");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("add_video", pptPath, slideIndex: 0));
        Assert.Contains("videoPath is required", ex.Message);
    }

    [Fact]
    public void AddVideo_WithNonExistentFile_ShouldThrowFileNotFoundException()
    {
        var pptPath = CreateTestPresentation("test_add_video_notfound.pptx");
        var ex = Assert.Throws<FileNotFoundException>(() =>
            _tool.Execute("add_video", pptPath, slideIndex: 0, videoPath: "nonexistent.mp4"));
        Assert.Contains("Video file not found", ex.Message);
    }

    [Fact]
    public void AddVideo_WithInvalidSlideIndex_ShouldThrowArgumentException()
    {
        var pptPath = CreateTestPresentation("test_add_video_invalid_slide.pptx");
        var videoPath = CreateFakeVideoFile("test_invalid_slide.mp4");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("add_video", pptPath, slideIndex: 99, videoPath: videoPath));
        Assert.Contains("slide", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void DeleteVideo_WithoutShapeIndex_ShouldThrowArgumentException()
    {
        var pptPath = CreatePresentationWithVideo("test_delete_video_no_index.pptx");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("delete_video", pptPath, slideIndex: 0));
        Assert.Contains("shapeIndex is required", ex.Message);
    }

    [Fact]
    public void DeleteVideo_WithNonVideoShape_ShouldThrowArgumentException()
    {
        var pptPath = CreatePresentationWithShape("test_delete_video_wrong_type.pptx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("delete_video", pptPath, slideIndex: 0, shapeIndex: 0));
        Assert.Contains("not a video frame", ex.Message);
    }

    [Fact]
    public void SetPlayback_WithoutShapeIndex_ShouldThrowArgumentException()
    {
        var pptPath = CreatePresentationWithAudio("test_playback_no_index.pptx");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("set_playback", pptPath, slideIndex: 0));
        Assert.Contains("shapeIndex is required", ex.Message);
    }

    [Fact]
    public void SetPlayback_WithInvalidVolume_ShouldThrowArgumentException()
    {
        var pptPath = CreatePresentationWithAudio("test_playback_invalid_volume.pptx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("set_playback", pptPath, slideIndex: 0, shapeIndex: 0, volume: "invalid"));
        Assert.Contains("Unknown volume", ex.Message);
        Assert.Contains("Supported values", ex.Message);
    }

    [Fact]
    public void SetPlayback_WithNonMediaShape_ShouldThrowArgumentException()
    {
        var pptPath = CreatePresentationWithShape("test_playback_wrong_type.pptx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("set_playback", pptPath, slideIndex: 0, shapeIndex: 0));
        Assert.Contains("not an audio or video frame", ex.Message);
    }

    #endregion

    #region Session

    [Fact]
    public void AddAudio_WithSessionId_ShouldAddInMemory()
    {
        var pptPath = CreateTestPresentation("test_session_add_audio.pptx");
        var audioPath = CreateFakeAudioFile("session_audio.mp3");
        var sessionId = OpenSession(pptPath);
        var ppt = SessionManager.GetDocument<Presentation>(sessionId);
        var initialCount = ppt.Slides[0].Shapes.OfType<IAudioFrame>().Count();
        var result = _tool.Execute("add_audio", sessionId: sessionId, slideIndex: 0, audioPath: audioPath, x: 100,
            y: 100);
        Assert.StartsWith("Audio embedded into slide", result);
        Assert.Contains("session", result);
        Assert.True(ppt.Slides[0].Shapes.OfType<IAudioFrame>().Count() > initialCount);
    }

    [Fact]
    public void AddVideo_WithSessionId_ShouldAddInMemory()
    {
        var pptPath = CreateTestPresentation("test_session_add_video.pptx");
        var videoPath = CreateFakeVideoFile("session_video.mp4");
        var sessionId = OpenSession(pptPath);
        var ppt = SessionManager.GetDocument<Presentation>(sessionId);
        var initialCount = ppt.Slides[0].Shapes.OfType<IVideoFrame>().Count();
        var result = _tool.Execute("add_video", sessionId: sessionId, slideIndex: 0, videoPath: videoPath, x: 100,
            y: 100, width: 320, height: 240);
        Assert.StartsWith("Video embedded into slide", result);
        Assert.Contains("session", result);
        Assert.True(ppt.Slides[0].Shapes.OfType<IVideoFrame>().Count() > initialCount);
    }

    [Fact]
    public void DeleteAudio_WithSessionId_ShouldDeleteInMemory()
    {
        var pptPath = CreatePresentationWithAudio("test_session_delete_audio.pptx");
        var sessionId = OpenSession(pptPath);
        var ppt = SessionManager.GetDocument<Presentation>(sessionId);
        Assert.NotEmpty(ppt.Slides[0].Shapes.OfType<IAudioFrame>());
        var result = _tool.Execute("delete_audio", sessionId: sessionId, slideIndex: 0, shapeIndex: 0);
        Assert.StartsWith("Audio deleted from slide", result);
        Assert.Contains("session", result);
        Assert.Empty(ppt.Slides[0].Shapes.OfType<IAudioFrame>());
    }

    [Fact]
    public void DeleteVideo_WithSessionId_ShouldDeleteInMemory()
    {
        var pptPath = CreatePresentationWithVideo("test_session_delete_video.pptx");
        var sessionId = OpenSession(pptPath);
        var ppt = SessionManager.GetDocument<Presentation>(sessionId);
        Assert.NotEmpty(ppt.Slides[0].Shapes.OfType<IVideoFrame>());
        var result = _tool.Execute("delete_video", sessionId: sessionId, slideIndex: 0, shapeIndex: 0);
        Assert.StartsWith("Video deleted from slide", result);
        Assert.Contains("session", result);
        Assert.Empty(ppt.Slides[0].Shapes.OfType<IVideoFrame>());
    }

    [Fact]
    public void SetPlayback_WithSessionId_ShouldUpdateInMemory()
    {
        var pptPath = CreatePresentationWithAudio("test_session_playback.pptx");
        var sessionId = OpenSession(pptPath);
        var result = _tool.Execute("set_playback", sessionId: sessionId, slideIndex: 0, shapeIndex: 0,
            playMode: "onclick", loop: true, volume: "loud");
        Assert.StartsWith("Playback settings updated", result);
        Assert.Contains("session", result);
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
        var pptPath1 = CreateTestPresentation("test_path_media.pptx");
        var pptPath2 = CreateTestPresentation("test_session_media.pptx");
        var audioPath = CreateFakeAudioFile("preference_audio.mp3");
        var sessionId = OpenSession(pptPath2);
        var ppt = SessionManager.GetDocument<Presentation>(sessionId);
        var result = _tool.Execute("add_audio", pptPath1, sessionId, slideIndex: 0, audioPath: audioPath, x: 100,
            y: 100);
        Assert.Contains("session", result);
        Assert.NotEmpty(ppt.Slides[0].Shapes.OfType<IAudioFrame>());
    }

    #endregion
}