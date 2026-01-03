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
        File.WriteAllBytes(audioPath, "ID3"u8.ToArray()); // ID3 header
        return audioPath;
    }

    private string CreateFakeVideoFile(string fileName)
    {
        var videoPath = CreateTestFilePath(fileName);
        // Minimal MP4 header bytes
        File.WriteAllBytes(videoPath, [0x00, 0x00, 0x00, 0x1C, 0x66, 0x74, 0x79, 0x70]);
        return videoPath;
    }

    #region General Tests

    [Fact]
    public void AddAudio_ShouldEmbedAudio()
    {
        var pptPath = CreateTestPresentation("test_add_audio.pptx");
        var audioPath = CreateFakeAudioFile("test_audio.mp3");
        var outputPath = CreateTestFilePath("test_add_audio_output.pptx");
        var result = _tool.Execute("add_audio", pptPath, slideIndex: 0, audioPath: audioPath, x: 100, y: 100,
            outputPath: outputPath);
        Assert.Contains("Audio embedded", result);
        using var presentation = new Presentation(outputPath);
        var audioFrames = presentation.Slides[0].Shapes.OfType<IAudioFrame>().ToList();
        Assert.NotEmpty(audioFrames);
    }

    [Fact]
    public void DeleteAudio_ShouldRemoveAudioFrame()
    {
        // Arrange - Create presentation with audio
        var pptPath = CreateTestFilePath("test_delete_audio.pptx");
        var audioPath = CreateFakeAudioFile("audio_to_delete.mp3");
        using (var presentation = new Presentation())
        {
            using var audioStream = File.OpenRead(audioPath);
            presentation.Slides[0].Shapes.AddAudioFrameEmbedded(100, 100, 80, 80, audioStream);
            presentation.Save(pptPath, SaveFormat.Pptx);
        }

        var outputPath = CreateTestFilePath("test_delete_audio_output.pptx");
        var result = _tool.Execute("delete_audio", pptPath, slideIndex: 0, shapeIndex: 0, outputPath: outputPath);
        Assert.Contains("Audio deleted", result);
        using var resultPres = new Presentation(outputPath);
        var audioFrames = resultPres.Slides[0].Shapes.OfType<IAudioFrame>().ToList();
        Assert.Empty(audioFrames);
    }

    [Fact]
    public void AddVideo_ShouldEmbedVideo()
    {
        var pptPath = CreateTestPresentation("test_add_video.pptx");
        var videoPath = CreateFakeVideoFile("test_video.mp4");
        var outputPath = CreateTestFilePath("test_add_video_output.pptx");
        var result = _tool.Execute("add_video", pptPath, slideIndex: 0, videoPath: videoPath, x: 100, y: 100,
            width: 320, height: 240, outputPath: outputPath);
        Assert.Contains("Video embedded", result);
        using var presentation = new Presentation(outputPath);
        var videoFrames = presentation.Slides[0].Shapes.OfType<IVideoFrame>().ToList();
        Assert.NotEmpty(videoFrames);
    }

    [Fact]
    public void DeleteVideo_ShouldRemoveVideoFrame()
    {
        // Arrange - Create presentation with video
        var pptPath = CreateTestFilePath("test_delete_video.pptx");
        var videoPath = CreateFakeVideoFile("video_to_delete.mp4");
        using (var presentation = new Presentation())
        {
            using var videoStream = File.OpenRead(videoPath);
            var videoObj = presentation.Videos.AddVideo(videoStream, LoadingStreamBehavior.KeepLocked);
            presentation.Slides[0].Shapes.AddVideoFrame(100, 100, 320, 240, videoObj);
            presentation.Save(pptPath, SaveFormat.Pptx);
        }

        var outputPath = CreateTestFilePath("test_delete_video_output.pptx");
        var result = _tool.Execute("delete_video", pptPath, slideIndex: 0, shapeIndex: 0, outputPath: outputPath);
        Assert.Contains("Video deleted", result);
        using var resultPres = new Presentation(outputPath);
        var videoFrames = resultPres.Slides[0].Shapes.OfType<IVideoFrame>().ToList();
        Assert.Empty(videoFrames);
    }

    [Fact]
    public void SetPlayback_ForAudio_ShouldUpdateSettings()
    {
        // Arrange - Create presentation with audio
        var pptPath = CreateTestFilePath("test_set_playback_audio.pptx");
        var audioPath = CreateFakeAudioFile("playback_audio.mp3");
        using (var presentation = new Presentation())
        {
            using var audioStream = File.OpenRead(audioPath);
            presentation.Slides[0].Shapes.AddAudioFrameEmbedded(100, 100, 80, 80, audioStream);
            presentation.Save(pptPath, SaveFormat.Pptx);
        }

        var outputPath = CreateTestFilePath("test_set_playback_audio_output.pptx");
        var result = _tool.Execute("set_playback", pptPath, slideIndex: 0, shapeIndex: 0, playMode: "onclick",
            loop: true, volume: "loud", outputPath: outputPath);
        Assert.Contains("Playback settings updated", result);
        Assert.Contains("playMode=onclick", result);
        Assert.Contains("volume=loud", result);
        Assert.Contains("loop=true", result);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void SetPlayback_ForVideo_ShouldUpdateSettings()
    {
        // Arrange - Create presentation with video
        var pptPath = CreateTestFilePath("test_set_playback_video.pptx");
        var videoPath = CreateFakeVideoFile("playback_video.mp4");
        using (var presentation = new Presentation())
        {
            using var videoStream = File.OpenRead(videoPath);
            var videoObj = presentation.Videos.AddVideo(videoStream, LoadingStreamBehavior.KeepLocked);
            presentation.Slides[0].Shapes.AddVideoFrame(100, 100, 320, 240, videoObj);
            presentation.Save(pptPath, SaveFormat.Pptx);
        }

        var outputPath = CreateTestFilePath("test_set_playback_video_output.pptx");
        var result = _tool.Execute("set_playback", pptPath, slideIndex: 0, shapeIndex: 0, playMode: "auto", loop: true,
            rewind: true, volume: "mute", outputPath: outputPath);
        Assert.Contains("Playback settings updated", result);
        Assert.Contains("rewind=true", result);
        Assert.Contains("loop=true", result);
        Assert.True(File.Exists(outputPath));
    }

    #endregion

    #region Exception Tests

    [Fact]
    public void ExecuteAsync_UnknownOperation_ShouldThrow()
    {
        var pptPath = CreateTestPresentation("test_unknown_op.pptx");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("unknown", pptPath, slideIndex: 0));
        Assert.Contains("Unknown operation", ex.Message);
    }

    [Fact]
    public void AddAudio_WithNonExistentFile_ShouldThrow()
    {
        var pptPath = CreateTestPresentation("test_add_audio_notfound.pptx");
        var ex = Assert.Throws<FileNotFoundException>(() =>
            _tool.Execute("add_audio", pptPath, slideIndex: 0, audioPath: "nonexistent_audio.mp3"));
        Assert.Contains("Audio file not found", ex.Message);
    }

    [Fact]
    public void AddAudio_WithInvalidSlideIndex_ShouldThrow()
    {
        var pptPath = CreateTestPresentation("test_add_audio_invalid_slide.pptx");
        var audioPath = CreateFakeAudioFile("test_audio2.mp3");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("add_audio", pptPath, slideIndex: 99, audioPath: audioPath));
        Assert.Contains("slide", ex.Message.ToLower());
    }

    [Fact]
    public void DeleteAudio_WithNonAudioShape_ShouldThrow()
    {
        // Arrange - Create presentation with a text shape
        var pptPath = CreateTestFilePath("test_delete_audio_wrong_type.pptx");
        using (var presentation = new Presentation())
        {
            presentation.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
            presentation.Save(pptPath, SaveFormat.Pptx);
        }

        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("delete_audio", pptPath, slideIndex: 0, shapeIndex: 0));
        Assert.Contains("not an audio frame", ex.Message);
    }

    [Fact]
    public void AddVideo_WithNonExistentFile_ShouldThrow()
    {
        var pptPath = CreateTestPresentation("test_add_video_notfound.pptx");
        var ex = Assert.Throws<FileNotFoundException>(() =>
            _tool.Execute("add_video", pptPath, slideIndex: 0, videoPath: "nonexistent_video.mp4"));
        Assert.Contains("Video file not found", ex.Message);
    }

    [Fact]
    public void DeleteVideo_WithNonVideoShape_ShouldThrow()
    {
        // Arrange - Create presentation with a rectangle shape
        var pptPath = CreateTestFilePath("test_delete_video_wrong_type.pptx");
        using (var presentation = new Presentation())
        {
            presentation.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
            presentation.Save(pptPath, SaveFormat.Pptx);
        }

        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("delete_video", pptPath, slideIndex: 0, shapeIndex: 0));
        Assert.Contains("not a video frame", ex.Message);
    }

    [Fact]
    public void SetPlayback_WithInvalidVolume_ShouldThrow()
    {
        // Arrange - Create presentation with audio
        var pptPath = CreateTestFilePath("test_set_playback_invalid_volume.pptx");
        var audioPath = CreateFakeAudioFile("invalid_volume_audio.mp3");
        using (var presentation = new Presentation())
        {
            using var audioStream = File.OpenRead(audioPath);
            presentation.Slides[0].Shapes.AddAudioFrameEmbedded(100, 100, 80, 80, audioStream);
            presentation.Save(pptPath, SaveFormat.Pptx);
        }

        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("set_playback", pptPath, slideIndex: 0, shapeIndex: 0, volume: "invalid_volume"));
        Assert.Contains("Unknown volume", ex.Message);
        Assert.Contains("Supported values", ex.Message);
    }

    [Fact]
    public void SetPlayback_WithNonMediaShape_ShouldThrow()
    {
        // Arrange - Create presentation with a rectangle shape
        var pptPath = CreateTestFilePath("test_set_playback_wrong_type.pptx");
        using (var presentation = new Presentation())
        {
            presentation.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
            presentation.Save(pptPath, SaveFormat.Pptx);
        }

        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("set_playback", pptPath, slideIndex: 0, shapeIndex: 0));
        Assert.Contains("not an audio or video frame", ex.Message);
    }

    #endregion

    #region Session ID Tests

    [Fact]
    public void AddAudio_WithSessionId_ShouldAddInMemory()
    {
        var pptPath = CreateTestPresentation("test_session_add_audio.pptx");
        var audioPath = CreateFakeAudioFile("session_audio.mp3");
        var sessionId = OpenSession(pptPath);
        var ppt = SessionManager.GetDocument<Presentation>(sessionId);
        var initialAudioCount = ppt.Slides[0].Shapes.OfType<IAudioFrame>().Count();
        var result = _tool.Execute("add_audio", sessionId: sessionId, slideIndex: 0, audioPath: audioPath, x: 100,
            y: 100);
        Assert.Contains("Audio embedded", result);
        Assert.Contains("session", result);

        // Verify in-memory changes
        var currentAudioCount = ppt.Slides[0].Shapes.OfType<IAudioFrame>().Count();
        Assert.True(currentAudioCount > initialAudioCount);
    }

    [Fact]
    public void AddVideo_WithSessionId_ShouldAddInMemory()
    {
        var pptPath = CreateTestPresentation("test_session_add_video.pptx");
        var videoPath = CreateFakeVideoFile("session_video.mp4");
        var sessionId = OpenSession(pptPath);
        var ppt = SessionManager.GetDocument<Presentation>(sessionId);
        var initialVideoCount = ppt.Slides[0].Shapes.OfType<IVideoFrame>().Count();
        var result = _tool.Execute("add_video", sessionId: sessionId, slideIndex: 0, videoPath: videoPath, x: 100,
            y: 100, width: 320, height: 240);
        Assert.Contains("Video embedded", result);
        Assert.Contains("session", result);

        // Verify in-memory changes
        var currentVideoCount = ppt.Slides[0].Shapes.OfType<IVideoFrame>().Count();
        Assert.True(currentVideoCount > initialVideoCount);
    }

    [Fact]
    public void SetPlayback_WithSessionId_ShouldUpdateInMemory()
    {
        // Arrange - Create presentation with audio
        var pptPath = CreateTestFilePath("test_session_set_playback.pptx");
        var audioPath = CreateFakeAudioFile("session_playback_audio.mp3");
        using (var presentation = new Presentation())
        {
            using var audioStream = File.OpenRead(audioPath);
            presentation.Slides[0].Shapes.AddAudioFrameEmbedded(100, 100, 80, 80, audioStream);
            presentation.Save(pptPath, SaveFormat.Pptx);
        }

        var sessionId = OpenSession(pptPath);
        var result = _tool.Execute("set_playback", sessionId: sessionId, slideIndex: 0, shapeIndex: 0,
            playMode: "onclick", loop: true, volume: "loud");
        Assert.Contains("Playback settings updated", result);
        Assert.Contains("session", result);
    }

    #endregion
}