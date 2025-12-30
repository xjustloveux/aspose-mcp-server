using System.Text.Json.Nodes;
using Aspose.Slides;
using Aspose.Slides.Export;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.PowerPoint;

namespace AsposeMcpServer.Tests.PowerPoint;

public class PptMediaToolTests : TestBase
{
    private readonly PptMediaTool _tool = new();

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

    #region Unknown Operation Tests

    [Fact]
    public async Task ExecuteAsync_UnknownOperation_ShouldThrow()
    {
        // Arrange
        var pptPath = CreateTestPresentation("test_unknown_op.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "unknown",
            ["path"] = pptPath,
            ["slideIndex"] = 0
        };

        // Act & Assert
        var ex = await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("Unknown operation", ex.Message);
    }

    #endregion

    #region AddAudio Tests

    [Fact]
    public async Task AddAudio_ShouldEmbedAudio()
    {
        // Arrange
        var pptPath = CreateTestPresentation("test_add_audio.pptx");
        var audioPath = CreateFakeAudioFile("test_audio.mp3");
        var outputPath = CreateTestFilePath("test_add_audio_output.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "add_audio",
            ["path"] = pptPath,
            ["outputPath"] = outputPath,
            ["slideIndex"] = 0,
            ["audioPath"] = audioPath,
            ["x"] = 100,
            ["y"] = 100
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("Audio embedded", result);
        using var presentation = new Presentation(outputPath);
        var audioFrames = presentation.Slides[0].Shapes.OfType<IAudioFrame>().ToList();
        Assert.NotEmpty(audioFrames);
    }

    [Fact]
    public async Task AddAudio_WithNonExistentFile_ShouldThrow()
    {
        // Arrange
        var pptPath = CreateTestPresentation("test_add_audio_notfound.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "add_audio",
            ["path"] = pptPath,
            ["slideIndex"] = 0,
            ["audioPath"] = "nonexistent_audio.mp3"
        };

        // Act & Assert
        var ex = await Assert.ThrowsAsync<FileNotFoundException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("Audio file not found", ex.Message);
    }

    [Fact]
    public async Task AddAudio_WithInvalidSlideIndex_ShouldThrow()
    {
        // Arrange
        var pptPath = CreateTestPresentation("test_add_audio_invalid_slide.pptx");
        var audioPath = CreateFakeAudioFile("test_audio2.mp3");
        var arguments = new JsonObject
        {
            ["operation"] = "add_audio",
            ["path"] = pptPath,
            ["slideIndex"] = 99,
            ["audioPath"] = audioPath
        };

        // Act & Assert
        var ex = await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("slide", ex.Message.ToLower());
    }

    #endregion

    #region DeleteAudio Tests

    [Fact]
    public async Task DeleteAudio_ShouldRemoveAudioFrame()
    {
        // Arrange - Create presentation with audio
        var pptPath = CreateTestFilePath("test_delete_audio.pptx");
        var audioPath = CreateFakeAudioFile("audio_to_delete.mp3");
        using (var presentation = new Presentation())
        {
            await using var audioStream = File.OpenRead(audioPath);
            presentation.Slides[0].Shapes.AddAudioFrameEmbedded(100, 100, 80, 80, audioStream);
            presentation.Save(pptPath, SaveFormat.Pptx);
        }

        var outputPath = CreateTestFilePath("test_delete_audio_output.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "delete_audio",
            ["path"] = pptPath,
            ["outputPath"] = outputPath,
            ["slideIndex"] = 0,
            ["shapeIndex"] = 0
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("Audio deleted", result);
        using var resultPres = new Presentation(outputPath);
        var audioFrames = resultPres.Slides[0].Shapes.OfType<IAudioFrame>().ToList();
        Assert.Empty(audioFrames);
    }

    [Fact]
    public async Task DeleteAudio_WithNonAudioShape_ShouldThrow()
    {
        // Arrange - Create presentation with a text shape
        var pptPath = CreateTestFilePath("test_delete_audio_wrong_type.pptx");
        using (var presentation = new Presentation())
        {
            presentation.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
            presentation.Save(pptPath, SaveFormat.Pptx);
        }

        var arguments = new JsonObject
        {
            ["operation"] = "delete_audio",
            ["path"] = pptPath,
            ["slideIndex"] = 0,
            ["shapeIndex"] = 0
        };

        // Act & Assert
        var ex = await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("not an audio frame", ex.Message);
    }

    #endregion

    #region AddVideo Tests

    [Fact]
    public async Task AddVideo_ShouldEmbedVideo()
    {
        // Arrange
        var pptPath = CreateTestPresentation("test_add_video.pptx");
        var videoPath = CreateFakeVideoFile("test_video.mp4");
        var outputPath = CreateTestFilePath("test_add_video_output.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "add_video",
            ["path"] = pptPath,
            ["outputPath"] = outputPath,
            ["slideIndex"] = 0,
            ["videoPath"] = videoPath,
            ["x"] = 100,
            ["y"] = 100,
            ["width"] = 320,
            ["height"] = 240
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("Video embedded", result);
        using var presentation = new Presentation(outputPath);
        var videoFrames = presentation.Slides[0].Shapes.OfType<IVideoFrame>().ToList();
        Assert.NotEmpty(videoFrames);
    }

    [Fact]
    public async Task AddVideo_WithNonExistentFile_ShouldThrow()
    {
        // Arrange
        var pptPath = CreateTestPresentation("test_add_video_notfound.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "add_video",
            ["path"] = pptPath,
            ["slideIndex"] = 0,
            ["videoPath"] = "nonexistent_video.mp4"
        };

        // Act & Assert
        var ex = await Assert.ThrowsAsync<FileNotFoundException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("Video file not found", ex.Message);
    }

    #endregion

    #region DeleteVideo Tests

    [Fact]
    public async Task DeleteVideo_ShouldRemoveVideoFrame()
    {
        // Arrange - Create presentation with video
        var pptPath = CreateTestFilePath("test_delete_video.pptx");
        var videoPath = CreateFakeVideoFile("video_to_delete.mp4");
        using (var presentation = new Presentation())
        {
            await using var videoStream = File.OpenRead(videoPath);
            var videoObj = presentation.Videos.AddVideo(videoStream, LoadingStreamBehavior.KeepLocked);
            presentation.Slides[0].Shapes.AddVideoFrame(100, 100, 320, 240, videoObj);
            presentation.Save(pptPath, SaveFormat.Pptx);
        }

        var outputPath = CreateTestFilePath("test_delete_video_output.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "delete_video",
            ["path"] = pptPath,
            ["outputPath"] = outputPath,
            ["slideIndex"] = 0,
            ["shapeIndex"] = 0
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("Video deleted", result);
        using var resultPres = new Presentation(outputPath);
        var videoFrames = resultPres.Slides[0].Shapes.OfType<IVideoFrame>().ToList();
        Assert.Empty(videoFrames);
    }

    [Fact]
    public async Task DeleteVideo_WithNonVideoShape_ShouldThrow()
    {
        // Arrange - Create presentation with a rectangle shape
        var pptPath = CreateTestFilePath("test_delete_video_wrong_type.pptx");
        using (var presentation = new Presentation())
        {
            presentation.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
            presentation.Save(pptPath, SaveFormat.Pptx);
        }

        var arguments = new JsonObject
        {
            ["operation"] = "delete_video",
            ["path"] = pptPath,
            ["slideIndex"] = 0,
            ["shapeIndex"] = 0
        };

        // Act & Assert
        var ex = await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("not a video frame", ex.Message);
    }

    #endregion

    #region SetPlayback Tests

    [Fact]
    public async Task SetPlayback_ForAudio_ShouldUpdateSettings()
    {
        // Arrange - Create presentation with audio
        var pptPath = CreateTestFilePath("test_set_playback_audio.pptx");
        var audioPath = CreateFakeAudioFile("playback_audio.mp3");
        using (var presentation = new Presentation())
        {
            await using var audioStream = File.OpenRead(audioPath);
            presentation.Slides[0].Shapes.AddAudioFrameEmbedded(100, 100, 80, 80, audioStream);
            presentation.Save(pptPath, SaveFormat.Pptx);
        }

        var outputPath = CreateTestFilePath("test_set_playback_audio_output.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "set_playback",
            ["path"] = pptPath,
            ["outputPath"] = outputPath,
            ["slideIndex"] = 0,
            ["shapeIndex"] = 0,
            ["playMode"] = "onclick",
            ["loop"] = true,
            ["volume"] = "loud"
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("Playback settings updated", result);
        Assert.Contains("playMode=onclick", result);
        Assert.Contains("volume=loud", result);
        Assert.Contains("loop=true", result);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public async Task SetPlayback_ForVideo_ShouldUpdateSettings()
    {
        // Arrange - Create presentation with video
        var pptPath = CreateTestFilePath("test_set_playback_video.pptx");
        var videoPath = CreateFakeVideoFile("playback_video.mp4");
        using (var presentation = new Presentation())
        {
            await using var videoStream = File.OpenRead(videoPath);
            var videoObj = presentation.Videos.AddVideo(videoStream, LoadingStreamBehavior.KeepLocked);
            presentation.Slides[0].Shapes.AddVideoFrame(100, 100, 320, 240, videoObj);
            presentation.Save(pptPath, SaveFormat.Pptx);
        }

        var outputPath = CreateTestFilePath("test_set_playback_video_output.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "set_playback",
            ["path"] = pptPath,
            ["outputPath"] = outputPath,
            ["slideIndex"] = 0,
            ["shapeIndex"] = 0,
            ["playMode"] = "auto",
            ["loop"] = true,
            ["rewind"] = true,
            ["volume"] = "mute"
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("Playback settings updated", result);
        Assert.Contains("rewind=true", result);
        Assert.Contains("loop=true", result);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public async Task SetPlayback_WithInvalidVolume_ShouldThrow()
    {
        // Arrange - Create presentation with audio
        var pptPath = CreateTestFilePath("test_set_playback_invalid_volume.pptx");
        var audioPath = CreateFakeAudioFile("invalid_volume_audio.mp3");
        using (var presentation = new Presentation())
        {
            await using var audioStream = File.OpenRead(audioPath);
            presentation.Slides[0].Shapes.AddAudioFrameEmbedded(100, 100, 80, 80, audioStream);
            presentation.Save(pptPath, SaveFormat.Pptx);
        }

        var arguments = new JsonObject
        {
            ["operation"] = "set_playback",
            ["path"] = pptPath,
            ["slideIndex"] = 0,
            ["shapeIndex"] = 0,
            ["volume"] = "invalid_volume"
        };

        // Act & Assert
        var ex = await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("Unknown volume", ex.Message);
        Assert.Contains("Supported values", ex.Message);
    }

    [Fact]
    public async Task SetPlayback_WithNonMediaShape_ShouldThrow()
    {
        // Arrange - Create presentation with a rectangle shape
        var pptPath = CreateTestFilePath("test_set_playback_wrong_type.pptx");
        using (var presentation = new Presentation())
        {
            presentation.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
            presentation.Save(pptPath, SaveFormat.Pptx);
        }

        var arguments = new JsonObject
        {
            ["operation"] = "set_playback",
            ["path"] = pptPath,
            ["slideIndex"] = 0,
            ["shapeIndex"] = 0
        };

        // Act & Assert
        var ex = await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("not an audio or video frame", ex.Message);
    }

    #endregion
}