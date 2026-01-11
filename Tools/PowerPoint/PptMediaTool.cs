using System.ComponentModel;
using Aspose.Slides;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Session;
using AsposeMcpServer.Handlers.PowerPoint.Media;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.PowerPoint;

/// <summary>
///     Unified tool for managing PowerPoint media.
///     Supports: add_audio, delete_audio, add_video, delete_video, set_playback
/// </summary>
[McpServerToolType]
public class PptMediaTool
{
    /// <summary>
    ///     Handler registry for media operations.
    /// </summary>
    private readonly HandlerRegistry<Presentation> _handlerRegistry;

    /// <summary>
    ///     Identity accessor for session isolation.
    /// </summary>
    private readonly ISessionIdentityAccessor? _identityAccessor;

    /// <summary>
    ///     Session manager for document lifecycle management.
    /// </summary>
    private readonly DocumentSessionManager? _sessionManager;

    /// <summary>
    ///     Initializes a new instance of the <see cref="PptMediaTool" /> class.
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory document editing.</param>
    /// <param name="identityAccessor">Optional identity accessor for session isolation.</param>
    public PptMediaTool(DocumentSessionManager? sessionManager = null,
        ISessionIdentityAccessor? identityAccessor = null)
    {
        _sessionManager = sessionManager;
        _identityAccessor = identityAccessor;
        _handlerRegistry = PptMediaHandlerRegistry.Create();
    }

    /// <summary>
    ///     Executes a PowerPoint media operation (add_audio, delete_audio, add_video, delete_video, set_playback).
    /// </summary>
    /// <param name="operation">The operation to perform: add_audio, delete_audio, add_video, delete_video, set_playback.</param>
    /// <param name="path">Presentation file path (required if no sessionId).</param>
    /// <param name="sessionId">Session ID for in-memory editing.</param>
    /// <param name="outputPath">Output file path (file mode only).</param>
    /// <param name="slideIndex">Slide index (0-based, required for all operations).</param>
    /// <param name="shapeIndex">Shape index (0-based, required for delete/set_playback).</param>
    /// <param name="audioPath">Audio file path to embed (required for add_audio).</param>
    /// <param name="videoPath">Video file path to embed (required for add_video).</param>
    /// <param name="x">X position in points from top-left corner (optional, default: 50).</param>
    /// <param name="y">Y position in points from top-left corner (optional, default: 50).</param>
    /// <param name="width">Width in points (optional, default: 80 for audio, 320 for video).</param>
    /// <param name="height">Height in points (optional, default: 80 for audio, 240 for video).</param>
    /// <param name="playMode">Playback mode: auto|onclick (optional, default: auto).</param>
    /// <param name="loop">Loop playback (optional, default: false).</param>
    /// <param name="rewind">Rewind video after play (optional, default: false).</param>
    /// <param name="volume">Volume level: mute|low|medium|loud (optional, default: medium).</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or the operation is unknown.</exception>
    [McpServerTool(Name = "ppt_media")]
    [Description(
        @"Manage PowerPoint media. Supports 5 operations: add_audio, delete_audio, add_video, delete_video, set_playback.

Coordinate system: Origin is top-left corner of slide. Units are in Points (1 inch = 72 points).
Standard slide size: 720 x 540 points (10 x 7.5 inches).

Usage examples:
- Add audio: ppt_media(operation='add_audio', path='presentation.pptx', slideIndex=0, audioPath='audio.mp3', x=100, y=100)
- Delete audio: ppt_media(operation='delete_audio', path='presentation.pptx', slideIndex=0, shapeIndex=0)
- Add video: ppt_media(operation='add_video', path='presentation.pptx', slideIndex=0, videoPath='video.mp4', x=100, y=100)
- Delete video: ppt_media(operation='delete_video', path='presentation.pptx', slideIndex=0, shapeIndex=0)
- Set playback: ppt_media(operation='set_playback', path='presentation.pptx', slideIndex=0, shapeIndex=0, playMode='auto', loop=true)")]
    public string Execute(
        [Description("Operation: add_audio, delete_audio, add_video, delete_video, set_playback")]
        string operation,
        [Description("Presentation file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Output file path (file mode only)")]
        string? outputPath = null,
        [Description("Slide index (0-based, required for all operations)")]
        int slideIndex = 0,
        [Description("Shape index (0-based, required for delete/set_playback)")]
        int? shapeIndex = null,
        [Description("Audio file path to embed (required for add_audio)")]
        string? audioPath = null,
        [Description("Video file path to embed (required for add_video)")]
        string? videoPath = null,
        [Description("X position in points from top-left corner (optional, default: 50)")]
        float x = 50,
        [Description("Y position in points from top-left corner (optional, default: 50)")]
        float y = 50,
        [Description("Width in points (optional, default: 80 for audio, 320 for video)")]
        float? width = null,
        [Description("Height in points (optional, default: 80 for audio, 240 for video)")]
        float? height = null,
        [Description("Playback mode: auto|onclick (optional, default: auto)")]
        string playMode = "auto",
        [Description("Loop playback (optional, default: false)")]
        bool loop = false,
        [Description("Rewind video after play (optional, default: false)")]
        bool rewind = false,
        [Description("Volume level: mute|low|medium|loud (optional, default: medium)")]
        string volume = "medium")
    {
        using var ctx = DocumentContext<Presentation>.Create(_sessionManager, sessionId, path, _identityAccessor);

        var parameters = BuildParameters(operation, slideIndex, shapeIndex, audioPath, videoPath,
            x, y, width, height, playMode, loop, rewind, volume);

        var handler = _handlerRegistry.GetHandler(operation);

        var operationContext = new OperationContext<Presentation>
        {
            Document = ctx.Document,
            SessionManager = _sessionManager,
            IdentityAccessor = _identityAccessor,
            SessionId = sessionId,
            SourcePath = path,
            OutputPath = outputPath
        };

        var result = handler.Execute(operationContext, parameters);

        if (operationContext.IsModified)
            ctx.Save(outputPath);

        return $"{result}\n{ctx.GetOutputMessage(outputPath)}";
    }

    /// <summary>
    ///     Builds OperationParameters from method parameters.
    /// </summary>
    private static OperationParameters BuildParameters(
        string operation,
        int slideIndex,
        int? shapeIndex,
        string? audioPath,
        string? videoPath,
        float x,
        float y,
        float? width,
        float? height,
        string playMode,
        bool loop,
        bool rewind,
        string volume)
    {
        var parameters = new OperationParameters();
        parameters.Set("slideIndex", slideIndex);

        switch (operation.ToLowerInvariant())
        {
            case "add_audio":
                if (audioPath != null) parameters.Set("audioPath", audioPath);
                parameters.Set("x", x);
                parameters.Set("y", y);
                parameters.Set("width", width ?? 80f);
                parameters.Set("height", height ?? 80f);
                break;

            case "add_video":
                if (videoPath != null) parameters.Set("videoPath", videoPath);
                parameters.Set("x", x);
                parameters.Set("y", y);
                parameters.Set("width", width ?? 320f);
                parameters.Set("height", height ?? 240f);
                break;

            case "delete_audio":
            case "delete_video":
                if (shapeIndex.HasValue) parameters.Set("shapeIndex", shapeIndex.Value);
                break;

            case "set_playback":
                if (shapeIndex.HasValue) parameters.Set("shapeIndex", shapeIndex.Value);
                parameters.Set("playMode", playMode);
                parameters.Set("loop", loop);
                parameters.Set("rewind", rewind);
                parameters.Set("volume", volume);
                break;
        }

        return parameters;
    }
}
