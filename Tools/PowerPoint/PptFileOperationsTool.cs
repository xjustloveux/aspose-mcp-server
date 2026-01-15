using System.ComponentModel;
using Aspose.Slides;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Session;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.PowerPoint;

/// <summary>
///     Unified tool for PowerPoint file operations (create, convert, merge, split).
///     Merges: PptCreateTool, PptConvertTool, PptMergeTool, PptSplitTool.
/// </summary>
[McpServerToolType]
public class PptFileOperationsTool
{
    /// <summary>
    ///     Handler registry for file operations.
    /// </summary>
    private readonly HandlerRegistry<Presentation> _handlerRegistry;

    /// <summary>
    ///     The session identity accessor for session isolation.
    /// </summary>
    private readonly ISessionIdentityAccessor? _identityAccessor;

    /// <summary>
    ///     The document session manager for managing in-memory document sessions.
    /// </summary>
    private readonly DocumentSessionManager? _sessionManager;

    /// <summary>
    ///     Initializes a new instance of the <see cref="PptFileOperationsTool" /> class.
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory document operations.</param>
    /// <param name="identityAccessor">Optional identity accessor for session isolation.</param>
    public PptFileOperationsTool(DocumentSessionManager? sessionManager = null,
        ISessionIdentityAccessor? identityAccessor = null)
    {
        _sessionManager = sessionManager;
        _identityAccessor = identityAccessor;
        _handlerRegistry =
            HandlerRegistry<Presentation>.CreateFromNamespace("AsposeMcpServer.Handlers.PowerPoint.FileOperations");
    }

    /// <summary>
    ///     Executes a PowerPoint file operation (create, convert, merge, or split).
    /// </summary>
    /// <param name="operation">The operation to perform: create, convert, merge, or split.</param>
    /// <param name="sessionId">Session ID to read presentation from session (for convert, split).</param>
    /// <param name="path">File path (output path for create operation).</param>
    /// <param name="outputPath">Output file path (required for convert).</param>
    /// <param name="inputPath">Input file path (required for convert/split).</param>
    /// <param name="outputDirectory">Output directory path (required for split).</param>
    /// <param name="format">Output format: pdf, html, pptx, jpg, png, etc. (required for convert).</param>
    /// <param name="slideIndex">Slide index to convert (0-based, for image format conversion).</param>
    /// <param name="inputPaths">Array of input presentation file paths (required for merge).</param>
    /// <param name="keepSourceFormatting">Keep source formatting when merging slides.</param>
    /// <param name="slidesPerFile">Number of slides per output file (for split).</param>
    /// <param name="startSlideIndex">Start slide index (0-based, for split).</param>
    /// <param name="endSlideIndex">End slide index (0-based, for split).</param>
    /// <param name="outputFileNamePattern">Output file name pattern with {index} placeholder.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when the operation is unknown or required parameters are missing.</exception>
    /// <exception cref="InvalidOperationException">Thrown when session management is not enabled but sessionId is provided.</exception>
    [McpServerTool(Name = "ppt_file_operations")]
    [Description(@"PowerPoint file operations. Supports 4 operations: create, convert, merge, split.

Usage examples:
- Create presentation: ppt_file_operations(operation='create', path='new.pptx')
- Convert format: ppt_file_operations(operation='convert', inputPath='presentation.pptx', outputPath='presentation.pdf', format='pdf')
- Convert from session: ppt_file_operations(operation='convert', sessionId='sess_xxx', outputPath='presentation.pdf', format='pdf')
- Convert slide to image: ppt_file_operations(operation='convert', inputPath='presentation.pptx', outputPath='slide.png', format='png', slideIndex=0)
- Merge presentations: ppt_file_operations(operation='merge', inputPath='presentation1.pptx', outputPath='merged.pptx', inputPaths=['presentation2.pptx'])
- Split presentation: ppt_file_operations(operation='split', inputPath='presentation.pptx', outputDirectory='output/')
- Split from session: ppt_file_operations(operation='split', sessionId='sess_xxx', outputDirectory='output/')")]
    public string Execute(
        [Description("Operation: create, convert, merge, split")]
        string operation,
        [Description("Session ID to read presentation from session (for convert, split)")]
        string? sessionId = null,
        [Description("File path (output path for create operation)")]
        string? path = null,
        [Description("Output file path (required for convert)")]
        string? outputPath = null,
        [Description("Input file path (required for convert/split)")]
        string? inputPath = null,
        [Description("Output directory path (required for split)")]
        string? outputDirectory = null,
        [Description("Output format (pdf, html, pptx, jpg, png, etc., required for convert)")]
        string? format = null,
        [Description("Slide index to convert (0-based, optional for convert to image format, default: 0)")]
        int slideIndex = 0,
        [Description("Array of input presentation file paths (required for merge)")]
        string[]? inputPaths = null,
        [Description("Keep source formatting (optional, for merge, default: true)")]
        bool keepSourceFormatting = true,
        [Description("Number of slides per output file (optional, for split, default: 1)")]
        int slidesPerFile = 1,
        [Description("Start slide index (0-based, optional, for split)")]
        int? startSlideIndex = null,
        [Description("End slide index (0-based, optional, for split)")]
        int? endSlideIndex = null,
        [Description(
            "Output file name pattern, use {index} for slide number (optional, for split, default: 'slide_{index}.pptx')")]
        string outputFileNamePattern = "slide_{index}.pptx")
    {
        var parameters = BuildParameters(operation, sessionId, path, outputPath, inputPath, outputDirectory,
            format, slideIndex, inputPaths, keepSourceFormatting, slidesPerFile,
            startSlideIndex, endSlideIndex, outputFileNamePattern);

        var handler = _handlerRegistry.GetHandler(operation);

        var operationContext = new OperationContext<Presentation>
        {
            Document = null!,
            SessionManager = _sessionManager,
            IdentityAccessor = _identityAccessor,
            SessionId = sessionId,
            SourcePath = inputPath ?? path,
            OutputPath = outputPath
        };

        return handler.Execute(operationContext, parameters);
    }

    /// <summary>
    ///     Builds OperationParameters from method parameters.
    /// </summary>
    private static OperationParameters BuildParameters(
        string operation,
        string? sessionId,
        string? path,
        string? outputPath,
        string? inputPath,
        string? outputDirectory,
        string? format,
        int slideIndex,
        string[]? inputPaths,
        bool keepSourceFormatting,
        int slidesPerFile,
        int? startSlideIndex,
        int? endSlideIndex,
        string outputFileNamePattern)
    {
        var parameters = new OperationParameters();

        switch (operation.ToLowerInvariant())
        {
            case "create":
                if (path != null) parameters.Set("path", path);
                if (outputPath != null) parameters.Set("outputPath", outputPath);
                break;

            case "convert":
                if (inputPath != null) parameters.Set("inputPath", inputPath);
                if (path != null) parameters.Set("path", path);
                if (sessionId != null) parameters.Set("sessionId", sessionId);
                if (outputPath != null) parameters.Set("outputPath", outputPath);
                if (format != null) parameters.Set("format", format);
                parameters.Set("slideIndex", slideIndex);
                break;

            case "merge":
                if (path != null) parameters.Set("path", path);
                if (outputPath != null) parameters.Set("outputPath", outputPath);
                if (inputPaths != null) parameters.Set("inputPaths", inputPaths);
                parameters.Set("keepSourceFormatting", keepSourceFormatting);
                break;

            case "split":
                if (inputPath != null) parameters.Set("inputPath", inputPath);
                if (path != null) parameters.Set("path", path);
                if (sessionId != null) parameters.Set("sessionId", sessionId);
                if (outputDirectory != null) parameters.Set("outputDirectory", outputDirectory);
                parameters.Set("slidesPerFile", slidesPerFile);
                if (startSlideIndex.HasValue) parameters.Set("startSlideIndex", startSlideIndex.Value);
                if (endSlideIndex.HasValue) parameters.Set("endSlideIndex", endSlideIndex.Value);
                parameters.Set("outputFileNamePattern", outputFileNamePattern);
                break;
        }

        return parameters;
    }
}
