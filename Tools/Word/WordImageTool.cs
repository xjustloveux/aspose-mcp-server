using System.ComponentModel;
using Aspose.Words;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Session;
using AsposeMcpServer.Helpers;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.Word;

/// <summary>
///     Unified tool for managing Word images (add, edit, delete, get, replace, extract)
///     Merges: WordAddImageTool, WordEditImageTool, WordDeleteImageTool, WordGetImagesTool, WordReplaceImageTool,
///     WordExtractImagesTool
/// </summary>
[ToolHandlerMapping("AsposeMcpServer.Handlers.Word.Image")]
[McpServerToolType]
public class WordImageTool
{
    /// <summary>
    ///     Handler registry for image operations
    /// </summary>
    private readonly HandlerRegistry<Document> _handlerRegistry;

    /// <summary>
    ///     Identity accessor for session isolation
    /// </summary>
    private readonly ISessionIdentityAccessor? _identityAccessor;

    /// <summary>
    ///     Session manager for document session operations
    /// </summary>
    private readonly DocumentSessionManager? _sessionManager;

    /// <summary>
    ///     Initializes a new instance of the WordImageTool class
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory document operations</param>
    /// <param name="identityAccessor">Optional identity accessor for session isolation</param>
    public WordImageTool(DocumentSessionManager? sessionManager = null,
        ISessionIdentityAccessor? identityAccessor = null)
    {
        _sessionManager = sessionManager;
        _identityAccessor = identityAccessor;
        _handlerRegistry = HandlerRegistry<Document>.CreateFromNamespace("AsposeMcpServer.Handlers.Word.Image");
    }

    /// <summary>
    ///     Executes a Word image operation (add, edit, delete, get, replace, extract).
    /// </summary>
    /// <param name="operation">The operation to perform: add, edit, delete, get, replace, extract.</param>
    /// <param name="path">Word document file path (required if no sessionId).</param>
    /// <param name="sessionId">Session ID for in-memory editing.</param>
    /// <param name="outputPath">Output file path (file mode only).</param>
    /// <param name="outputDir">Output directory (for extract).</param>
    /// <param name="imagePath">Image file path (for add/replace).</param>
    /// <param name="imageIndex">Image index (0-based, for edit/delete/replace).</param>
    /// <param name="sectionIndex">Section index (0-based, use -1 for all sections).</param>
    /// <param name="width">Image width in points.</param>
    /// <param name="height">Image height in points.</param>
    /// <param name="alignment">Horizontal alignment: left, center, right.</param>
    /// <param name="textWrapping">Text wrapping: inline, square, tight, through, topAndBottom, none.</param>
    /// <param name="caption">Image caption text (for add).</param>
    /// <param name="captionPosition">Caption position: above, below (for add).</param>
    /// <param name="aspectRatioLocked">Lock aspect ratio (for edit).</param>
    /// <param name="horizontalAlignment">Horizontal alignment for floating images (for edit).</param>
    /// <param name="verticalAlignment">Vertical alignment for floating images (for edit).</param>
    /// <param name="alternativeText">Alternative text for accessibility.</param>
    /// <param name="title">Image title.</param>
    /// <param name="linkUrl">Hyperlink URL for the image.</param>
    /// <param name="newImagePath">New image file path for replace operation.</param>
    /// <param name="preserveSize">Preserve original image size for replace operation.</param>
    /// <param name="smartFit">Smart fit to avoid distortion for replace operation.</param>
    /// <param name="preservePosition">Preserve original image position for replace operation.</param>
    /// <param name="prefix">Filename prefix for extracted images.</param>
    /// <param name="extractImageIndex">Specific image index to extract.</param>
    /// <returns>A message indicating the result of the operation, or JSON data for get operations.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or the operation is unknown.</exception>
    [McpServerTool(
        Name = "word_image",
        Title = "Word Image Operations",
        Destructive = true,
        Idempotent = false,
        OpenWorld = false,
        ReadOnly = false,
        UseStructuredContent = true)]
    [Description(@"Manage Word document images. Supports 6 operations: add, edit, delete, get, replace, extract.

Usage examples:
- Add image: word_image(operation='add', path='doc.docx', imagePath='image.png', width=200)
- Edit image: word_image(operation='edit', path='doc.docx', imageIndex=0, width=300, height=200)
- Delete image: word_image(operation='delete', path='doc.docx', imageIndex=0)
- Get all images: word_image(operation='get', path='doc.docx')
- Replace image: word_image(operation='replace', path='doc.docx', imageIndex=0, imagePath='new_image.png')
- Extract images: word_image(operation='extract', path='doc.docx', outputDir='images/')")]
    public object Execute(
        [Description(@"Operation to perform.
- 'add': Add a new image (required params: path, imagePath)
- 'edit': Edit existing image (required params: path, imageIndex)
- 'delete': Delete an image (required params: path, imageIndex)
- 'get': Get all images info (required params: path)
- 'replace': Replace an image (required params: path, imageIndex, imagePath)
- 'extract': Extract all images (required params: path, outputDir)")]
        string operation,
        [Description("Document file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Output file path (file mode only)")]
        string? outputPath = null,
        [Description("Output directory (required for extract operation)")]
        string? outputDir = null,
        [Description("Image file path (required for add and replace operations)")]
        string? imagePath = null,
        [Description(
            "Image index (0-based, required for edit, delete, and replace operations). Note: After delete operations, subsequent image indices will shift automatically. Use 'get' operation to refresh indices.")]
        int? imageIndex = null,
        [Description("Section index (0-based, optional, default: 0, use -1 to search all sections)")]
        int sectionIndex = 0,
        [Description("Image width in points (72 pts = 1 inch, optional, for add/edit operations)")]
        double? width = null,
        [Description("Image height in points (72 pts = 1 inch, optional, for add/edit operations)")]
        double? height = null,
        [Description("Horizontal alignment: left, center, right (optional, for add/edit operations)")]
        string alignment = "left",
        [Description(
            "Text wrapping: inline, square, tight, through, topAndBottom, none (optional, for add/edit operations)")]
        string textWrapping = "inline",
        [Description("Image caption text (optional, for add operation)")]
        string? caption = null,
        [Description("Caption position: above, below (optional, for add operation)")]
        string captionPosition = "below",
        [Description("Lock aspect ratio (optional, for edit operation)")]
        bool? aspectRatioLocked = null,
        [Description("Horizontal alignment for floating images: left, center, right (optional, for edit operation)")]
        string? horizontalAlignment = null,
        [Description("Vertical alignment for floating images: top, center, bottom (optional, for edit operation)")]
        string? verticalAlignment = null,
        [Description("Alternative text for accessibility (optional, for add/edit operation)")]
        string? alternativeText = null,
        [Description("Image title (optional, for add/edit operation)")]
        string? title = null,
        [Description(
            "Hyperlink URL for the image. When clicked, opens the specified URL (optional, for add/edit operation). Use empty string to remove existing hyperlink.")]
        string? linkUrl = null,
        [Description("New image file path (required for replace operation)")]
        string? newImagePath = null,
        [Description("Preserve original image size (default: true, for replace operation)")]
        bool preserveSize = true,
        [Description(
            "When true, keeps original width and calculates height proportionally based on new image aspect ratio (avoids distortion when aspect ratios differ, default: false, for replace operation). Only applies when preserveSize is true.")]
        bool smartFit = false,
        [Description("Preserve original image position and wrapping (default: true, for replace operation)")]
        bool preservePosition = true,
        [Description("Filename prefix for extracted images (optional, default: 'image', for extract operation)")]
        string prefix = "image",
        [Description(
            "Specific image index to extract (0-based, optional, for extract operation). If not provided, extracts all images.")]
        int? extractImageIndex = null)
    {
        using var ctx = DocumentContext<Document>.Create(_sessionManager, sessionId, path, _identityAccessor);

        var parameters = BuildParameters(operation, imagePath, imageIndex, sectionIndex, width, height,
            alignment, textWrapping, caption, captionPosition, aspectRatioLocked, horizontalAlignment,
            verticalAlignment, alternativeText, title, linkUrl, newImagePath, preserveSize, smartFit,
            preservePosition, outputDir, prefix, extractImageIndex);

        var handler = _handlerRegistry.GetHandler(operation);

        var operationContext = new OperationContext<Document>
        {
            Document = ctx.Document,
            SessionManager = _sessionManager,
            IdentityAccessor = _identityAccessor,
            SessionId = sessionId,
            SourcePath = path,
            OutputPath = outputPath
        };

        var result = handler.Execute(operationContext, parameters);

        // Read-only operations don't need to save
        if (operation.ToLower() is "get" or "extract")
            return ResultHelper.FinalizeResult((dynamic)result, ctx, outputPath);

        if (operationContext.IsModified)
            ctx.Save(outputPath);

        return ResultHelper.FinalizeResult((dynamic)result, ctx, outputPath);
    }

    /// <summary>
    ///     Builds OperationParameters from method parameters using strategy pattern.
    ///     Parameters are documented on the Execute method.
    /// </summary>
    /// <returns>OperationParameters configured with all input values.</returns>
    private static OperationParameters BuildParameters(
        string operation,
        string? imagePath,
        int? imageIndex,
        int sectionIndex,
        double? width,
        double? height,
        string alignment,
        string textWrapping,
        string? caption,
        string captionPosition,
        bool? aspectRatioLocked,
        string? horizontalAlignment,
        string? verticalAlignment,
        string? alternativeText,
        string? title,
        string? linkUrl,
        string? newImagePath,
        bool preserveSize,
        bool smartFit,
        bool preservePosition,
        string? outputDir,
        string prefix,
        int? extractImageIndex)
    {
        return operation.ToLower() switch
        {
            "add" => BuildAddParameters(imagePath, width, height, alignment, textWrapping, caption, captionPosition,
                linkUrl, alternativeText, title),
            "edit" => BuildEditParameters(imageIndex, sectionIndex, width, height, alignment, textWrapping,
                aspectRatioLocked, horizontalAlignment, verticalAlignment, alternativeText, title, linkUrl),
            "delete" => BuildDeleteParameters(imageIndex, sectionIndex),
            "get" => BuildGetParameters(sectionIndex),
            "replace" => BuildReplaceParameters(imageIndex, imagePath, newImagePath, preserveSize, smartFit,
                preservePosition, sectionIndex),
            "extract" => BuildExtractParameters(outputDir, prefix, extractImageIndex),
            _ => new OperationParameters()
        };
    }

    /// <summary>
    ///     Builds parameters for the add operation.
    /// </summary>
    /// <param name="imagePath">The image file path.</param>
    /// <param name="width">The image width in points.</param>
    /// <param name="height">The image height in points.</param>
    /// <param name="alignment">The horizontal alignment: left, center, right.</param>
    /// <param name="textWrapping">The text wrapping: inline, square, tight, through, topAndBottom, none.</param>
    /// <param name="caption">The image caption text.</param>
    /// <param name="captionPosition">The caption position: above, below.</param>
    /// <param name="linkUrl">The hyperlink URL for the image.</param>
    /// <param name="alternativeText">The alternative text for accessibility.</param>
    /// <param name="title">The image title.</param>
    /// <returns>OperationParameters configured for adding an image.</returns>
    private static OperationParameters BuildAddParameters(
        string? imagePath, double? width, double? height,
        string alignment, string textWrapping, string? caption, string captionPosition, string? linkUrl,
        string? alternativeText, string? title)
    {
        var parameters = new OperationParameters();
        if (imagePath != null) parameters.Set("imagePath", imagePath);
        if (width.HasValue) parameters.Set("width", width.Value);
        if (height.HasValue) parameters.Set("height", height.Value);
        parameters.Set("alignment", alignment);
        parameters.Set("textWrapping", textWrapping);
        if (caption != null) parameters.Set("caption", caption);
        parameters.Set("captionPosition", captionPosition);
        if (linkUrl != null) parameters.Set("linkUrl", linkUrl);
        if (alternativeText != null) parameters.Set("alternativeText", alternativeText);
        if (title != null) parameters.Set("title", title);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the edit operation.
    /// </summary>
    /// <param name="imageIndex">The image index (0-based).</param>
    /// <param name="sectionIndex">The section index (0-based).</param>
    /// <param name="width">The image width in points.</param>
    /// <param name="height">The image height in points.</param>
    /// <param name="alignment">The horizontal alignment: left, center, right.</param>
    /// <param name="textWrapping">The text wrapping: inline, square, tight, through, topAndBottom, none.</param>
    /// <param name="aspectRatioLocked">Whether to lock aspect ratio.</param>
    /// <param name="horizontalAlignment">The horizontal alignment for floating images.</param>
    /// <param name="verticalAlignment">The vertical alignment for floating images.</param>
    /// <param name="alternativeText">The alternative text for accessibility.</param>
    /// <param name="title">The image title.</param>
    /// <param name="linkUrl">The hyperlink URL for the image.</param>
    /// <returns>OperationParameters configured for editing an image.</returns>
    private static OperationParameters BuildEditParameters(
        int? imageIndex, int sectionIndex, double? width,
        double? height, string alignment, string textWrapping, bool? aspectRatioLocked, string? horizontalAlignment,
        string? verticalAlignment, string? alternativeText, string? title, string? linkUrl)
    {
        var parameters = new OperationParameters();
        if (imageIndex.HasValue) parameters.Set("imageIndex", imageIndex.Value);
        parameters.Set("sectionIndex", sectionIndex);
        if (width.HasValue) parameters.Set("width", width.Value);
        if (height.HasValue) parameters.Set("height", height.Value);
        if (alignment != "left") parameters.Set("alignment", alignment);
        if (textWrapping != "inline") parameters.Set("textWrapping", textWrapping);
        if (aspectRatioLocked.HasValue) parameters.Set("aspectRatioLocked", aspectRatioLocked.Value);
        if (horizontalAlignment != null) parameters.Set("horizontalAlignment", horizontalAlignment);
        if (verticalAlignment != null) parameters.Set("verticalAlignment", verticalAlignment);
        if (alternativeText != null) parameters.Set("alternativeText", alternativeText);
        if (title != null) parameters.Set("title", title);
        if (linkUrl != null) parameters.Set("linkUrl", linkUrl);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the delete operation.
    /// </summary>
    /// <param name="imageIndex">The image index (0-based).</param>
    /// <param name="sectionIndex">The section index (0-based).</param>
    /// <returns>OperationParameters configured for deleting an image.</returns>
    private static OperationParameters BuildDeleteParameters(int? imageIndex, int sectionIndex)
    {
        var parameters = new OperationParameters();
        if (imageIndex.HasValue) parameters.Set("imageIndex", imageIndex.Value);
        parameters.Set("sectionIndex", sectionIndex);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the get operation.
    /// </summary>
    /// <param name="sectionIndex">The section index (0-based).</param>
    /// <returns>OperationParameters configured for getting image information.</returns>
    private static OperationParameters BuildGetParameters(int sectionIndex)
    {
        var parameters = new OperationParameters();
        parameters.Set("sectionIndex", sectionIndex);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the replace operation.
    /// </summary>
    /// <param name="imageIndex">The image index (0-based).</param>
    /// <param name="imagePath">The image file path (alternative to newImagePath).</param>
    /// <param name="newImagePath">The new image file path.</param>
    /// <param name="preserveSize">Whether to preserve original image size.</param>
    /// <param name="smartFit">Whether to use smart fit to avoid distortion.</param>
    /// <param name="preservePosition">Whether to preserve original image position.</param>
    /// <param name="sectionIndex">The section index (0-based).</param>
    /// <returns>OperationParameters configured for replacing an image.</returns>
    private static OperationParameters BuildReplaceParameters(int? imageIndex, string? imagePath, string? newImagePath,
        bool preserveSize, bool smartFit, bool preservePosition, int sectionIndex)
    {
        var parameters = new OperationParameters();
        if (imageIndex.HasValue) parameters.Set("imageIndex", imageIndex.Value);
        if (newImagePath != null) parameters.Set("newImagePath", newImagePath);
        else if (imagePath != null) parameters.Set("imagePath", imagePath);
        parameters.Set("preserveSize", preserveSize);
        parameters.Set("smartFit", smartFit);
        parameters.Set("preservePosition", preservePosition);
        parameters.Set("sectionIndex", sectionIndex);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the extract operation.
    /// </summary>
    /// <param name="outputDir">The output directory for extracted images.</param>
    /// <param name="prefix">The filename prefix for extracted images.</param>
    /// <param name="extractImageIndex">The specific image index to extract (0-based).</param>
    /// <returns>OperationParameters configured for extracting images.</returns>
    private static OperationParameters BuildExtractParameters(string? outputDir, string prefix, int? extractImageIndex)
    {
        var parameters = new OperationParameters();
        if (outputDir != null) parameters.Set("outputDir", outputDir);
        parameters.Set("prefix", prefix);
        if (extractImageIndex.HasValue) parameters.Set("extractImageIndex", extractImageIndex.Value);
        return parameters;
    }
}
