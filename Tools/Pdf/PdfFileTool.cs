using System.ComponentModel;
using Aspose.Pdf;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Session;
using AsposeMcpServer.Helpers;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.Pdf;

/// <summary>
///     Tool for performing file operations on PDF documents (create, merge, split, compress, encrypt, linearize)
/// </summary>
[ToolHandlerMapping("AsposeMcpServer.Handlers.Pdf.FileOperations")]
[McpServerToolType]
public class PdfFileTool
{
    /// <summary>
    ///     Handler registry for file operations.
    /// </summary>
    private readonly HandlerRegistry<Document> _handlerRegistry;

    /// <summary>
    ///     The session identity accessor for session isolation.
    /// </summary>
    private readonly ISessionIdentityAccessor? _identityAccessor;

    /// <summary>
    ///     The document session manager for managing in-memory document sessions.
    /// </summary>
    private readonly DocumentSessionManager? _sessionManager;

    /// <summary>
    ///     Initializes a new instance of the <see cref="PdfFileTool" /> class.
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory document editing.</param>
    /// <param name="identityAccessor">Optional session identity accessor for session isolation.</param>
    public PdfFileTool(DocumentSessionManager? sessionManager = null, ISessionIdentityAccessor? identityAccessor = null)
    {
        _sessionManager = sessionManager;
        _identityAccessor = identityAccessor;
        _handlerRegistry = HandlerRegistry<Document>.CreateFromNamespace("AsposeMcpServer.Handlers.Pdf.FileOperations");
    }

    /// <summary>
    ///     Executes a PDF file operation (create, merge, split, compress, encrypt, linearize).
    /// </summary>
    /// <param name="operation">The operation to perform: create, merge, split, compress, encrypt, linearize.</param>
    /// <param name="path">Input file path (required for split, compress, encrypt, and linearize operations).</param>
    /// <param name="sessionId">Session ID for in-memory editing.</param>
    /// <param name="outputPath">Output file path (required for create, merge, compress, encrypt, and linearize operations).</param>
    /// <param name="inputPaths">Array of input file paths to merge (required for merge).</param>
    /// <param name="outputDir">Output directory for split files (required for split).</param>
    /// <param name="pagesPerFile">Number of pages per file (for split, default: 1).</param>
    /// <param name="startPage">Start page number, 1-based (for split, optional).</param>
    /// <param name="endPage">End page number, 1-based inclusive (for split, optional).</param>
    /// <param name="compressImages">Compress images (for compress, default: true).</param>
    /// <param name="compressFonts">Compress fonts (for compress, default: true).</param>
    /// <param name="removeUnusedObjects">Remove unused objects (for compress, default: true).</param>
    /// <param name="userPassword">User password for opening PDF (required for encrypt).</param>
    /// <param name="ownerPassword">Owner password for permissions control (required for encrypt).</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or the operation is unknown.</exception>
    [McpServerTool(
        Name = "pdf_file",
        Title = "PDF File Operations",
        Destructive = false,
        Idempotent = true,
        OpenWorld = false,
        ReadOnly = false,
        UseStructuredContent = true)]
    [Description(
        @"Perform file operations on PDF documents. Supports 6 operations: create, merge, split, compress, encrypt, linearize.

Usage examples:
- Create PDF: pdf_file(operation='create', outputPath='new.pdf')
- Merge PDFs: pdf_file(operation='merge', inputPaths=['doc1.pdf','doc2.pdf'], outputPath='merged.pdf')
- Split PDF: pdf_file(operation='split', path='doc.pdf', outputDir='output/', pagesPerFile=1)
- Split PDF (page range): pdf_file(operation='split', path='doc.pdf', outputDir='output/', startPage=2, endPage=5)
- Compress PDF: pdf_file(operation='compress', path='doc.pdf', outputPath='compressed.pdf', compressImages=true)
- Encrypt PDF: pdf_file(operation='encrypt', path='doc.pdf', outputPath='encrypted.pdf', userPassword='user', ownerPassword='owner')
- Linearize PDF: pdf_file(operation='linearize', path='doc.pdf', outputPath='linearized.pdf')")]
    public object Execute(
        [Description(@"Operation to perform.
- 'create': Create a new PDF (required params: outputPath)
- 'merge': Merge multiple PDFs (required params: inputPaths, outputPath)
- 'split': Split PDF into multiple files (required params: path, outputDir; optional: startPage, endPage, pagesPerFile)
- 'compress': Compress PDF file (required params: path, outputPath)
- 'encrypt': Encrypt PDF file (required params: path, outputPath, userPassword, ownerPassword)
- 'linearize': Optimize PDF for fast web view (required params: path, outputPath)")]
        string operation,
        [Description("Input file path (required for split, compress, encrypt, and linearize operations)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Output file path (required for create, merge, compress, encrypt, and linearize operations)")]
        string? outputPath = null,
        [Description("Array of input file paths to merge (required for merge)")]
        string[]? inputPaths = null,
        [Description("Output directory for split files (required for split)")]
        string? outputDir = null,
        [Description("Number of pages per file (for split, default: 1, ignored if startPage/endPage specified)")]
        int pagesPerFile = 1,
        [Description("Start page number, 1-based (for split, optional)")]
        int? startPage = null,
        [Description("End page number, 1-based inclusive (for split, optional)")]
        int? endPage = null,
        [Description("Compress images (for compress, default: true)")]
        bool compressImages = true,
        [Description("Compress fonts (for compress, default: true)")]
        bool compressFonts = true,
        [Description("Remove unused objects (for compress, default: true)")]
        bool removeUnusedObjects = true,
        [Description("User password for opening PDF (required for encrypt)")]
        string? userPassword = null,
        [Description("Owner password for permissions control (required for encrypt)")]
        string? ownerPassword = null)
    {
        var lowerOperation = operation.ToLowerInvariant();

        if (lowerOperation is "create" or "merge")
        {
            var message = ExecuteWithoutContext(lowerOperation, outputPath, inputPaths);
            return ResultHelper.FinalizeResult((dynamic)message, outputPath, sessionId);
        }

        var result = ExecuteWithContext(lowerOperation, path, sessionId, outputPath, outputDir, pagesPerFile, startPage,
            endPage, compressImages, compressFonts, removeUnusedObjects, userPassword, ownerPassword);
        return ResultHelper.FinalizeResult((dynamic)result, outputPath, sessionId);
    }

    /// <summary>
    ///     Executes operations that don't require an existing document context.
    /// </summary>
    private object ExecuteWithoutContext(string operation, string? outputPath, string[]? inputPaths)
    {
        var parameters = operation switch
        {
            "create" => BuildCreateParameters(outputPath),
            "merge" => BuildMergeParameters(outputPath, inputPaths),
            _ => new OperationParameters()
        };

        var handler = _handlerRegistry.GetHandler(operation);

        var operationContext = new OperationContext<Document>
        {
            Document = null!,
            SessionManager = _sessionManager,
            IdentityAccessor = _identityAccessor,
            OutputPath = outputPath
        };

        return handler.Execute(operationContext, parameters);
    }

    /// <summary>
    ///     Builds parameters for the create operation.
    /// </summary>
    /// <param name="outputPath">The output file path for the created PDF.</param>
    /// <returns>OperationParameters configured for creating a new PDF.</returns>
    private static OperationParameters BuildCreateParameters(string? outputPath)
    {
        var parameters = new OperationParameters();
        if (outputPath != null) parameters.Set("outputPath", outputPath);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the merge operation.
    /// </summary>
    /// <param name="outputPath">The output file path for the merged PDF.</param>
    /// <param name="inputPaths">Array of input file paths to merge.</param>
    /// <returns>OperationParameters configured for merging PDFs.</returns>
    private static OperationParameters BuildMergeParameters(string? outputPath, string[]? inputPaths)
    {
        var parameters = new OperationParameters();
        if (outputPath != null) parameters.Set("outputPath", outputPath);
        if (inputPaths != null) parameters.Set("inputPaths", inputPaths);
        return parameters;
    }

    /// <summary>
    ///     Executes operations that require an existing document context.
    /// </summary>
    private object ExecuteWithContext(string operation, string? path, string? sessionId, string? outputPath,
        string? outputDir, int pagesPerFile, int? startPage, int? endPage, bool compressImages, bool compressFonts,
        bool removeUnusedObjects, string? userPassword, string? ownerPassword)
    {
        using var ctx = DocumentContext<Document>.Create(_sessionManager, sessionId, path, _identityAccessor);

        var parameters = BuildParameters(operation, outputDir, pagesPerFile, startPage, endPage,
            compressImages, compressFonts, removeUnusedObjects, userPassword, ownerPassword,
            ctx.IsSession ? null : path);

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

        if (operationContext.IsModified)
            ctx.Save(outputPath);

        // For compress/linearize file operations, return detailed size info
        if (operation == "compress" && !ctx.IsSession && path != null && outputPath != null)
        {
            var originalSize = new FileInfo(path).Length;
            var compressedSize = new FileInfo(outputPath).Length;
            var reduction = (double)(originalSize - compressedSize) / originalSize * 100;
            return $"PDF compressed ({reduction:F2}% reduction, {originalSize} -> {compressedSize} bytes)";
        }

        if (operation == "linearize" && !ctx.IsSession && path != null && outputPath != null)
        {
            var originalSize = new FileInfo(path).Length;
            var optimizedSize = new FileInfo(outputPath).Length;
            return
                $"PDF linearized for fast web view. Original: {originalSize} bytes, Optimized: {optimizedSize} bytes";
        }

        return result;
    }

    /// <summary>
    ///     Builds OperationParameters from method parameters using strategy pattern.
    ///     Parameters are documented on the Execute method.
    /// </summary>
    /// <returns>OperationParameters configured with all input values.</returns>
    private static OperationParameters BuildParameters(
        string operation,
        string? outputDir,
        int pagesPerFile,
        int? startPage,
        int? endPage,
        bool compressImages,
        bool compressFonts,
        bool removeUnusedObjects,
        string? userPassword,
        string? ownerPassword,
        string? fileBaseName)
    {
        return operation switch
        {
            "split" => BuildSplitParameters(outputDir, pagesPerFile, startPage, endPage, fileBaseName),
            "compress" => BuildCompressParameters(compressImages, compressFonts, removeUnusedObjects),
            "encrypt" => BuildEncryptParameters(userPassword, ownerPassword),
            _ => new OperationParameters()
        };
    }

    /// <summary>
    ///     Builds parameters for the split operation.
    /// </summary>
    /// <param name="outputDir">The output directory for split files.</param>
    /// <param name="pagesPerFile">Number of pages per split file.</param>
    /// <param name="startPage">The starting page number (1-based).</param>
    /// <param name="endPage">The ending page number (1-based, inclusive).</param>
    /// <param name="fileBaseName">The base name for output files.</param>
    /// <returns>OperationParameters configured for splitting a PDF.</returns>
    private static OperationParameters BuildSplitParameters(string? outputDir, int pagesPerFile, int? startPage,
        int? endPage, string? fileBaseName)
    {
        var parameters = new OperationParameters();
        if (outputDir != null) parameters.Set("outputDir", outputDir);
        parameters.Set("pagesPerFile", pagesPerFile);
        if (startPage.HasValue) parameters.Set("startPage", startPage.Value);
        if (endPage.HasValue) parameters.Set("endPage", endPage.Value);
        if (fileBaseName != null) parameters.Set("fileBaseName", fileBaseName);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the compress operation.
    /// </summary>
    /// <param name="compressImages">Whether to compress images.</param>
    /// <param name="compressFonts">Whether to compress fonts.</param>
    /// <param name="removeUnusedObjects">Whether to remove unused objects.</param>
    /// <returns>OperationParameters configured for compressing a PDF.</returns>
    private static OperationParameters BuildCompressParameters(bool compressImages, bool compressFonts,
        bool removeUnusedObjects)
    {
        var parameters = new OperationParameters();
        parameters.Set("compressImages", compressImages);
        parameters.Set("compressFonts", compressFonts);
        parameters.Set("removeUnusedObjects", removeUnusedObjects);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the encrypt operation.
    /// </summary>
    /// <param name="userPassword">The user password for opening the PDF.</param>
    /// <param name="ownerPassword">The owner password for permissions control.</param>
    /// <returns>OperationParameters configured for encrypting a PDF.</returns>
    private static OperationParameters BuildEncryptParameters(string? userPassword, string? ownerPassword)
    {
        var parameters = new OperationParameters();
        if (userPassword != null) parameters.Set("userPassword", userPassword);
        if (ownerPassword != null) parameters.Set("ownerPassword", ownerPassword);
        return parameters;
    }
}
