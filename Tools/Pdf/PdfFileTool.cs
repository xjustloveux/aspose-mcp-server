using System.ComponentModel;
using Aspose.Pdf;
using Aspose.Pdf.Optimization;
using AsposeMcpServer.Core.Helpers;
using AsposeMcpServer.Core.Session;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.Pdf;

/// <summary>
///     Tool for performing file operations on PDF documents (create, merge, split, compress, encrypt, linearize)
/// </summary>
[McpServerToolType]
public class PdfFileTool
{
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
    [McpServerTool(Name = "pdf_file")]
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
    public string Execute(
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
        return operation.ToLower() switch
        {
            "create" => CreateDocument(outputPath),
            "merge" => MergeDocuments(outputPath, inputPaths),
            "split" => SplitDocument(sessionId, path, outputDir, pagesPerFile, startPage, endPage),
            "compress" => CompressDocument(sessionId, path, outputPath, compressImages, compressFonts,
                removeUnusedObjects),
            "encrypt" => EncryptDocument(sessionId, path, outputPath, userPassword, ownerPassword),
            "linearize" => LinearizeDocument(sessionId, path, outputPath),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Creates a new empty PDF document with one blank page.
    /// </summary>
    /// <param name="outputPath">The output file path.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when outputPath is not provided.</exception>
    private static string CreateDocument(string? outputPath)
    {
        if (string.IsNullOrEmpty(outputPath))
            throw new ArgumentException("outputPath is required for create operation");

        SecurityHelper.ValidateFilePath(outputPath, "outputPath", true);

        using var document = new Document();
        document.Pages.Add();
        document.Save(outputPath);
        return $"PDF document created. Output: {outputPath}";
    }

    /// <summary>
    ///     Merges multiple PDF documents into a single document.
    /// </summary>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="inputPaths">Array of input PDF file paths to merge.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing.</exception>
    private static string MergeDocuments(string? outputPath, string[]? inputPaths)
    {
        if (string.IsNullOrEmpty(outputPath))
            throw new ArgumentException("outputPath is required for merge operation");
        if (inputPaths == null || inputPaths.Length == 0)
            throw new ArgumentException("inputPaths is required for merge operation");

        SecurityHelper.ValidateArraySize(inputPaths, "inputPaths");

        var validPaths = inputPaths.Where(p => !string.IsNullOrEmpty(p)).ToList();
        if (validPaths.Count == 0)
            throw new ArgumentException("At least one input path is required");

        foreach (var inputPath in validPaths)
            SecurityHelper.ValidateFilePath(inputPath, "inputPaths", true);
        SecurityHelper.ValidateFilePath(outputPath, "outputPath", true);

        using var mergedDocument = new Document(validPaths[0]);
        for (var i = 1; i < validPaths.Count; i++)
        {
            using var doc = new Document(validPaths[i]);
            mergedDocument.Pages.Add(doc.Pages);
        }

        mergedDocument.Save(outputPath);
        return $"Merged {validPaths.Count} PDF documents. Output: {outputPath}";
    }

    /// <summary>
    ///     Splits a PDF document into multiple files.
    /// </summary>
    /// <param name="sessionId">The session ID for in-memory editing.</param>
    /// <param name="path">The PDF file path.</param>
    /// <param name="outputDir">The output directory for split files.</param>
    /// <param name="pagesPerFile">Number of pages per output file.</param>
    /// <param name="startPage">Optional start page for extraction.</param>
    /// <param name="endPage">Optional end page for extraction.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or invalid.</exception>
    private string SplitDocument(string? sessionId, string? path, string? outputDir, int pagesPerFile, int? startPage,
        int? endPage)
    {
        if (string.IsNullOrEmpty(outputDir))
            throw new ArgumentException("outputDir is required for split operation");

        SecurityHelper.ValidateFilePath(outputDir, "outputDir", true);

        if (pagesPerFile < 1 || pagesPerFile > 1000)
            throw new ArgumentException("pagesPerFile must be between 1 and 1000");

        Directory.CreateDirectory(outputDir);

        using var ctx = DocumentContext<Document>.Create(_sessionManager, sessionId, path, _identityAccessor);
        var document = ctx.Document;
        var totalPages = document.Pages.Count;
        var fileBaseName = ctx.IsSession
            ? "document"
            : SecurityHelper.SanitizeFileName(Path.GetFileNameWithoutExtension(path!));
        var fileCount = 0;

        var actualStartPage = startPage ?? 1;
        var actualEndPage = endPage ?? totalPages;

        if (actualStartPage < 1 || actualStartPage > totalPages)
            throw new ArgumentException($"startPage must be between 1 and {totalPages}");
        if (actualEndPage < actualStartPage || actualEndPage > totalPages)
            throw new ArgumentException($"endPage must be between {actualStartPage} and {totalPages}");

        if (startPage.HasValue || endPage.HasValue)
        {
            using var newDocument = new Document();
            for (var pageNum = actualStartPage; pageNum <= actualEndPage; pageNum++)
                newDocument.Pages.Add(document.Pages[pageNum]);

            var safeFileName =
                SecurityHelper.SanitizeFileName($"{fileBaseName}_pages_{actualStartPage}-{actualEndPage}.pdf");
            var splitOutputPath = Path.Combine(outputDir, safeFileName);
            newDocument.Save(splitOutputPath);
            return
                $"PDF extracted pages {actualStartPage}-{actualEndPage} ({actualEndPage - actualStartPage + 1} pages). Output: {splitOutputPath}";
        }

        for (var i = 0; i < totalPages; i += pagesPerFile)
        {
            using var newDocument = new Document();
            for (var j = 0; j < pagesPerFile && i + j < totalPages; j++)
                newDocument.Pages.Add(document.Pages[i + j + 1]);

            var safeFileName = SecurityHelper.SanitizeFileName($"{fileBaseName}_part_{++fileCount}.pdf");
            var splitOutputPath = Path.Combine(outputDir, safeFileName);
            newDocument.Save(splitOutputPath);
        }

        return $"PDF split into {fileCount} files. Output: {outputDir}";
    }

    /// <summary>
    ///     Compresses a PDF document by optimizing images, fonts, and removing unused objects.
    /// </summary>
    /// <param name="sessionId">The session ID for in-memory editing.</param>
    /// <param name="path">The PDF file path.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="compressImages">Whether to compress images.</param>
    /// <param name="compressFonts">Whether to subset fonts.</param>
    /// <param name="removeUnusedObjects">Whether to remove unused objects.</param>
    /// <returns>A message indicating the result and compression statistics.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing.</exception>
    private string CompressDocument(string? sessionId, string? path, string? outputPath, bool compressImages,
        bool compressFonts, bool removeUnusedObjects)
    {
        if (string.IsNullOrEmpty(outputPath) && string.IsNullOrEmpty(sessionId))
            throw new ArgumentException("outputPath is required for compress operation in file mode");

        using var ctx = DocumentContext<Document>.Create(_sessionManager, sessionId, path, _identityAccessor);
        var document = ctx.Document;
        var optimizationOptions = new OptimizationOptions();

        if (compressImages)
        {
            optimizationOptions.ImageCompressionOptions.CompressImages = true;
            optimizationOptions.ImageCompressionOptions.ImageQuality = 75;
        }

        if (compressFonts)
            optimizationOptions.SubsetFonts = true;

        if (removeUnusedObjects)
        {
            optimizationOptions.LinkDuplcateStreams = true;
            optimizationOptions.RemoveUnusedObjects = true;
            optimizationOptions.AllowReusePageContent = true;
        }

        document.OptimizeResources(optimizationOptions);

        ctx.Save(outputPath);

        if (ctx.IsSession) return $"PDF compressed. {ctx.GetOutputMessage(outputPath)}";

        var originalSize = new FileInfo(path!).Length;
        var compressedSize = new FileInfo(outputPath!).Length;
        var reduction = (double)(originalSize - compressedSize) / originalSize * 100;

        return
            $"PDF compressed ({reduction:F2}% reduction, {originalSize} -> {compressedSize} bytes). {ctx.GetOutputMessage(outputPath)}";
    }

    /// <summary>
    ///     Encrypts a PDF document with user and owner passwords.
    /// </summary>
    /// <param name="sessionId">The session ID for in-memory editing.</param>
    /// <param name="path">The PDF file path.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="userPassword">The user password for opening the document.</param>
    /// <param name="ownerPassword">The owner password for full access.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when passwords are not provided.</exception>
    private string EncryptDocument(string? sessionId, string? path, string? outputPath, string? userPassword,
        string? ownerPassword)
    {
        if (string.IsNullOrEmpty(userPassword))
            throw new ArgumentException("userPassword is required for encrypt operation");
        if (string.IsNullOrEmpty(ownerPassword))
            throw new ArgumentException("ownerPassword is required for encrypt operation");

        using var ctx = DocumentContext<Document>.Create(_sessionManager, sessionId, path, _identityAccessor);
        var document = ctx.Document;
        document.Encrypt(userPassword, ownerPassword, Permissions.PrintDocument | Permissions.ModifyContent,
            CryptoAlgorithm.AESx256);
        ctx.Save(outputPath);
        return $"PDF encrypted with password. {ctx.GetOutputMessage(outputPath)}";
    }

    /// <summary>
    ///     Linearizes a PDF document for fast web viewing.
    /// </summary>
    /// <param name="sessionId">The session ID for in-memory editing.</param>
    /// <param name="path">The PDF file path.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <returns>A message indicating the result and file size information.</returns>
    private string LinearizeDocument(string? sessionId, string? path, string? outputPath)
    {
        using var ctx = DocumentContext<Document>.Create(_sessionManager, sessionId, path, _identityAccessor);
        var document = ctx.Document;
        document.Optimize();
        ctx.Save(outputPath);

        if (ctx.IsSession) return $"PDF linearized for fast web view. {ctx.GetOutputMessage(outputPath)}";

        var originalSize = new FileInfo(path!).Length;
        var optimizedSize = new FileInfo(outputPath!).Length;

        return
            $"PDF linearized for fast web view. Original: {originalSize} bytes, Optimized: {optimizedSize} bytes. {ctx.GetOutputMessage(outputPath)}";
    }
}