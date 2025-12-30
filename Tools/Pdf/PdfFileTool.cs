using System.Text.Json.Nodes;
using Aspose.Pdf;
using Aspose.Pdf.Optimization;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools.Pdf;

/// <summary>
///     Tool for performing file operations on PDF documents (create, merge, split, compress, encrypt, linearize)
/// </summary>
public class PdfFileTool : IAsposeTool
{
    public string Description =>
        @"Perform file operations on PDF documents. Supports 6 operations: create, merge, split, compress, encrypt, linearize.

Usage examples:
- Create PDF: pdf_file(operation='create', outputPath='new.pdf')
- Merge PDFs: pdf_file(operation='merge', inputPaths=['doc1.pdf','doc2.pdf'], outputPath='merged.pdf')
- Split PDF: pdf_file(operation='split', path='doc.pdf', outputDir='output/', pagesPerFile=1)
- Split PDF (page range): pdf_file(operation='split', path='doc.pdf', outputDir='output/', startPage=2, endPage=5)
- Compress PDF: pdf_file(operation='compress', path='doc.pdf', outputPath='compressed.pdf', compressImages=true)
- Encrypt PDF: pdf_file(operation='encrypt', path='doc.pdf', outputPath='encrypted.pdf', userPassword='user', ownerPassword='owner')
- Linearize PDF: pdf_file(operation='linearize', path='doc.pdf', outputPath='linearized.pdf')";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            operation = new
            {
                type = "string",
                description = @"Operation to perform.
- 'create': Create a new PDF (required params: outputPath)
- 'merge': Merge multiple PDFs (required params: inputPaths, outputPath)
- 'split': Split PDF into multiple files (required params: path, outputDir; optional: startPage, endPage, pagesPerFile)
- 'compress': Compress PDF file (required params: path, outputPath)
- 'encrypt': Encrypt PDF file (required params: path, outputPath, userPassword, ownerPassword)
- 'linearize': Optimize PDF for fast web view (required params: path, outputPath)",
                @enum = new[] { "create", "merge", "split", "compress", "encrypt", "linearize" }
            },
            path = new
            {
                type = "string",
                description = "Input file path (required for split, compress, encrypt, and linearize operations)"
            },
            outputPath = new
            {
                type = "string",
                description =
                    "Output file path (required for create, merge, compress, encrypt, and linearize operations)"
            },
            inputPaths = new
            {
                type = "array",
                description = "Array of input file paths to merge (required for merge)",
                items = new { type = "string" }
            },
            outputDir = new
            {
                type = "string",
                description = "Output directory for split files (required for split)"
            },
            pagesPerFile = new
            {
                type = "number",
                description = "Number of pages per file (for split, default: 1, ignored if startPage/endPage specified)"
            },
            startPage = new
            {
                type = "number",
                description = "Start page number, 1-based (for split, optional)"
            },
            endPage = new
            {
                type = "number",
                description = "End page number, 1-based inclusive (for split, optional)"
            },
            compressImages = new
            {
                type = "boolean",
                description = "Compress images (for compress, default: true)"
            },
            compressFonts = new
            {
                type = "boolean",
                description = "Compress fonts (for compress, default: true)"
            },
            removeUnusedObjects = new
            {
                type = "boolean",
                description = "Remove unused objects (for compress, default: true)"
            },
            userPassword = new
            {
                type = "string",
                description = "User password for opening PDF (required for encrypt)"
            },
            ownerPassword = new
            {
                type = "string",
                description = "Owner password for permissions control (required for encrypt)"
            }
        },
        required = new[] { "operation" }
    };

    /// <summary>
    ///     Executes the tool operation with the provided JSON arguments
    /// </summary>
    /// <param name="arguments">JSON arguments object containing operation parameters</param>
    /// <returns>Result message as a string</returns>
    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var operation = ArgumentHelper.GetString(arguments, "operation");

        // Get path and outputPath based on operation type
        string? path = null;
        string? outputPath = null;

        switch (operation.ToLower())
        {
            case "create":
                outputPath = ArgumentHelper.GetString(arguments, "outputPath", "path", "outputPath");
                break;
            case "merge":
                outputPath = ArgumentHelper.GetString(arguments, "outputPath");
                break;
            case "split":
                path = ArgumentHelper.GetAndValidatePath(arguments);
                break;
            case "compress":
            case "encrypt":
            case "linearize":
                path = ArgumentHelper.GetAndValidatePath(arguments);
                outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);
                break;
        }

        return operation.ToLower() switch
        {
            "create" => await CreateDocument(outputPath!),
            "merge" => await MergeDocuments(outputPath!, arguments),
            "split" => await SplitDocument(path!, arguments),
            "compress" => await CompressDocument(path!, outputPath!, arguments),
            "encrypt" => await EncryptDocument(path!, outputPath!, arguments),
            "linearize" => await LinearizeDocument(path!, outputPath!),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Creates a new PDF document
    /// </summary>
    /// <param name="outputPath">Output file path</param>
    /// <returns>Success message with file path</returns>
    private Task<string> CreateDocument(string outputPath)
    {
        return Task.Run(() =>
        {
            SecurityHelper.ValidateFilePath(outputPath, "outputPath", true);

            using var document = new Document();
            document.Pages.Add();
            document.Save(outputPath);
            return $"PDF document created. Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Merges multiple PDF documents into one
    /// </summary>
    /// <param name="outputPath">Output file path</param>
    /// <param name="arguments">JSON arguments containing inputPaths array</param>
    /// <returns>Success message with merged file path</returns>
    private Task<string> MergeDocuments(string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var inputPathsArray = ArgumentHelper.GetArray(arguments, "inputPaths");

            // Validate array size
            SecurityHelper.ValidateArraySize(inputPathsArray, "inputPaths");

            var inputPaths = inputPathsArray.Select(p => p?.GetValue<string>()).Where(p => p != null).ToList();
            if (inputPaths.Count == 0)
                throw new ArgumentException("At least one input path is required");

            foreach (var inputPath in inputPaths) SecurityHelper.ValidateFilePath(inputPath!, "inputPaths", true);
            SecurityHelper.ValidateFilePath(outputPath, "outputPath", true);

            using var mergedDocument = new Document(inputPaths[0]);
            for (var i = 1; i < inputPaths.Count; i++)
            {
                using var doc = new Document(inputPaths[i]);
                mergedDocument.Pages.Add(doc.Pages);
            }

            mergedDocument.Save(outputPath);
            return $"Merged {inputPaths.Count} PDF documents. Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Splits PDF into multiple files
    /// </summary>
    /// <param name="path">Input file path</param>
    /// <param name="arguments">JSON arguments containing outputDir, pagesPerFile, startPage, endPage</param>
    /// <returns>Success message with split file count</returns>
    private Task<string> SplitDocument(string path, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var outputDir = ArgumentHelper.GetString(arguments, "outputDir");
            var pagesPerFile = ArgumentHelper.GetInt(arguments, "pagesPerFile", 1);
            var startPage = ArgumentHelper.GetIntNullable(arguments, "startPage");
            var endPage = ArgumentHelper.GetIntNullable(arguments, "endPage");

            SecurityHelper.ValidateFilePath(outputDir, "outputDir", true);

            if (pagesPerFile < 1 || pagesPerFile > 1000)
                throw new ArgumentException("pagesPerFile must be between 1 and 1000");

            Directory.CreateDirectory(outputDir);
            using var document = new Document(path);
            var totalPages = document.Pages.Count;
            var fileBaseName = SecurityHelper.SanitizeFileName(Path.GetFileNameWithoutExtension(path));
            var fileCount = 0;

            // Determine page range
            var actualStartPage = startPage ?? 1;
            var actualEndPage = endPage ?? totalPages;

            if (actualStartPage < 1 || actualStartPage > totalPages)
                throw new ArgumentException($"startPage must be between 1 and {totalPages}");
            if (actualEndPage < actualStartPage || actualEndPage > totalPages)
                throw new ArgumentException($"endPage must be between {actualStartPage} and {totalPages}");

            // If startPage/endPage specified, extract as single file
            if (startPage.HasValue || endPage.HasValue)
            {
                using var newDocument = new Document();
                for (var pageNum = actualStartPage; pageNum <= actualEndPage; pageNum++)
                    newDocument.Pages.Add(document.Pages[pageNum]);

                var safeFileName =
                    SecurityHelper.SanitizeFileName($"{fileBaseName}_pages_{actualStartPage}-{actualEndPage}.pdf");
                var outputPath = Path.Combine(outputDir, safeFileName);
                newDocument.Save(outputPath);
                return
                    $"PDF extracted pages {actualStartPage}-{actualEndPage} ({actualEndPage - actualStartPage + 1} pages). Output: {outputPath}";
            }

            // Original split by pagesPerFile
            for (var i = 0; i < totalPages; i += pagesPerFile)
            {
                using var newDocument = new Document();
                for (var j = 0; j < pagesPerFile && i + j < totalPages; j++)
                    newDocument.Pages.Add(document.Pages[i + j + 1]);

                var safeFileName = SecurityHelper.SanitizeFileName($"{fileBaseName}_part_{++fileCount}.pdf");
                var outputPath = Path.Combine(outputDir, safeFileName);
                newDocument.Save(outputPath);
            }

            return $"PDF split into {fileCount} files. Output: {outputDir}";
        });
    }

    /// <summary>
    ///     Compresses a PDF document
    /// </summary>
    /// <param name="path">Input file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="arguments">JSON arguments containing compression options</param>
    /// <returns>Success message</returns>
    private Task<string> CompressDocument(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var compressImages = ArgumentHelper.GetBool(arguments, "compressImages", true);
            var compressFonts = ArgumentHelper.GetBool(arguments, "compressFonts", true);
            var removeUnusedObjects = ArgumentHelper.GetBool(arguments, "removeUnusedObjects", true);

            using var document = new Document(path);
            var optimizationOptions = new OptimizationOptions();

            if (compressImages)
            {
                optimizationOptions.ImageCompressionOptions.CompressImages = true;
                optimizationOptions.ImageCompressionOptions.ImageQuality = 75;
            }

            if (compressFonts) optimizationOptions.SubsetFonts = true;

            if (removeUnusedObjects)
            {
                optimizationOptions.LinkDuplcateStreams = true;
                optimizationOptions.RemoveUnusedObjects = true;
                optimizationOptions.AllowReusePageContent = true;
            }

            document.OptimizeResources(optimizationOptions);
            document.Save(outputPath);

            var originalSize = new FileInfo(path).Length;
            var compressedSize = new FileInfo(outputPath).Length;
            var reduction = (double)(originalSize - compressedSize) / originalSize * 100;

            return
                $"PDF compressed ({reduction:F2}% reduction, {originalSize} -> {compressedSize} bytes). Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Encrypts a PDF document
    /// </summary>
    /// <param name="path">Input file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="arguments">JSON arguments containing userPassword, ownerPassword</param>
    /// <returns>Success message</returns>
    private Task<string> EncryptDocument(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var userPassword = ArgumentHelper.GetString(arguments, "userPassword");
            var ownerPassword = ArgumentHelper.GetString(arguments, "ownerPassword");

            using var document = new Document(path);
            document.Encrypt(userPassword, ownerPassword, Permissions.PrintDocument | Permissions.ModifyContent,
                CryptoAlgorithm.AESx256);
            document.Save(outputPath);
            return $"PDF encrypted with password. Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Linearizes a PDF document for fast web view
    /// </summary>
    /// <param name="path">Input file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <returns>Success message</returns>
    private Task<string> LinearizeDocument(string path, string outputPath)
    {
        return Task.Run(() =>
        {
            using var document = new Document(path);
            document.Optimize();
            document.Save(outputPath);

            var originalSize = new FileInfo(path).Length;
            var optimizedSize = new FileInfo(outputPath).Length;

            return
                $"PDF linearized for fast web view. Original: {originalSize} bytes, Optimized: {optimizedSize} bytes. Output: {outputPath}";
        });
    }
}