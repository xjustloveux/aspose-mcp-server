using System.Text.Json.Nodes;
using Aspose.Pdf;
using Aspose.Pdf.Optimization;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools;

public class PdfFileTool : IAsposeTool
{
    public string Description => @"Perform file operations on PDF documents. Supports 5 operations: create, merge, split, compress, encrypt.

Usage examples:
- Create PDF: pdf_file(operation='create', outputPath='new.pdf')
- Merge PDFs: pdf_file(operation='merge', inputPaths=['doc1.pdf','doc2.pdf'], outputPath='merged.pdf')
- Split PDF: pdf_file(operation='split', path='doc.pdf', outputDir='output/', pagesPerFile=1)
- Compress PDF: pdf_file(operation='compress', path='doc.pdf', outputPath='compressed.pdf', compressImages=true)
- Encrypt PDF: pdf_file(operation='encrypt', path='doc.pdf', outputPath='encrypted.pdf', password='password')";

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
- 'split': Split PDF into multiple files (required params: path, outputDir)
- 'compress': Compress PDF file (required params: path, outputPath)
- 'encrypt': Encrypt PDF file (required params: path, outputPath, password)",
                @enum = new[] { "create", "merge", "split", "compress", "encrypt" }
            },
            path = new
            {
                type = "string",
                description = "Input file path (required for split, compress, and encrypt operations)"
            },
            outputPath = new
            {
                type = "string",
                description = "Output file path (required for create, merge, compress, and encrypt operations)"
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
                description = "Number of pages per file (for split, default: 1)"
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
                description = "User password (required for encrypt)"
            },
            ownerPassword = new
            {
                type = "string",
                description = "Owner password (required for encrypt)"
            }
        },
        required = new[] { "operation" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var operation = arguments?["operation"]?.GetValue<string>() ?? throw new ArgumentException("operation is required");

        return operation.ToLower() switch
        {
            "create" => await CreateDocument(arguments),
            "merge" => await MergeDocuments(arguments),
            "split" => await SplitDocument(arguments),
            "compress" => await CompressDocument(arguments),
            "encrypt" => await EncryptDocument(arguments),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    private async Task<string> CreateDocument(JsonObject? arguments)
    {
        var outputPath = arguments?["outputPath"]?.GetValue<string>() ?? arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("outputPath is required");

        SecurityHelper.ValidateFilePath(outputPath, "outputPath");

        using var document = new Document();
        document.Pages.Add();
        document.Save(outputPath);
        return await Task.FromResult($"PDF document created successfully at: {outputPath}");
    }

    private async Task<string> MergeDocuments(JsonObject? arguments)
    {
        var inputPathsArray = arguments?["inputPaths"]?.AsArray() ?? throw new ArgumentException("inputPaths is required");
        var outputPath = arguments?["outputPath"]?.GetValue<string>() ?? throw new ArgumentException("outputPath is required");

        // Validate array size
        SecurityHelper.ValidateArraySize(inputPathsArray, "inputPaths");

        var inputPaths = inputPathsArray.Select(p => p?.GetValue<string>()).Where(p => p != null).ToList();
        if (inputPaths.Count == 0)
            throw new ArgumentException("At least one input path is required");

        // Validate all input paths
        foreach (var inputPath in inputPaths)
        {
            SecurityHelper.ValidateFilePath(inputPath!, "inputPaths");
        }

        // Validate output path
        SecurityHelper.ValidateFilePath(outputPath, "outputPath");

        using var mergedDocument = new Document(inputPaths[0]);
        for (int i = 1; i < inputPaths.Count; i++)
        {
            using var doc = new Document(inputPaths[i]);
            mergedDocument.Pages.Add(doc.Pages);
        }

        mergedDocument.Save(outputPath);
        return await Task.FromResult($"Merged {inputPaths.Count} PDF documents into: {outputPath}");
    }

    private async Task<string> SplitDocument(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var outputDir = arguments?["outputDir"]?.GetValue<string>() ?? throw new ArgumentException("outputDir is required");
        var pagesPerFile = arguments?["pagesPerFile"]?.GetValue<int>() ?? 1;

        SecurityHelper.ValidateFilePath(path, "path");
        SecurityHelper.ValidateFilePath(outputDir, "outputDir");

        if (pagesPerFile < 1 || pagesPerFile > 1000)
            throw new ArgumentException("pagesPerFile must be between 1 and 1000");

        Directory.CreateDirectory(outputDir);
        using var document = new Document(path);
        var fileBaseName = SecurityHelper.SanitizeFileName(Path.GetFileNameWithoutExtension(path));
        int fileCount = 0;

        for (int i = 0; i < document.Pages.Count; i += pagesPerFile)
        {
            using var newDocument = new Document();
            for (int j = 0; j < pagesPerFile && (i + j) < document.Pages.Count; j++)
                newDocument.Pages.Add(document.Pages[i + j + 1]);

            var safeFileName = SecurityHelper.SanitizeFileName($"{fileBaseName}_part_{++fileCount}.pdf");
            var outputPath = Path.Combine(outputDir, safeFileName);
            newDocument.Save(outputPath);
        }

        return await Task.FromResult($"PDF split into {fileCount} files in: {outputDir}");
    }

    private async Task<string> CompressDocument(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var outputPath = arguments?["outputPath"]?.GetValue<string>() ?? path;
        var compressImages = arguments?["compressImages"]?.GetValue<bool>() ?? true;
        var compressFonts = arguments?["compressFonts"]?.GetValue<bool>() ?? true;
        var removeUnusedObjects = arguments?["removeUnusedObjects"]?.GetValue<bool>() ?? true;

        SecurityHelper.ValidateFilePath(path, "path");
        SecurityHelper.ValidateFilePath(outputPath, "outputPath");

        var document = new Document(path);
        var optimizationOptions = new OptimizationOptions();

        if (compressImages)
        {
            optimizationOptions.ImageCompressionOptions.CompressImages = true;
            optimizationOptions.ImageCompressionOptions.ImageQuality = 75;
        }

        if (removeUnusedObjects)
        {
            optimizationOptions.LinkDuplcateStreams = true;
            optimizationOptions.RemoveUnusedObjects = true;
        }

        document.OptimizeResources(optimizationOptions);
        document.Save(outputPath);
        document.Dispose();

        var originalSize = new FileInfo(path).Length;
        var compressedSize = new FileInfo(outputPath).Length;
        var reduction = ((double)(originalSize - compressedSize) / originalSize) * 100;

        return await Task.FromResult($"PDF compressed. Size reduction: {reduction:F2}% ({originalSize} -> {compressedSize} bytes). Output: {outputPath}");
    }

    private async Task<string> EncryptDocument(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var outputPath = arguments?["outputPath"]?.GetValue<string>() ?? path;
        var userPassword = arguments?["userPassword"]?.GetValue<string>() ?? throw new ArgumentException("userPassword is required");
        var ownerPassword = arguments?["ownerPassword"]?.GetValue<string>() ?? throw new ArgumentException("ownerPassword is required");

        SecurityHelper.ValidateFilePath(path, "path");
        SecurityHelper.ValidateFilePath(outputPath, "outputPath");

        using var document = new Document(path);
        document.Encrypt(userPassword, ownerPassword, Permissions.PrintDocument | Permissions.ModifyContent, CryptoAlgorithm.AESx256);
        document.Save(outputPath);
        return await Task.FromResult($"PDF encrypted with password. Output: {outputPath}");
    }
}

