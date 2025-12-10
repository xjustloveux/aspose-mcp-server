using System.Text.Json.Nodes;
using Aspose.Pdf;
using Aspose.Pdf.Optimization;

namespace AsposeMcpServer.Tools;

public class PdfCompressTool : IAsposeTool
{
    public string Description => "Compress PDF document to reduce file size";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            path = new
            {
                type = "string",
                description = "PDF file path"
            },
            outputPath = new
            {
                type = "string",
                description = "Output file path (optional, if not provided overwrites input)"
            },
            compressImages = new
            {
                type = "boolean",
                description = "Compress images (optional, default: true)"
            },
            compressFonts = new
            {
                type = "boolean",
                description = "Compress fonts (optional, default: true)"
            },
            removeUnusedObjects = new
            {
                type = "boolean",
                description = "Remove unused objects (optional, default: true)"
            }
        },
        required = new[] { "path" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var outputPath = arguments?["outputPath"]?.GetValue<string>() ?? path;
        var compressImages = arguments?["compressImages"]?.GetValue<bool?>() ?? true;
        var compressFonts = arguments?["compressFonts"]?.GetValue<bool?>() ?? true;
        var removeUnusedObjects = arguments?["removeUnusedObjects"]?.GetValue<bool?>() ?? true;

        var document = new Document(path);
        var optimizationOptions = new OptimizationOptions();

        if (compressImages)
        {
            optimizationOptions.ImageCompressionOptions.CompressImages = true;
            optimizationOptions.ImageCompressionOptions.ImageQuality = 75;
        }

        // Note: FontOptimizationOptions may not be available in all versions
        // Font compression is handled automatically by OptimizeResources

        if (removeUnusedObjects)
        {
            optimizationOptions.LinkDuplcateStreams = true;
            optimizationOptions.RemoveUnusedObjects = true;
        }

        document.OptimizeResources(optimizationOptions);
        document.Save(outputPath);

        var originalSize = new FileInfo(path).Length;
        var compressedSize = new FileInfo(outputPath).Length;
        var reduction = ((double)(originalSize - compressedSize) / originalSize) * 100;

        return await Task.FromResult($"PDF compressed. Size reduction: {reduction:F2}% ({originalSize} -> {compressedSize} bytes): {outputPath}");
    }
}

