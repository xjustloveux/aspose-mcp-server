using System.ComponentModel;
using System.Drawing;
using System.Drawing.Imaging;
using Aspose.Slides;
using Aspose.Slides.Export;
using AsposeMcpServer.Core.Helpers;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.PowerPoint;

/// <summary>
///     Unified tool for PowerPoint file operations (create, convert, merge, split)
///     Merges: PptCreateTool, PptConvertTool, PptMergeTool, PptSplitTool
/// </summary>
[McpServerToolType]
public class PptFileOperationsTool
{
    [McpServerTool(Name = "ppt_file_operations")]
    [Description(@"PowerPoint file operations. Supports 4 operations: create, convert, merge, split.

Usage examples:
- Create presentation: ppt_file_operations(operation='create', path='new.pptx')
- Convert format: ppt_file_operations(operation='convert', inputPath='presentation.pptx', outputPath='presentation.pdf', format='pdf')
- Convert slide to image: ppt_file_operations(operation='convert', inputPath='presentation.pptx', outputPath='slide.png', format='png', slideIndex=0)
- Merge presentations: ppt_file_operations(operation='merge', inputPath='presentation1.pptx', outputPath='merged.pptx', inputPaths=['presentation2.pptx'])
- Split presentation: ppt_file_operations(operation='split', inputPath='presentation.pptx', outputDirectory='output/')")]
    public string Execute(
        [Description("Operation: create, convert, merge, split")]
        string operation,
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
        return operation.ToLower() switch
        {
            "create" => CreatePresentation(path, outputPath),
            "convert" => ConvertPresentation(inputPath, path, outputPath, format, slideIndex),
            "merge" => MergePresentations(path, outputPath, inputPaths, keepSourceFormatting),
            "split" => SplitPresentation(inputPath, path, outputDirectory, slidesPerFile, startSlideIndex,
                endSlideIndex, outputFileNamePattern),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Creates a new presentation.
    /// </summary>
    /// <param name="path">The file path to save the presentation.</param>
    /// <param name="outputPath">Alternative output file path (used if path is null).</param>
    /// <returns>A message indicating the presentation was created successfully.</returns>
    /// <exception cref="ArgumentException">Thrown when neither path nor outputPath is provided.</exception>
    private static string CreatePresentation(string? path, string? outputPath)
    {
        var savePath = path ?? outputPath;
        if (string.IsNullOrEmpty(savePath))
            throw new ArgumentException("path or outputPath is required for create operation");

        SecurityHelper.ValidateFilePath(savePath, allowAbsolutePaths: true);

        using var presentation = new Presentation();
        presentation.Save(savePath, SaveFormat.Pptx);

        return $"PowerPoint presentation created successfully. Output: {savePath}";
    }

    /// <summary>
    ///     Converts presentation to another format.
    /// </summary>
    /// <param name="inputPath">The input presentation file path.</param>
    /// <param name="path">Alternative input file path (used if inputPath is null).</param>
    /// <param name="outputPath">The output file path for the converted presentation.</param>
    /// <param name="format">The target format (pdf, html, pptx, jpg, png, etc.).</param>
    /// <param name="slideIndex">The zero-based slide index to convert (for image formats).</param>
    /// <returns>A message indicating the conversion was successful.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or format is unsupported.</exception>
    private static string ConvertPresentation(string? inputPath, string? path, string? outputPath, string? format,
        int slideIndex)
    {
        var sourcePath = inputPath ?? path;
        if (string.IsNullOrEmpty(sourcePath))
            throw new ArgumentException("inputPath or path is required for convert operation");
        if (string.IsNullOrEmpty(outputPath))
            throw new ArgumentException("outputPath is required for convert operation");
        if (string.IsNullOrEmpty(format))
            throw new ArgumentException("format is required for convert operation");

        SecurityHelper.ValidateFilePath(sourcePath, "inputPath", true);
        SecurityHelper.ValidateFilePath(outputPath, "outputPath", true);

        format = format.ToLower();

        using var presentation = new Presentation(sourcePath);

        if (format is "jpg" or "jpeg" or "png")
        {
            var slide = PowerPointHelper.GetSlide(presentation, slideIndex);

            var slideSize = presentation.SlideSize.Size;
            var targetSize = new Size((int)slideSize.Width, (int)slideSize.Height);

#pragma warning disable CA1416 // Validate platform compatibility
            using var bitmap = slide.GetThumbnail(targetSize);
            var imageFormat = format == "png" ? ImageFormat.Png : ImageFormat.Jpeg;
            bitmap.Save(outputPath, imageFormat);
#pragma warning restore CA1416

            var formatName = format == "png" ? "PNG" : "JPEG";
            return $"Slide {slideIndex} converted to {formatName}. Output: {outputPath}";
        }

        var saveFormat = format switch
        {
            "pdf" => SaveFormat.Pdf,
            "html" => SaveFormat.Html,
            "pptx" => SaveFormat.Pptx,
            "ppt" => SaveFormat.Ppt,
            "odp" => SaveFormat.Odp,
            "xps" => SaveFormat.Xps,
            "tiff" => SaveFormat.Tiff,
            _ => throw new ArgumentException($"Unsupported format: {format}")
        };

        presentation.Save(outputPath, saveFormat);

        return $"Presentation converted to {format.ToUpper()} format. Output: {outputPath}";
    }

    /// <summary>
    ///     Merges multiple presentations into one.
    /// </summary>
    /// <param name="path">The output file path to save the merged presentation.</param>
    /// <param name="outputPath">Alternative output file path (used if path is null).</param>
    /// <param name="inputPaths">Array of input presentation file paths to merge.</param>
    /// <param name="keepSourceFormatting">Whether to keep source formatting when merging slides.</param>
    /// <returns>A message indicating the merge was successful with file count information.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or no valid input paths are provided.</exception>
    private static string MergePresentations(string? path, string? outputPath, string[]? inputPaths,
        bool keepSourceFormatting)
    {
        var savePath = path ?? outputPath;
        if (string.IsNullOrEmpty(savePath))
            throw new ArgumentException("path or outputPath is required for merge operation");
        if (inputPaths == null || inputPaths.Length == 0)
            throw new ArgumentException("inputPaths is required for merge operation");

        SecurityHelper.ValidateFilePath(savePath, "outputPath", true);

        var validPaths = inputPaths.Where(p => !string.IsNullOrEmpty(p)).ToList();
        if (validPaths.Count == 0)
            throw new ArgumentException("No valid input paths provided");

        using var masterPresentation = new Presentation(validPaths[0]);

        for (var i = 1; i < validPaths.Count; i++)
        {
            var inputPath = validPaths[i];
            if (string.IsNullOrEmpty(inputPath) || !File.Exists(inputPath)) continue;

            using var sourcePresentation = new Presentation(inputPath);
            foreach (var slide in sourcePresentation.Slides)
                if (keepSourceFormatting)
                    masterPresentation.Slides.AddClone(slide);
                else
                    masterPresentation.Slides.AddClone(slide, masterPresentation.Masters[0], true);
        }

        masterPresentation.Save(savePath, SaveFormat.Pptx);
        return
            $"Merged {validPaths.Count} presentations (Total slides: {masterPresentation.Slides.Count}). Output: {savePath}";
    }

    /// <summary>
    ///     Splits presentation into multiple files.
    /// </summary>
    /// <param name="inputPath">The input presentation file path.</param>
    /// <param name="path">Alternative input file path (used if inputPath is null).</param>
    /// <param name="outputDirectory">The output directory path for split files.</param>
    /// <param name="slidesPerFile">Number of slides per output file.</param>
    /// <param name="startSlideIndex">The zero-based start slide index (null for beginning).</param>
    /// <param name="endSlideIndex">The zero-based end slide index (null for end).</param>
    /// <param name="fileNamePattern">Output file name pattern with {index} placeholder.</param>
    /// <returns>A message indicating the split was successful with file count information.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or slide range is invalid.</exception>
    private static string SplitPresentation(string? inputPath, string? path, string? outputDirectory,
        int slidesPerFile, int? startSlideIndex, int? endSlideIndex, string fileNamePattern)
    {
        var sourcePath = inputPath ?? path;
        if (string.IsNullOrEmpty(sourcePath))
            throw new ArgumentException("inputPath or path is required for split operation");
        if (string.IsNullOrEmpty(outputDirectory))
            throw new ArgumentException("outputDirectory is required for split operation");

        if (!Directory.Exists(outputDirectory))
            Directory.CreateDirectory(outputDirectory);

        using var presentation = new Presentation(sourcePath);
        var totalSlides = presentation.Slides.Count;

        var start = startSlideIndex ?? 0;
        var end = endSlideIndex ?? totalSlides - 1;

        if (start < 0 || start >= totalSlides || end < 0 || end >= totalSlides || start > end)
            throw new ArgumentException($"Invalid slide range: start={start}, end={end}, total={totalSlides}");

        var fileCount = 0;
        for (var i = start; i <= end; i += slidesPerFile)
        {
            using var newPresentation = new Presentation();
            newPresentation.Slides.RemoveAt(0);

            for (var j = 0; j < slidesPerFile && i + j <= end; j++)
                newPresentation.Slides.AddClone(presentation.Slides[i + j]);

            var outputFileName = fileNamePattern.Replace("{index}", fileCount.ToString());
            outputFileName = SecurityHelper.SanitizeFileName(outputFileName);
            var outPath = Path.Combine(outputDirectory, outputFileName);
            newPresentation.Save(outPath, SaveFormat.Pptx);
            fileCount++;
        }

        return $"Split presentation into {fileCount} file(s). Output: {outputDirectory}";
    }
}