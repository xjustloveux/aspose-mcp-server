using System.Drawing;
using System.Drawing.Imaging;
using System.Text.Json.Nodes;
using Aspose.Slides;
using Aspose.Slides.Export;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools.PowerPoint;

/// <summary>
///     Unified tool for PowerPoint file operations (create, convert, merge, split)
///     Merges: PptCreateTool, PptConvertTool, PptMergeTool, PptSplitTool
/// </summary>
public class PptFileOperationsTool : IAsposeTool
{
    public string Description => @"PowerPoint file operations. Supports 4 operations: create, convert, merge, split.

Usage examples:
- Create presentation: ppt_file_operations(operation='create', path='new.pptx')
- Convert format: ppt_file_operations(operation='convert', inputPath='presentation.pptx', outputPath='presentation.pdf', format='pdf')
- Convert slide to image: ppt_file_operations(operation='convert', inputPath='presentation.pptx', outputPath='slide.png', format='png', slideIndex=0)
- Merge presentations: ppt_file_operations(operation='merge', inputPath='presentation1.pptx', outputPath='merged.pptx', inputPaths=['presentation2.pptx'])
- Split presentation: ppt_file_operations(operation='split', inputPath='presentation.pptx', outputDirectory='output/')";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            operation = new
            {
                type = "string",
                description = @"Operation to perform.
- 'create': Create a new presentation (required params: path)
- 'convert': Convert presentation format (required params: inputPath, outputPath, format)
- 'merge': Merge presentations (required params: outputPath, inputPaths)
- 'split': Split presentation (required params: inputPath, outputDirectory)",
                @enum = new[] { "create", "convert", "merge", "split" }
            },
            path = new
            {
                type = "string",
                description = "File path (output path for create operation, input path for convert/split operations)"
            },
            outputPath = new
            {
                type = "string",
                description = "Output file path (required for convert, optional for create, defaults to input path)"
            },
            inputPath = new
            {
                type = "string",
                description = "Input file path (required for convert/split)"
            },
            outputDirectory = new
            {
                type = "string",
                description = "Output directory path (required for split)"
            },
            format = new
            {
                type = "string",
                description = "Output format (pdf, html, pptx, jpg, png, etc., required for convert)"
            },
            slideIndex = new
            {
                type = "number",
                description = "Slide index to convert (0-based, optional for convert to image format, default: 0)"
            },
            inputPaths = new
            {
                type = "array",
                items = new { type = "string" },
                description = "Array of input presentation file paths (required for merge)"
            },
            keepSourceFormatting = new
            {
                type = "boolean",
                description = "Keep source formatting (optional, for merge, default: true)"
            },
            slidesPerFile = new
            {
                type = "number",
                description = "Number of slides per output file (optional, for split, default: 1)"
            },
            startSlideIndex = new
            {
                type = "number",
                description = "Start slide index (0-based, optional, for split)"
            },
            endSlideIndex = new
            {
                type = "number",
                description = "End slide index (0-based, optional, for split)"
            },
            outputFileNamePattern = new
            {
                type = "string",
                description =
                    "Output file name pattern, use {index} for slide number (optional, for split, default: 'slide_{index}.pptx')"
            }
        },
        required = new[] { "operation" }
    };

    /// <summary>
    ///     Executes the tool operation with the provided JSON arguments.
    /// </summary>
    /// <param name="arguments">JSON arguments object containing operation parameters.</param>
    /// <returns>Result message as a string.</returns>
    /// <exception cref="ArgumentException">Thrown when operation is unknown.</exception>
    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var operation = ArgumentHelper.GetString(arguments, "operation");

        return operation.ToLower() switch
        {
            "create" => await CreatePresentationAsync(arguments),
            "convert" => await ConvertPresentationAsync(arguments),
            "merge" => await MergePresentationsAsync(arguments),
            "split" => await SplitPresentationAsync(arguments),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Creates a new presentation.
    /// </summary>
    /// <param name="arguments">JSON arguments containing path or outputPath.</param>
    /// <returns>Success message with file path.</returns>
    private Task<string> CreatePresentationAsync(JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var path = ArgumentHelper.GetString(arguments, "path", "outputPath", "path or outputPath");
            SecurityHelper.ValidateFilePath(path, allowAbsolutePaths: true);

            using var presentation = new Presentation();
            presentation.Save(path, SaveFormat.Pptx);

            return $"PowerPoint presentation created successfully. Output: {path}";
        });
    }

    /// <summary>
    ///     Converts presentation to another format.
    /// </summary>
    /// <param name="arguments">JSON arguments containing inputPath, outputPath, format, and optional slideIndex.</param>
    /// <returns>Success message with output path.</returns>
    /// <exception cref="ArgumentException">Thrown when format is unsupported or slideIndex is out of range.</exception>
    private Task<string> ConvertPresentationAsync(JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var inputPath = ArgumentHelper.GetString(arguments, "inputPath", "path", "inputPath or path");
            SecurityHelper.ValidateFilePath(inputPath, "inputPath", true);
            var outputPath = ArgumentHelper.GetString(arguments, "outputPath");
            SecurityHelper.ValidateFilePath(outputPath, "outputPath", true);
            var format = ArgumentHelper.GetString(arguments, "format").ToLower();

            using var presentation = new Presentation(inputPath);

            if (format is "jpg" or "jpeg" or "png")
            {
                var slideIndex = ArgumentHelper.GetInt(arguments, "slideIndex", 0);
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
        });
    }

    /// <summary>
    ///     Merges multiple presentations into one.
    /// </summary>
    /// <param name="arguments">JSON arguments containing inputPaths array, outputPath, and optional keepSourceFormatting.</param>
    /// <returns>Success message with merged file path.</returns>
    /// <exception cref="ArgumentException">Thrown when no valid input paths are provided.</exception>
    private Task<string> MergePresentationsAsync(JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var outputPath = ArgumentHelper.GetString(arguments, "path", "outputPath", "path or outputPath");
            SecurityHelper.ValidateFilePath(outputPath, "outputPath", true);
            var inputPathsArray = ArgumentHelper.GetArray(arguments, "inputPaths");
            SecurityHelper.ValidateArraySize(inputPathsArray, "inputPaths");
            var keepSourceFormatting = ArgumentHelper.GetBool(arguments, "keepSourceFormatting");

            if (inputPathsArray.Count == 0) throw new ArgumentException("At least one input path is required");

            var inputPaths = inputPathsArray.Select(p => p?.GetValue<string>()).Where(p => !string.IsNullOrEmpty(p))
                .ToList();
            if (inputPaths.Count == 0) throw new ArgumentException("No valid input paths provided");

            using var masterPresentation = new Presentation(inputPaths[0]!);

            for (var i = 1; i < inputPaths.Count; i++)
            {
                var inputPath = inputPaths[i];
                if (string.IsNullOrEmpty(inputPath) || !File.Exists(inputPath)) continue;

                using var sourcePresentation = new Presentation(inputPath);
                foreach (var slide in sourcePresentation.Slides)
                    if (keepSourceFormatting)
                        masterPresentation.Slides.AddClone(slide);
                    else
                        masterPresentation.Slides.AddClone(slide, masterPresentation.Masters[0], true);
            }

            masterPresentation.Save(outputPath, SaveFormat.Pptx);
            return
                $"Merged {inputPaths.Count} presentations (Total slides: {masterPresentation.Slides.Count}). Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Splits presentation into multiple files.
    /// </summary>
    /// <param name="arguments">
    ///     JSON arguments containing inputPath, outputDirectory, and optional slidesPerFile,
    ///     startSlideIndex, endSlideIndex, outputFileNamePattern.
    /// </param>
    /// <returns>Success message with split file count.</returns>
    /// <exception cref="ArgumentException">Thrown when slide range is invalid.</exception>
    private Task<string> SplitPresentationAsync(JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var inputPath = ArgumentHelper.GetString(arguments, "inputPath", "path", "inputPath or path");
            var outputDirectory = ArgumentHelper.GetString(arguments, "outputDirectory");
            var slidesPerFile = ArgumentHelper.GetInt(arguments, "slidesPerFile", 1);
            var startSlideIndex = ArgumentHelper.GetIntNullable(arguments, "startSlideIndex");
            var endSlideIndex = ArgumentHelper.GetIntNullable(arguments, "endSlideIndex");
            var fileNamePattern = ArgumentHelper.GetString(arguments, "outputFileNamePattern", "slide_{index}.pptx");

            if (!Directory.Exists(outputDirectory)) Directory.CreateDirectory(outputDirectory);

            using var presentation = new Presentation(inputPath);
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
                var outputPath = Path.Combine(outputDirectory, outputFileName);
                newPresentation.Save(outputPath, SaveFormat.Pptx);
                fileCount++;
            }

            return $"Split presentation into {fileCount} file(s). Output: {outputDirectory}";
        });
    }
}