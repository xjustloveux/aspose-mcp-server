using Aspose.OCR;
using Aspose.OCR.Models.PreprocessingFilters;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers;
using AsposeMcpServer.Results.Ocr;

namespace AsposeMcpServer.Handlers.Ocr.Preprocessing;

/// <summary>
///     Base class for OCR preprocessing handlers providing shared preprocessing logic.
/// </summary>
public abstract class OcrPreprocessingHandlerBase : OperationHandlerBase<AsposeOcr>
{
    /// <summary>
    ///     Applies preprocessing filters to an input image and saves the result.
    /// </summary>
    /// <param name="inputPath">The input image file path.</param>
    /// <param name="outputPath">The output file path for the preprocessed image.</param>
    /// <param name="filters">The preprocessing filters to apply.</param>
    /// <exception cref="InvalidOperationException">Thrown when preprocessing produces no output files.</exception>
    protected static void SavePreprocessedImage(string inputPath, string outputPath,
        PreprocessingFilter filters)
    {
        var input = new OcrInput(InputType.SingleImage, filters);
        input.Add(inputPath);

        var tempDir = Path.Combine(Path.GetTempPath(), $"ocr_preprocess_{Guid.NewGuid()}");
        Directory.CreateDirectory(tempDir);
        try
        {
            ImageProcessing.Save(input, tempDir);

            var generatedFiles = Directory.GetFiles(tempDir);
            if (generatedFiles.Length == 0)
                throw new InvalidOperationException("Preprocessing produced no output files.");

            var outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                Directory.CreateDirectory(outputDir);

            File.Copy(generatedFiles[0], outputPath, true);
        }
        finally
        {
            try
            {
                Directory.Delete(tempDir, true);
            }
            catch
            {
                // Intentionally ignored: temp directory cleanup failure is non-critical
            }
        }
    }

    /// <summary>
    ///     Validates and extracts common preprocessing parameters from the operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted and validated preprocessing parameters.</returns>
    /// <exception cref="ArgumentException">Thrown when a required parameter is missing or a path is invalid.</exception>
    /// <exception cref="FileNotFoundException">Thrown when the input file does not exist.</exception>
    protected static PreprocessingParameters ExtractCommonParameters(OperationParameters parameters)
    {
        var path = parameters.GetRequired<string>("path");
        var outputPath = parameters.GetRequired<string>("outputPath");

        SecurityHelper.ValidateFilePath(path, "path", true);
        SecurityHelper.ValidateFilePath(outputPath, "outputPath", true);

        if (!File.Exists(path))
            throw new FileNotFoundException($"Input file not found: {path}");

        return new PreprocessingParameters(path, outputPath);
    }

    /// <summary>
    ///     Creates a standardized preprocessing result.
    /// </summary>
    /// <param name="preprocessingParams">The preprocessing parameters used.</param>
    /// <param name="operation">The preprocessing operation name.</param>
    /// <param name="description">A human-readable description of the operation performed.</param>
    /// <returns>An <see cref="OcrPreprocessingResult" /> with operation details and output file info.</returns>
    protected static OcrPreprocessingResult CreatePreprocessingResult(PreprocessingParameters preprocessingParams,
        string operation, string description)
    {
        return new OcrPreprocessingResult
        {
            SourcePath = preprocessingParams.Path,
            OutputPath = preprocessingParams.OutputPath,
            Operation = operation,
            FileSize = File.Exists(preprocessingParams.OutputPath)
                ? new FileInfo(preprocessingParams.OutputPath).Length
                : null,
            Message = $"{description}. Output saved to: {preprocessingParams.OutputPath}"
        };
    }

    /// <summary>
    ///     Common preprocessing parameters.
    /// </summary>
    /// <param name="Path">The input image file path.</param>
    /// <param name="OutputPath">The output file path for the preprocessed image.</param>
    protected sealed record PreprocessingParameters(string Path, string OutputPath);
}
