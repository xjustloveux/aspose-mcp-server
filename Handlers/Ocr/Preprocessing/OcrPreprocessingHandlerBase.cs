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
    /// <param name="allowedBasePaths">
    ///     The allowlist of base paths used for symlink resolution before the write sink.
    ///     Pass an empty list to skip allowlist enforcement (allowlist disabled).
    /// </param>
    /// <exception cref="InvalidOperationException">Thrown when preprocessing produces no output files.</exception>
    /// <exception cref="ArgumentException">Thrown when outputPath resolves to a location outside the allowlist.</exception>
    protected static void SavePreprocessedImage(string inputPath, string outputPath,
        PreprocessingFilter filters, IReadOnlyList<string> allowedBasePaths)
    {
        // Resolved immediately before the read sink, the way every other sink in this repository
        // does it. The caller has already authorised this path; re-resolving here is what closes
        // the window between that check and this open (R19-OCR01).
        var resolvedInputPath =
            SecurityHelper.ResolveAndEnsureWithinAllowlist(inputPath, allowedBasePaths,
                nameof(inputPath));

        using var input = new OcrInput(InputType.SingleImage, filters);
        input.Add(resolvedInputPath);

        var tempDir = Path.Combine(Path.GetTempPath(), $"ocr_preprocess_{Guid.NewGuid()}");
        Directory.CreateDirectory(tempDir);
        try
        {
            ImageProcessing.Save(input, tempDir);

            var generatedFiles = Directory.GetFiles(tempDir);
            if (generatedFiles.Length == 0)
                throw new InvalidOperationException("Preprocessing produced no output files.");

            // H44: resolve symlinks immediately before the write sink (bug 20260415-symlink-toctou-sweep).
            var resolvedOutputPath =
                SecurityHelper.ResolveAndEnsureWithinAllowlist(outputPath, allowedBasePaths,
                    nameof(outputPath));

            // After the resolution, not before it. Creating the parent first meant a request that
            // was about to be refused had already made a directory wherever the caller named
            // (R19-OCR02).
            var outputDir = Path.GetDirectoryName(resolvedOutputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                Directory.CreateDirectory(outputDir);

            File.Copy(generatedFiles[0], resolvedOutputPath, true);
        }
        finally
        {
            try
            {
                SecurityHelper.SafeRecursiveDelete(tempDir, [], nameof(tempDir));
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
    /// <param name="allowedBasePaths">
    ///     The allowlist both paths must resolve inside. Empty means no allowlist is configured,
    ///     which the resolver already treats as unrestricted.
    /// </param>
    /// <returns>The extracted and validated preprocessing parameters, both paths resolved.</returns>
    /// <exception cref="ArgumentException">
    ///     Thrown when a required parameter is missing, a path is malformed, or a path resolves
    ///     outside the allowlist.
    /// </exception>
    /// <exception cref="FileNotFoundException">Thrown when the input file does not exist.</exception>
    /// <remarks>
    ///     Both paths are resolved here rather than only checked for shape. `ValidateFilePath`
    ///     says a string is well formed and `File.Exists` says something is there; neither says
    ///     the file is one this server may touch, and the input reached `OcrInput.Add` on the
    ///     caller's own spelling (R19-OCR01). Existence is asked of the resolved path, so the
    ///     answer is about the file that will actually be read.
    /// </remarks>
    protected static PreprocessingParameters ExtractCommonParameters(OperationParameters parameters,
        IReadOnlyList<string> allowedBasePaths)
    {
        var path = parameters.GetRequired<string>("path");
        var outputPath = parameters.GetRequired<string>("outputPath");

        SecurityHelper.ValidateFilePath(path, "path", true);
        SecurityHelper.ValidateFilePath(outputPath, "outputPath", true);

        var resolvedPath =
            SecurityHelper.ResolveAndEnsureWithinAllowlist(path, allowedBasePaths, "path");
        var resolvedOutputPath =
            SecurityHelper.ResolveAndEnsureWithinAllowlist(outputPath, allowedBasePaths,
                "outputPath");

        if (!File.Exists(resolvedPath))
            throw new FileNotFoundException("The specified file was not found.");

        return new PreprocessingParameters(resolvedPath, resolvedOutputPath);
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
