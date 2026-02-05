using System.Runtime.InteropServices;
using System.Text;
using Aspose.Drawing;
using Aspose.OCR;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers;
using AsposeMcpServer.Results.Ocr;

namespace AsposeMcpServer.Handlers.Ocr.Recognition;

/// <summary>
///     Handler for recognizing text from images and PDF files using OCR.
/// </summary>
[ResultType(typeof(OcrRecognitionResult))]
public class RecognizeHandler : OperationHandlerBase<AsposeOcr>
{
    /// <summary>
    ///     Supported image file extensions for OCR recognition.
    /// </summary>
    private static readonly HashSet<string> SupportedImageExtensions = new(StringComparer.OrdinalIgnoreCase)
    {
        ".png", ".jpg", ".jpeg", ".bmp", ".tiff", ".tif", ".gif"
    };

    /// <summary>
    ///     Mapping from common language names to Aspose.OCR Language enum values.
    /// </summary>
    private static readonly Dictionary<string, Language> CommonLanguageNames = new(StringComparer.OrdinalIgnoreCase)
    {
        { "English", Language.Eng },
        { "German", Language.Deu },
        { "Portuguese", Language.Por },
        { "Spanish", Language.Spa },
        { "French", Language.Fra },
        { "Italian", Language.Ita },
        { "Czech", Language.Cze },
        { "Danish", Language.Dan },
        { "Dutch", Language.Dum },
        { "Estonian", Language.Est },
        { "Finnish", Language.Fin },
        { "Latvian", Language.Lav },
        { "Lithuanian", Language.Lit },
        { "Norwegian", Language.Nor },
        { "Polish", Language.Pol },
        { "Romanian", Language.Rum },
        { "Slovak", Language.Slk },
        { "Slovenian", Language.Slv },
        { "Swedish", Language.Swe },
        { "Chinese", Language.Chi },
        { "Russian", Language.Rus },
        { "Ukrainian", Language.Ukr },
        { "Hindi", Language.Hin }
    };

    /// <inheritdoc />
    public override string Operation => "recognize";

    /// <summary>
    ///     Recognizes text from an image or PDF file using OCR.
    /// </summary>
    /// <param name="context">The OCR engine context.</param>
    /// <param name="parameters">
    ///     Required: path (image or PDF file path).
    ///     Optional: language (default: "English"), includeWords (default: false).
    /// </param>
    /// <returns>An <see cref="OcrRecognitionResult" /> containing recognized text and metadata.</returns>
    /// <exception cref="PlatformNotSupportedException">Thrown on unsupported platforms (Linux ARM64).</exception>
    /// <exception cref="FileNotFoundException">Thrown when the input file does not exist.</exception>
    /// <exception cref="ArgumentException">Thrown when the file format is not supported.</exception>
    public override object Execute(OperationContext<AsposeOcr> context, OperationParameters parameters)
    {
        ValidatePlatformSupport();

        var p = ExtractParameters(parameters);
        SecurityHelper.ValidateFilePath(p.Path, "path", true);

        if (!File.Exists(p.Path))
            throw new FileNotFoundException($"File not found: {p.Path}");

        var ocr = context.Document;
        var inputType = DetectInputType(p.Path);
        var input = new OcrInput(inputType);
        input.Add(p.Path);

        var settings = new RecognitionSettings
        {
            Language = ParseLanguage(p.Language)
        };

        var results = ocr.Recognize(input, settings);
        return BuildRecognitionResult(results, p.IncludeWords);
    }

    /// <summary>
    ///     Validates that the current platform supports OCR operations.
    ///     ONNX Runtime does not provide Linux ARM64 binaries.
    /// </summary>
    /// <exception cref="PlatformNotSupportedException">Thrown on Linux ARM64.</exception>
    internal static void ValidatePlatformSupport()
    {
        if (RuntimeInformation.IsOSPlatform(OSPlatform.Linux) &&
            RuntimeInformation.ProcessArchitecture == Architecture.Arm64)
            throw new PlatformNotSupportedException(
                "OCR is not supported on Linux ARM64. " +
                "ONNX Runtime does not provide Linux ARM64 binaries. " +
                "Please use x64 architecture.");
    }

    /// <summary>
    ///     Parses a language string to an Aspose.OCR Language enum value.
    ///     Accepts both abbreviated (Eng, Chi, Deu) and common names (English, Chinese, German).
    /// </summary>
    /// <param name="language">The language name or abbreviation.</param>
    /// <returns>The parsed Language enum value, defaults to Eng if parsing fails.</returns>
    internal static Language ParseLanguage(string language)
    {
        if (Enum.TryParse<Language>(language, true, out var result))
            return result;

        if (CommonLanguageNames.TryGetValue(language, out var mapped))
            return mapped;

        return Language.Eng;
    }

    /// <summary>
    ///     Detects the OCR input type based on file extension.
    /// </summary>
    /// <param name="path">The file path to detect.</param>
    /// <returns>The detected input type.</returns>
    /// <exception cref="ArgumentException">Thrown when the file format is not supported.</exception>
    private static InputType DetectInputType(string path)
    {
        var ext = Path.GetExtension(path);
        if (string.Equals(ext, ".pdf", StringComparison.OrdinalIgnoreCase))
            return InputType.PDF;

        if (SupportedImageExtensions.Contains(ext))
            return InputType.SingleImage;

        throw new ArgumentException(
            $"Unsupported file format: {ext}. Supported formats: PNG, JPG, BMP, TIFF, GIF, PDF.");
    }

    /// <summary>
    ///     Builds an <see cref="OcrRecognitionResult" /> from the OCR recognition results.
    /// </summary>
    /// <param name="results">The list of recognition results (one per page/image).</param>
    /// <param name="includeWords">Whether to include word-level details with bounding boxes.</param>
    /// <returns>The structured recognition result.</returns>
    internal static OcrRecognitionResult BuildRecognitionResult(List<RecognitionResult> results, bool includeWords)
    {
        var pages = new List<OcrPageResult>();
        var allText = new StringBuilder();

        for (var i = 0; i < results.Count; i++)
        {
            var result = results[i];
            var pageText = result.RecognitionText;
            allText.AppendLine(pageText);

            var page = new OcrPageResult
            {
                PageIndex = i,
                Text = pageText,
                Confidence = 0,
                Words = includeWords ? BuildWordInfos(result) : null
            };
            pages.Add(page);
        }

        return new OcrRecognitionResult
        {
            Text = allText.ToString().TrimEnd(),
            Confidence = 0,
            PageCount = results.Count,
            Pages = pages
        };
    }

    /// <summary>
    ///     Builds word-level information from recognition result areas.
    /// </summary>
    /// <param name="result">The recognition result for a single page.</param>
    /// <returns>A list of word info objects, or null if no area data is available.</returns>
    private static List<OcrWordInfo>? BuildWordInfos(RecognitionResult result)
    {
        if (result.RecognitionAreasText == null || result.RecognitionAreasRectangles == null)
            return null;

        var words = new List<OcrWordInfo>();
        var count = Math.Min(result.RecognitionAreasText.Count, result.RecognitionAreasRectangles.Count);

        for (var i = 0; i < count; i++)
        {
            var text = result.RecognitionAreasText[i];
            if (string.IsNullOrWhiteSpace(text))
                continue;

            var rect = result.RecognitionAreasRectangles[i];
            words.Add(new OcrWordInfo
            {
                Text = text,
                Confidence = 0,
                BoundingBox = RectangleToBoundingBox(rect)
            });
        }

        return words.Count > 0 ? words : null;
    }

    /// <summary>
    ///     Converts an <see cref="Aspose.Drawing.Rectangle" /> to an <see cref="OcrBoundingBox" />.
    /// </summary>
    /// <param name="rect">The rectangle to convert.</param>
    /// <returns>The equivalent bounding box.</returns>
    private static OcrBoundingBox RectangleToBoundingBox(Rectangle rect)
    {
        return new OcrBoundingBox
        {
            X = rect.X,
            Y = rect.Y,
            Width = rect.Width,
            Height = rect.Height
        };
    }

    /// <summary>
    ///     Extracts recognize parameters from the operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted parameters.</returns>
    private static RecognizeParameters ExtractParameters(OperationParameters parameters)
    {
        return new RecognizeParameters(
            parameters.GetRequired<string>("path"),
            parameters.GetOptional("language", "Eng"),
            parameters.GetOptional("includeWords", false));
    }

    /// <summary>
    ///     Parameters for the recognize operation.
    /// </summary>
    /// <param name="Path">The input file path.</param>
    /// <param name="Language">The recognition language.</param>
    /// <param name="IncludeWords">Whether to include word-level details.</param>
    private sealed record RecognizeParameters(string Path, string Language, bool IncludeWords);
}
