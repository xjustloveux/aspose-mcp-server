using Aspose.Words;
using Aspose.Words.Settings;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;
using WordParagraph = Aspose.Words.Paragraph;

namespace AsposeMcpServer.Handlers.Word.File;

/// <summary>
///     Handler for creating new Word documents.
/// </summary>
public class CreateWordDocumentHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "create";

    /// <summary>
    ///     Creates a new Word document with specified settings.
    /// </summary>
    /// <param name="context">The operation context (Document may be null for this operation).</param>
    /// <param name="parameters">
    ///     Required: outputPath
    ///     Optional: content, skipInitialContent, marginTop, marginBottom, marginLeft, marginRight,
    ///     compatibilityMode, paperSize, pageWidth, pageHeight, headerDistance, footerDistance
    /// </param>
    /// <returns>Success message with output path.</returns>
    /// <exception cref="ArgumentException">Thrown when outputPath is missing.</exception>
    public override string
        Execute(OperationContext<Document> context,
            OperationParameters parameters)
    {
        var p = ExtractCreateParameters(parameters);

        if (string.IsNullOrEmpty(p.OutputPath))
            throw new ArgumentException("outputPath is required for create operation");

        SecurityHelper.ValidateFilePath(p.OutputPath, "outputPath", true);

        var outputDir = Path.GetDirectoryName(p.OutputPath);
        if (!string.IsNullOrEmpty(outputDir))
            Directory.CreateDirectory(outputDir);

        var doc = new Document();

        var wordVersion = p.CompatibilityMode switch
        {
            "Word2019" => MsWordVersion.Word2019,
            "Word2016" => MsWordVersion.Word2016,
            "Word2013" => MsWordVersion.Word2013,
            "Word2010" => MsWordVersion.Word2010,
            "Word2007" => MsWordVersion.Word2007,
            _ => MsWordVersion.Word2019
        };
        doc.CompatibilityOptions.OptimizeFor(wordVersion);

        var section = doc.FirstSection;
        if (section != null)
        {
            var pageSetup = section.PageSetup;

            if (!string.IsNullOrEmpty(p.PaperSize) && p.PageWidth == null && p.PageHeight == null)
            {
                pageSetup.PaperSize = p.PaperSize.ToUpper() switch
                {
                    "A4" => PaperSize.A4,
                    "LETTER" => PaperSize.Letter,
                    "A3" => PaperSize.A3,
                    "LEGAL" => PaperSize.Legal,
                    _ => PaperSize.A4
                };
            }
            else if (p.PageWidth != null || p.PageHeight != null)
            {
                pageSetup.PaperSize = PaperSize.Custom;
                pageSetup.PageWidth = p.PageWidth ?? 595.3;
                pageSetup.PageHeight = p.PageHeight ?? 841.9;
            }
            else
            {
                pageSetup.PaperSize = PaperSize.A4;
            }

            pageSetup.TopMargin = p.MarginTop;
            pageSetup.BottomMargin = p.MarginBottom;
            pageSetup.LeftMargin = p.MarginLeft;
            pageSetup.RightMargin = p.MarginRight;
            pageSetup.HeaderDistance = p.HeaderDistance;
            pageSetup.FooterDistance = p.FooterDistance;
        }

        var builder = new DocumentBuilder(doc);

        if (p.SkipInitialContent)
        {
            if (doc.FirstSection is { Body: not null })
            {
                doc.FirstSection.Body.RemoveAllChildren();
                var firstPara = new WordParagraph(doc)
                {
                    ParagraphFormat =
                    {
                        SpaceBefore = 0,
                        SpaceAfter = 0,
                        LineSpacing = 12
                    }
                };
                doc.FirstSection.Body.AppendChild(firstPara);
            }
        }
        else if (!string.IsNullOrEmpty(p.Content))
        {
            builder.Write(p.Content);
        }

        doc.Save(p.OutputPath);
        return $"Word document created successfully at: {p.OutputPath}";
    }

    private static CreateParameters ExtractCreateParameters(OperationParameters parameters)
    {
        return new CreateParameters(
            parameters.GetOptional<string?>("outputPath"),
            parameters.GetOptional<string?>("content"),
            parameters.GetOptional("skipInitialContent", false),
            parameters.GetOptional("marginTop", 70.87),
            parameters.GetOptional("marginBottom", 70.87),
            parameters.GetOptional("marginLeft", 70.87),
            parameters.GetOptional("marginRight", 70.87),
            parameters.GetOptional("compatibilityMode", "Word2019"),
            parameters.GetOptional("paperSize", "A4"),
            parameters.GetOptional<double?>("pageWidth"),
            parameters.GetOptional<double?>("pageHeight"),
            parameters.GetOptional("headerDistance", 35.4),
            parameters.GetOptional("footerDistance", 35.4));
    }

    private sealed record CreateParameters(
        string? OutputPath,
        string? Content,
        bool SkipInitialContent,
        double MarginTop,
        double MarginBottom,
        double MarginLeft,
        double MarginRight,
        string CompatibilityMode,
        string PaperSize,
        double? PageWidth,
        double? PageHeight,
        double HeaderDistance,
        double FooterDistance);
}
