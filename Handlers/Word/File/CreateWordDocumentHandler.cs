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
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var outputPath = parameters.GetOptional<string?>("outputPath");
        var content = parameters.GetOptional<string?>("content");
        var skipInitialContent = parameters.GetOptional("skipInitialContent", false);
        var marginTop = parameters.GetOptional("marginTop", 70.87);
        var marginBottom = parameters.GetOptional("marginBottom", 70.87);
        var marginLeft = parameters.GetOptional("marginLeft", 70.87);
        var marginRight = parameters.GetOptional("marginRight", 70.87);
        var compatibilityMode = parameters.GetOptional("compatibilityMode", "Word2019");
        var paperSize = parameters.GetOptional("paperSize", "A4");
        var pageWidth = parameters.GetOptional<double?>("pageWidth");
        var pageHeight = parameters.GetOptional<double?>("pageHeight");
        var headerDistance = parameters.GetOptional("headerDistance", 35.4);
        var footerDistance = parameters.GetOptional("footerDistance", 35.4);

        if (string.IsNullOrEmpty(outputPath))
            throw new ArgumentException("outputPath is required for create operation");

        SecurityHelper.ValidateFilePath(outputPath, "outputPath", true);

        var outputDir = Path.GetDirectoryName(outputPath);
        if (!string.IsNullOrEmpty(outputDir))
            Directory.CreateDirectory(outputDir);

        var doc = new Document();

        var wordVersion = compatibilityMode switch
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

            if (!string.IsNullOrEmpty(paperSize) && pageWidth == null && pageHeight == null)
            {
                pageSetup.PaperSize = paperSize.ToUpper() switch
                {
                    "A4" => PaperSize.A4,
                    "LETTER" => PaperSize.Letter,
                    "A3" => PaperSize.A3,
                    "LEGAL" => PaperSize.Legal,
                    _ => PaperSize.A4
                };
            }
            else if (pageWidth != null || pageHeight != null)
            {
                pageSetup.PaperSize = PaperSize.Custom;
                pageSetup.PageWidth = pageWidth ?? 595.3;
                pageSetup.PageHeight = pageHeight ?? 841.9;
            }
            else
            {
                pageSetup.PaperSize = PaperSize.A4;
            }

            pageSetup.TopMargin = marginTop;
            pageSetup.BottomMargin = marginBottom;
            pageSetup.LeftMargin = marginLeft;
            pageSetup.RightMargin = marginRight;
            pageSetup.HeaderDistance = headerDistance;
            pageSetup.FooterDistance = footerDistance;
        }

        var builder = new DocumentBuilder(doc);

        if (skipInitialContent)
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
        else if (!string.IsNullOrEmpty(content))
        {
            builder.Write(content);
        }

        doc.Save(outputPath);
        return $"Word document created successfully at: {outputPath}";
    }
}
