using System.Text.Json.Nodes;
using Aspose.Words;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Word.Page;

/// <summary>
///     Handler for setting page size in Word documents.
/// </summary>
public class SetSizeWordHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "set_size";

    /// <summary>
    ///     Sets page size using custom dimensions or predefined paper sizes.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Optional: width, height (page dimensions in points)
    ///     Optional: paperSize (A4, Letter, Legal, A3, A5)
    ///     Optional: sectionIndex, sectionIndices
    /// </param>
    /// <returns>Success message with size details.</returns>
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var width = parameters.GetOptional<double?>("width");
        var height = parameters.GetOptional<double?>("height");
        var paperSize = parameters.GetOptional<string?>("paperSize");
        var sectionIndex = parameters.GetOptional<int?>("sectionIndex");
        var sectionIndices = parameters.GetOptional<JsonArray?>("sectionIndices");

        var doc = context.Document;
        var sectionsToUpdate = WordPageHelper.GetTargetSections(doc, sectionIndex, sectionIndices);

        foreach (var idx in sectionsToUpdate)
        {
            var pageSetup = doc.Sections[idx].PageSetup;

            if (!string.IsNullOrEmpty(paperSize))
            {
                pageSetup.PaperSize = paperSize.ToUpper() switch
                {
                    "A4" => PaperSize.A4,
                    "LETTER" => PaperSize.Letter,
                    "LEGAL" => PaperSize.Legal,
                    "A3" => PaperSize.A3,
                    "A5" => PaperSize.A5,
                    _ => PaperSize.A4
                };
            }
            else if (width.HasValue && height.HasValue)
            {
                pageSetup.PageWidth = width.Value;
                pageSetup.PageHeight = height.Value;
            }
            else
            {
                throw new ArgumentException("Either paperSize or both width and height must be provided");
            }
        }

        MarkModified(context);

        return Success($"Page size updated for {sectionsToUpdate.Count} section(s)");
    }
}
