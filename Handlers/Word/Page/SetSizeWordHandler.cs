using System.Text.Json.Nodes;
using Aspose.Words;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers.Word;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.Word.Page;

/// <summary>
///     Handler for setting page size in Word documents.
/// </summary>
[ResultType(typeof(SuccessResult))]
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
    public override object Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var setParams = ExtractSetSizeParameters(parameters);

        var doc = context.Document;
        var sectionsToUpdate = WordPageHelper.GetTargetSections(doc, setParams.SectionIndex, setParams.SectionIndices);

        foreach (var idx in sectionsToUpdate)
        {
            var pageSetup = doc.Sections[idx].PageSetup;

            if (!string.IsNullOrEmpty(setParams.PaperSize))
            {
                pageSetup.PaperSize = setParams.PaperSize.ToUpper() switch
                {
                    "A4" => PaperSize.A4,
                    "LETTER" => PaperSize.Letter,
                    "LEGAL" => PaperSize.Legal,
                    "A3" => PaperSize.A3,
                    "A5" => PaperSize.A5,
                    _ => PaperSize.A4
                };
            }
            else if (setParams is { Width: { } width, Height: { } height })
            {
                pageSetup.PageWidth = width;
                pageSetup.PageHeight = height;
            }
            else
            {
                throw new ArgumentException("Either paperSize or both width and height must be provided");
            }
        }

        MarkModified(context);

        return new SuccessResult { Message = $"Page size updated for {sectionsToUpdate.Count} section(s)" };
    }

    /// <summary>
    ///     Extracts set size parameters from operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted set size parameters.</returns>
    private static SetSizeParameters ExtractSetSizeParameters(OperationParameters parameters)
    {
        return new SetSizeParameters(
            parameters.GetOptional<double?>("width"),
            parameters.GetOptional<double?>("height"),
            parameters.GetOptional<string?>("paperSize"),
            parameters.GetOptional<int?>("sectionIndex"),
            parameters.GetOptional<JsonArray?>("sectionIndices")
        );
    }

    /// <summary>
    ///     Record to hold set size parameters.
    /// </summary>
    /// <param name="Width">The page width in points.</param>
    /// <param name="Height">The page height in points.</param>
    /// <param name="PaperSize">The predefined paper size (A4, Letter, Legal, A3, A5).</param>
    /// <param name="SectionIndex">The section index to apply size to.</param>
    /// <param name="SectionIndices">The array of section indices to apply size to.</param>
    private sealed record SetSizeParameters(
        double? Width,
        double? Height,
        string? PaperSize,
        int? SectionIndex,
        JsonArray? SectionIndices);
}
