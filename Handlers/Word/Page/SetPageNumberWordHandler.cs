using Aspose.Words;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Word.Page;

/// <summary>
///     Handler for setting page number format and starting number in Word documents.
/// </summary>
public class SetPageNumberWordHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "set_page_number";

    /// <summary>
    ///     Sets page number format and starting number for a section.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Optional: pageNumberFormat (arabic, roman, letter)
    ///     Optional: startingPageNumber
    ///     Optional: sectionIndex
    /// </param>
    /// <returns>Success message with page number details.</returns>
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var pageNumberFormat = parameters.GetOptional<string?>("pageNumberFormat");
        var startingPageNumber = parameters.GetOptional<int?>("startingPageNumber");
        var sectionIndex = parameters.GetOptional<int?>("sectionIndex");

        var doc = context.Document;
        List<int> sectionsToUpdate;

        if (sectionIndex.HasValue)
        {
            if (sectionIndex.Value < 0 || sectionIndex.Value >= doc.Sections.Count)
                throw new ArgumentException($"sectionIndex must be between 0 and {doc.Sections.Count - 1}");
            sectionsToUpdate = [sectionIndex.Value];
        }
        else
        {
            sectionsToUpdate = [0];
        }

        foreach (var idx in sectionsToUpdate)
        {
            var pageSetup = doc.Sections[idx].PageSetup;

            if (!string.IsNullOrEmpty(pageNumberFormat))
            {
                var numStyle = pageNumberFormat.ToLower() switch
                {
                    "roman" => NumberStyle.UppercaseRoman,
                    "letter" => NumberStyle.UppercaseLetter,
                    _ => NumberStyle.Arabic
                };
                pageSetup.PageNumberStyle = numStyle;
            }

            if (startingPageNumber.HasValue)
            {
                pageSetup.RestartPageNumbering = true;
                pageSetup.PageStartingNumber = startingPageNumber.Value;
            }
        }

        MarkModified(context);

        return Success($"Page number settings updated for {sectionsToUpdate.Count} section(s)");
    }
}
