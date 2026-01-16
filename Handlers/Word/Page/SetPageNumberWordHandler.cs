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
        var setParams = ExtractSetPageNumberParameters(parameters);

        var doc = context.Document;
        List<int> sectionsToUpdate;

        if (setParams.SectionIndex.HasValue)
        {
            if (setParams.SectionIndex.Value < 0 || setParams.SectionIndex.Value >= doc.Sections.Count)
                throw new ArgumentException($"sectionIndex must be between 0 and {doc.Sections.Count - 1}");
            sectionsToUpdate = [setParams.SectionIndex.Value];
        }
        else
        {
            sectionsToUpdate = [0];
        }

        foreach (var idx in sectionsToUpdate)
        {
            var pageSetup = doc.Sections[idx].PageSetup;

            if (!string.IsNullOrEmpty(setParams.PageNumberFormat))
            {
                var numStyle = setParams.PageNumberFormat.ToLower() switch
                {
                    "roman" => NumberStyle.UppercaseRoman,
                    "letter" => NumberStyle.UppercaseLetter,
                    _ => NumberStyle.Arabic
                };
                pageSetup.PageNumberStyle = numStyle;
            }

            if (setParams.StartingPageNumber.HasValue)
            {
                pageSetup.RestartPageNumbering = true;
                pageSetup.PageStartingNumber = setParams.StartingPageNumber.Value;
            }
        }

        MarkModified(context);

        return Success($"Page number settings updated for {sectionsToUpdate.Count} section(s)");
    }

    /// <summary>
    ///     Extracts set page number parameters from operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted set page number parameters.</returns>
    private static SetPageNumberParameters ExtractSetPageNumberParameters(OperationParameters parameters)
    {
        return new SetPageNumberParameters(
            parameters.GetOptional<string?>("pageNumberFormat"),
            parameters.GetOptional<int?>("startingPageNumber"),
            parameters.GetOptional<int?>("sectionIndex")
        );
    }

    /// <summary>
    ///     Record to hold set page number parameters.
    /// </summary>
    /// <param name="PageNumberFormat">The page number format (arabic, roman, letter).</param>
    /// <param name="StartingPageNumber">The starting page number.</param>
    /// <param name="SectionIndex">The section index to apply settings to.</param>
    private sealed record SetPageNumberParameters(string? PageNumberFormat, int? StartingPageNumber, int? SectionIndex);
}
