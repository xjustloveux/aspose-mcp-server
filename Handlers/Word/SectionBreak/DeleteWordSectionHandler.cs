using Aspose.Words;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.Word.SectionBreak;

/// <summary>
///     Handler for deleting sections from Word documents.
/// </summary>
[ResultType(typeof(SuccessResult))]
public class DeleteWordSectionHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "delete";

    /// <summary>
    ///     Deletes one or more sections from the document (including all content within).
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Optional: sectionIndex or sectionIndices (at least one required)
    /// </param>
    /// <returns>Success message with deletion details.</returns>
    public override object Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var p = ExtractDeleteWordSectionParameters(parameters);

        var doc = context.Document;
        if (doc.Sections.Count <= 1)
            throw new ArgumentException("Cannot delete the last section. Document must have at least one section.");

        List<int> sectionsToDelete;
        if (p.SectionIndices is { Length: > 0 })
            sectionsToDelete = p.SectionIndices.OrderByDescending(s => s).ToList();
        else if (p.SectionIndex.HasValue)
            sectionsToDelete = [p.SectionIndex.Value];
        else
            throw new ArgumentException(
                "Either sectionIndex or sectionIndices must be provided for delete operation");

        var deletedCount = 0;
        foreach (var idx in sectionsToDelete)
        {
            if (idx < 0 || idx >= doc.Sections.Count) continue;
            if (doc.Sections.Count <= 1) break;
            doc.Sections.RemoveAt(idx);
            deletedCount++;
        }

        MarkModified(context);

        return new SuccessResult
        {
            Message =
                $"Deleted {deletedCount} section(s) with their content. Remaining sections: {doc.Sections.Count}."
        };
    }

    private static DeleteWordSectionParameters ExtractDeleteWordSectionParameters(OperationParameters parameters)
    {
        return new DeleteWordSectionParameters(
            parameters.GetOptional<int?>("sectionIndex"),
            parameters.GetOptional<int[]?>("sectionIndices"));
    }

    private sealed record DeleteWordSectionParameters(
        int? SectionIndex,
        int[]? SectionIndices);
}
