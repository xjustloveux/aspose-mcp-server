using Aspose.Words;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Word.SectionBreak;

/// <summary>
///     Handler for deleting sections from Word documents.
/// </summary>
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
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var sectionIndex = parameters.GetOptional<int?>("sectionIndex");
        var sectionIndices = parameters.GetOptional<int[]?>("sectionIndices");

        var doc = context.Document;
        if (doc.Sections.Count <= 1)
            throw new ArgumentException("Cannot delete the last section. Document must have at least one section.");

        List<int> sectionsToDelete;
        if (sectionIndices is { Length: > 0 })
            sectionsToDelete = sectionIndices.OrderByDescending(s => s).ToList();
        else if (sectionIndex.HasValue)
            sectionsToDelete = [sectionIndex.Value];
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

        return Success(
            $"Deleted {deletedCount} section(s) with their content. Remaining sections: {doc.Sections.Count}.");
    }
}
