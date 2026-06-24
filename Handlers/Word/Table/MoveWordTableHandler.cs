using Aspose.Words;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers.Word;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.Word.Table;

/// <summary>
///     Handler for moving tables in Word documents.
/// </summary>
[ResultType(typeof(SuccessResult))]
public class MoveWordTableHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "move_table";

    /// <summary>
    ///     Moves a table to a different position.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Optional: tableIndex (default 0), targetParagraphIndex (default -1 for end), sectionIndex.
    /// </param>
    /// <returns>Success message.</returns>
    /// <exception cref="ArgumentException">Thrown when indices are out of range or target paragraph cannot be found.</exception>
    public override object Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var p = ExtractMoveWordTableParameters(parameters);

        var doc = context.Document;
        var sectionIdx = p.SectionIndex ?? 0;
        if (sectionIdx < 0 || sectionIdx >= doc.Sections.Count)
            throw new ArgumentException($"sectionIndex must be between 0 and {doc.Sections.Count - 1}");

        var section = doc.Sections[sectionIdx];
        var tables = section.Body.GetChildNodes(NodeType.Table, true).Cast<Aspose.Words.Tables.Table>().ToList();

        if (p.TableIndex < 0 || p.TableIndex >= tables.Count)
            throw new ArgumentException($"tableIndex must be between 0 and {tables.Count - 1}");

        var table = tables[p.TableIndex];

        var targetPara = ParagraphResolver
            .Resolve(doc, ParagraphAddress.From(parameters, p.TargetParagraphIndex)).Paragraph;

        // The resolver can address paragraphs nested in table cells, but a table is inserted relative
        // to a body-level block. Walk up to the direct child of the section body so the structural
        // insert has a valid sibling anchor instead of throwing on a cell paragraph.
        Node anchor = targetPara;
        while (anchor.ParentNode != null && anchor.ParentNode != section.Body)
            anchor = anchor.ParentNode;
        if (anchor.ParentNode != section.Body)
            throw new ArgumentException(
                $"targetParagraphIndex {p.TargetParagraphIndex} does not resolve to a paragraph within " +
                $"section {sectionIdx}'s body.");
        if (anchor == table)
            throw new ArgumentException(
                "Cannot move a table to a target paragraph inside the table being moved.");

        section.Body.InsertAfter(table, anchor);

        MarkModified(context);

        return new SuccessResult { Message = $"Successfully moved table {p.TableIndex}." };
    }

    private static MoveWordTableParameters ExtractMoveWordTableParameters(OperationParameters parameters)
    {
        var tableIndex = parameters.GetOptional("tableIndex", 0);
        var targetParagraphIndex = parameters.GetOptional("targetParagraphIndex", -1);
        var sectionIndex = parameters.GetOptional<int?>("sectionIndex");

        return new MoveWordTableParameters(tableIndex, targetParagraphIndex, sectionIndex);
    }

    private sealed record MoveWordTableParameters(int TableIndex, int TargetParagraphIndex, int? SectionIndex);
}
