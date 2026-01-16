using Aspose.Words;
using AsposeMcpServer.Core.Handlers;
using WordParagraph = Aspose.Words.Paragraph;

namespace AsposeMcpServer.Handlers.Word.Table;

/// <summary>
///     Handler for moving tables in Word documents.
/// </summary>
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
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var p = ExtractMoveWordTableParameters(parameters);

        var doc = context.Document;
        var sectionIdx = p.SectionIndex ?? 0;
        if (sectionIdx < 0 || sectionIdx >= doc.Sections.Count)
            throw new ArgumentException($"sectionIndex must be between 0 and {doc.Sections.Count - 1}");

        var section = doc.Sections[sectionIdx];
        var tables = section.Body.GetChildNodes(NodeType.Table, true).Cast<Aspose.Words.Tables.Table>().ToList();
        var paragraphs = section.Body.GetChildNodes(NodeType.Paragraph, false).Cast<WordParagraph>().ToList();

        if (p.TableIndex < 0 || p.TableIndex >= tables.Count)
            throw new ArgumentException($"tableIndex must be between 0 and {tables.Count - 1}");

        var table = tables[p.TableIndex];
        WordParagraph? targetPara;

        if (p.TargetParagraphIndex == -1)
        {
            if (paragraphs.Count > 0)
                targetPara = paragraphs[^1];
            else
                throw new ArgumentException(
                    "Cannot move table: section has no paragraphs. Use a valid paragraph index.");
        }
        else if (p.TargetParagraphIndex < 0 || p.TargetParagraphIndex >= paragraphs.Count)
        {
            throw new ArgumentException(
                $"targetParagraphIndex must be between 0 and {paragraphs.Count - 1}, or use -1 for document end");
        }
        else
        {
            targetPara = paragraphs[p.TargetParagraphIndex];
        }

        if (targetPara == null)
            throw new ArgumentException("Cannot find target paragraph");

        section.Body.InsertAfter(table, targetPara);

        MarkModified(context);

        return Success($"Successfully moved table {p.TableIndex}.");
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
