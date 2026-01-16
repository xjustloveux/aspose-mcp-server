using Aspose.Words;
using AsposeMcpServer.Core.Handlers;
using WordParagraph = Aspose.Words.Paragraph;
using WordTable = Aspose.Words.Tables.Table;

namespace AsposeMcpServer.Handlers.Word.Table;

/// <summary>
///     Handler for copying tables in Word documents.
/// </summary>
public class CopyWordTableHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "copy_table";

    /// <summary>
    ///     Copies a table to another location.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Optional: tableIndex (default 0), targetParagraphIndex (default -1 for end),
    ///     sourceSectionIndex, targetSectionIndex.
    /// </param>
    /// <returns>Success message.</returns>
    /// <exception cref="ArgumentException">Thrown when indices are out of range or target paragraph cannot be found.</exception>
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var tableIndex = parameters.GetOptional("tableIndex", 0);
        var targetParagraphIndex = parameters.GetOptional("targetParagraphIndex", -1);
        var sourceSectionIndex = parameters.GetOptional<int?>("sourceSectionIndex");
        var targetSectionIndex = parameters.GetOptional<int?>("targetSectionIndex");

        var doc = context.Document;
        var sourceSectionIdx = sourceSectionIndex ?? 0;
        var targetSectionIdx = targetSectionIndex ?? 0;

        if (sourceSectionIdx < 0 || sourceSectionIdx >= doc.Sections.Count)
            throw new ArgumentException($"sourceSectionIndex must be between 0 and {doc.Sections.Count - 1}");
        if (targetSectionIdx < 0 || targetSectionIdx >= doc.Sections.Count)
            throw new ArgumentException($"targetSectionIndex must be between 0 and {doc.Sections.Count - 1}");

        var sourceSection = doc.Sections[sourceSectionIdx];
        var sourceTables = sourceSection.Body.GetChildNodes(NodeType.Table, true).Cast<WordTable>().ToList();
        if (tableIndex < 0 || tableIndex >= sourceTables.Count)
            throw new ArgumentException($"sourceTableIndex must be between 0 and {sourceTables.Count - 1}");

        var sourceTable = sourceTables[tableIndex];
        var targetSection = doc.Sections[targetSectionIdx];
        var targetParagraphs =
            targetSection.Body.GetChildNodes(NodeType.Paragraph, true).Cast<WordParagraph>().ToList();

        WordParagraph? targetPara;
        if (targetParagraphIndex == -1)
        {
            if (targetParagraphs.Count > 0)
                targetPara = targetParagraphs[^1];
            else
                throw new ArgumentException(
                    "Cannot copy table: target section has no paragraphs. Use a valid paragraph index.");
        }
        else if (targetParagraphIndex < 0 || targetParagraphIndex >= targetParagraphs.Count)
        {
            throw new ArgumentException(
                $"targetParagraphIndex must be between 0 and {targetParagraphs.Count - 1}, or use -1 for document end");
        }
        else
        {
            targetPara = targetParagraphs[targetParagraphIndex];
        }

        if (targetPara == null)
            throw new ArgumentException("Cannot find target paragraph");

        var insertionPoint = FindInsertionPoint(targetSection, targetPara);

        if (insertionPoint == null)
            throw new ArgumentException(
                $"Unable to find valid insertion point (targetParagraphIndex: {targetParagraphIndex})");

        var clonedTable = (WordTable)sourceTable.Clone(true);
        targetSection.Body.InsertAfter(clonedTable, insertionPoint);

        MarkModified(context);

        return Success($"Successfully copied table {tableIndex} to paragraph {targetParagraphIndex}.");
    }

    /// <summary>
    ///     Finds the insertion point in the target section.
    /// </summary>
    /// <param name="targetSection">The target section.</param>
    /// <param name="targetPara">The target paragraph.</param>
    /// <returns>The insertion point node.</returns>
    private static Node? FindInsertionPoint(Section targetSection, WordParagraph targetPara)
    {
        if (targetPara.ParentNode == targetSection.Body)
            return targetPara;

        var bodyParagraphs = targetSection.Body.GetChildNodes(NodeType.Paragraph, false);
        var directPara = bodyParagraphs.Cast<WordParagraph>().FirstOrDefault(para => para == targetPara);

        if (directPara == null && bodyParagraphs.Count > 0)
            directPara = bodyParagraphs[^1] as WordParagraph;

        return directPara ?? targetSection.Body.LastChild;
    }
}
