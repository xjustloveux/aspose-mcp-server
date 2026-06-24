using Aspose.Words;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers.Word;
using AsposeMcpServer.Results.Common;
using WordParagraph = Aspose.Words.Paragraph;
using WordTable = Aspose.Words.Tables.Table;

namespace AsposeMcpServer.Handlers.Word.Table;

/// <summary>
///     Handler for copying tables in Word documents.
/// </summary>
[ResultType(typeof(SuccessResult))]
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
    public override object Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var p = ExtractCopyWordTableParameters(parameters);

        var doc = context.Document;
        var sourceSectionIdx = p.SourceSectionIndex ?? 0;
        var targetSectionIdx = p.TargetSectionIndex ?? 0;

        if (sourceSectionIdx < 0 || sourceSectionIdx >= doc.Sections.Count)
            throw new ArgumentException($"sourceSectionIndex must be between 0 and {doc.Sections.Count - 1}");
        if (targetSectionIdx < 0 || targetSectionIdx >= doc.Sections.Count)
            throw new ArgumentException($"targetSectionIndex must be between 0 and {doc.Sections.Count - 1}");

        var sourceSection = doc.Sections[sourceSectionIdx];
        var sourceTables = sourceSection.Body.GetChildNodes(NodeType.Table, true).Cast<WordTable>().ToList();
        if (p.TableIndex < 0 || p.TableIndex >= sourceTables.Count)
            throw new ArgumentException($"sourceTableIndex must be between 0 and {sourceTables.Count - 1}");

        var sourceTable = sourceTables[p.TableIndex];
        var targetSection = doc.Sections[targetSectionIdx];

        var targetPara = ParagraphResolver
            .Resolve(doc, new ParagraphAddress(p.TargetParagraphIndex, StoryTypes.Body, targetSectionIdx)).Paragraph;

        var insertionPoint = FindInsertionPoint(targetSection, targetPara);

        if (insertionPoint == null)
            throw new ArgumentException(
                $"Unable to find valid insertion point (targetParagraphIndex: {p.TargetParagraphIndex})");

        var clonedTable = (WordTable)sourceTable.Clone(true);
        targetSection.Body.InsertAfter(clonedTable, insertionPoint);

        MarkModified(context);

        return new SuccessResult
            { Message = $"Successfully copied table {p.TableIndex} to paragraph {p.TargetParagraphIndex}." };
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

    private static CopyWordTableParameters ExtractCopyWordTableParameters(OperationParameters parameters)
    {
        var tableIndex = parameters.GetOptional("tableIndex", 0);
        var targetParagraphIndex = parameters.GetOptional("targetParagraphIndex", -1);
        var sourceSectionIndex = parameters.GetOptional<int?>("sourceSectionIndex");
        var targetSectionIndex = parameters.GetOptional<int?>("targetSectionIndex");

        return new CopyWordTableParameters(tableIndex, targetParagraphIndex, sourceSectionIndex, targetSectionIndex);
    }

    private sealed record CopyWordTableParameters(
        int TableIndex,
        int TargetParagraphIndex,
        int? SourceSectionIndex,
        int? TargetSectionIndex);
}
