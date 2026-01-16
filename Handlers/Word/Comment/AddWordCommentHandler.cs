using Aspose.Words;
using AsposeMcpServer.Core.Handlers;
using WordParagraph = Aspose.Words.Paragraph;

namespace AsposeMcpServer.Handlers.Word.Comment;

/// <summary>
///     Handler for adding comments to Word documents.
/// </summary>
public class AddWordCommentHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "add";

    /// <summary>
    ///     Adds a comment to the document at the specified paragraph and run range.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: text
    ///     Optional: author, authorInitial, paragraphIndex, startRunIndex, endRunIndex
    /// </param>
    /// <returns>Success message with comment details.</returns>
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var text = parameters.GetOptional<string?>("text");
        var author = parameters.GetOptional("author", "Comment Author");
        var authorInitial = parameters.GetOptional<string?>("authorInitial");
        var paragraphIndex = parameters.GetOptional<int?>("paragraphIndex");
        var startRunIndex = parameters.GetOptional<int?>("startRunIndex");
        var endRunIndex = parameters.GetOptional<int?>("endRunIndex");

        if (string.IsNullOrEmpty(text))
            throw new ArgumentException("text is required for add operation");

        var doc = context.Document;
        var paragraphs = GetAllParagraphs(doc);
        var targetPara = GetTargetParagraph(doc, paragraphs, paragraphIndex);
        var (startRun, endRun) = GetCommentRunRange(doc, targetPara, startRunIndex, endRunIndex);

        var para = GetContainingParagraph(startRun);
        var comment = CreateComment(doc, text, author, authorInitial);

        InsertCommentNodes(doc, comment, startRun, endRun, para);

        doc.EnsureMinimum();
        MarkModified(context);

        return $"Comment added successfully\nAuthor: {author}\nContent: {text}";
    }

    private static List<WordParagraph> GetAllParagraphs(Document doc)
    {
        List<WordParagraph> paragraphs = [];
        foreach (var section in doc.Sections.Cast<Section>())
        {
            var bodyParagraphs = section.Body.GetChildNodes(NodeType.Paragraph, false).Cast<WordParagraph>().ToList();
            paragraphs.AddRange(bodyParagraphs);
        }

        return paragraphs;
    }

    private static WordParagraph GetTargetParagraph(Document doc, List<WordParagraph> paragraphs, int? paragraphIndex)
    {
        if (!paragraphIndex.HasValue)
            return CreateNewParagraph(doc);

        if (paragraphIndex.Value == -1)
        {
            if (paragraphs.Count == 0)
                throw new InvalidOperationException("Document has no paragraphs to add comment to");
            return paragraphs[^1];
        }

        if (paragraphIndex.Value < 0 || paragraphIndex.Value >= paragraphs.Count)
            throw new ArgumentException(
                $"Paragraph index {paragraphIndex.Value} is out of range (document has {paragraphs.Count} paragraphs)");

        return paragraphs[paragraphIndex.Value];
    }

    private static WordParagraph CreateNewParagraph(Document doc)
    {
        var newPara = new WordParagraph(doc);
        var newRun = new Run(doc, " ");
        newPara.AppendChild(newRun);
        doc.LastSection.Body.AppendChild(newPara);
        return newPara;
    }

    private static (Run startRun, Run endRun) GetCommentRunRange(Document doc, WordParagraph targetPara,
        int? startRunIndex, int? endRunIndex)
    {
        var runs = targetPara.GetChildNodes(NodeType.Run, false);

        if (runs == null || runs.Count == 0)
        {
            var placeholderRun = new Run(doc, " ");
            targetPara.AppendChild(placeholderRun);
            return (placeholderRun, placeholderRun);
        }

        if (startRunIndex.HasValue && endRunIndex.HasValue)
            return GetRunRangeWithBothIndices(runs, startRunIndex.Value, endRunIndex.Value);

        if (startRunIndex.HasValue)
            return GetRunRangeWithStartIndex(runs, startRunIndex.Value);

        var start = runs[0] as Run ?? throw new InvalidOperationException("Unable to determine comment range");
        var end = runs[^1] as Run ?? throw new InvalidOperationException("Unable to determine comment range");
        return (start, end);
    }

    private static (Run startRun, Run endRun) GetRunRangeWithBothIndices(NodeCollection runs, int startIndex,
        int endIndex)
    {
        if (startIndex < 0 || startIndex >= runs.Count ||
            endIndex < 0 || endIndex >= runs.Count ||
            startIndex > endIndex)
            throw new ArgumentException($"Run index is out of range (paragraph has {runs.Count} Runs)");

        var startRun = runs[startIndex] as Run ??
                       throw new InvalidOperationException("Unable to determine comment range");
        var endRun = runs[endIndex] as Run ?? throw new InvalidOperationException("Unable to determine comment range");
        return (startRun, endRun);
    }

    private static (Run startRun, Run endRun) GetRunRangeWithStartIndex(NodeCollection runs, int startIndex)
    {
        if (startIndex < 0 || startIndex >= runs.Count)
            throw new ArgumentException($"Run index is out of range (paragraph has {runs.Count} Runs)");

        var startRun = runs[startIndex] as Run ??
                       throw new InvalidOperationException("Unable to determine comment range");
        return (startRun, startRun);
    }

    private static WordParagraph GetContainingParagraph(Run run)
    {
        var para = run.ParentNode as WordParagraph ?? run.GetAncestor(NodeType.Paragraph) as WordParagraph;
        if (para == null)
            throw new InvalidOperationException("Unable to find paragraph node containing Run");
        return para;
    }

    private static Aspose.Words.Comment CreateComment(Document doc, string text, string author, string? authorInitial)
    {
        var initial = authorInitial ?? (author.Length >= 2 ? author[..2].ToUpper() : author.ToUpper());
        var comment = new Aspose.Words.Comment(doc, author, initial, DateTime.UtcNow);
        comment.Paragraphs.Add(new WordParagraph(doc));
        comment.FirstParagraph.AppendChild(new Run(doc, text));
        return comment;
    }

    private static void InsertCommentNodes(Document doc, Aspose.Words.Comment comment, Run startRun, Run endRun,
        WordParagraph para)
    {
        var rangeStart = new CommentRangeStart(doc, comment.Id);
        var rangeEnd = new CommentRangeEnd(doc, comment.Id);

        InsertRangeStart(rangeStart, startRun, para);
        var endPara = InsertRangeEnd(rangeEnd, endRun);
        InsertComment(comment, endPara);
    }

    private static void InsertRangeStart(CommentRangeStart rangeStart, Run startRun, WordParagraph para)
    {
        var startPara = para;
        if (startRun.ParentNode != startPara && startRun.ParentNode is WordParagraph parentPara)
            startPara = parentPara;

        var insertBefore = startRun.ParentNode == startPara ? startRun : startPara.FirstChild;
        startPara.InsertBefore(rangeStart, insertBefore);
    }

    private static WordParagraph InsertRangeEnd(CommentRangeEnd rangeEnd, Run endRun)
    {
        var endPara = endRun.ParentNode as WordParagraph ?? endRun.GetAncestor(NodeType.Paragraph) as WordParagraph;
        if (endPara == null)
            throw new InvalidOperationException("Unable to find paragraph containing endRun");

        if (endRun.ParentNode == endPara)
        {
            var nextSibling = endRun.NextSibling;
            if (nextSibling != null)
                endPara.InsertBefore(rangeEnd, nextSibling);
            else
                endPara.AppendChild(rangeEnd);
        }
        else
        {
            endPara.AppendChild(rangeEnd);
        }

        return endPara;
    }

    private static void InsertComment(Aspose.Words.Comment comment, WordParagraph endPara)
    {
        var rangeEndNode = endPara.GetChildNodes(NodeType.CommentRangeEnd, false)
            .Cast<CommentRangeEnd>()
            .FirstOrDefault(re => re.Id == comment.Id);

        if (rangeEndNode != null)
            InsertCommentAfterRangeEnd(comment, endPara, rangeEndNode);
        else
            InsertCommentAtEnd(comment, endPara);
    }

    private static void InsertCommentAfterRangeEnd(Aspose.Words.Comment comment, WordParagraph endPara,
        CommentRangeEnd rangeEndNode)
    {
        if (comment.ParentNode == null)
        {
            endPara.InsertAfter(comment, rangeEndNode);
        }
        else if (comment.ParentNode != endPara)
        {
            comment.Remove();
            endPara.InsertAfter(comment, rangeEndNode);
        }
    }

    private static void InsertCommentAtEnd(Aspose.Words.Comment comment, WordParagraph endPara)
    {
        if (comment.ParentNode == null)
        {
            endPara.AppendChild(comment);
        }
        else if (comment.ParentNode != endPara)
        {
            comment.Remove();
            endPara.AppendChild(comment);
        }
    }
}
