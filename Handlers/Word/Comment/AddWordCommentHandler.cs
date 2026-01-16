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
    /// <exception cref="ArgumentException">Thrown when text is not provided.</exception>
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var p = ExtractAddParameters(parameters);

        var doc = context.Document;
        var paragraphs = GetAllParagraphs(doc);
        var targetPara = GetTargetParagraph(doc, paragraphs, p.ParagraphIndex);
        var (startRun, endRun) = GetCommentRunRange(doc, targetPara, p.StartRunIndex, p.EndRunIndex);

        var para = GetContainingParagraph(startRun);
        var comment = CreateComment(doc, p.Text, p.Author, p.AuthorInitial);

        InsertCommentNodes(doc, comment, startRun, endRun, para);

        doc.EnsureMinimum();
        MarkModified(context);

        return $"Comment added successfully\nAuthor: {p.Author}\nContent: {p.Text}";
    }

    /// <summary>
    ///     Gets all paragraphs from all sections of the document.
    /// </summary>
    /// <param name="doc">The Word document.</param>
    /// <returns>A list of all paragraphs in the document.</returns>
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

    /// <summary>
    ///     Gets the target paragraph for the comment based on the paragraph index.
    /// </summary>
    /// <param name="doc">The Word document.</param>
    /// <param name="paragraphs">The list of all paragraphs.</param>
    /// <param name="paragraphIndex">The paragraph index, or null to create a new paragraph.</param>
    /// <returns>The target paragraph.</returns>
    /// <exception cref="InvalidOperationException">Thrown when document has no paragraphs.</exception>
    /// <exception cref="ArgumentException">Thrown when paragraph index is out of range.</exception>
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

    /// <summary>
    ///     Creates a new paragraph with a placeholder run.
    /// </summary>
    /// <param name="doc">The Word document.</param>
    /// <returns>The newly created paragraph.</returns>
    private static WordParagraph CreateNewParagraph(Document doc)
    {
        var newPara = new WordParagraph(doc);
        var newRun = new Run(doc, " ");
        newPara.AppendChild(newRun);
        doc.LastSection.Body.AppendChild(newPara);
        return newPara;
    }

    /// <summary>
    ///     Gets the run range for the comment.
    /// </summary>
    /// <param name="doc">The Word document.</param>
    /// <param name="targetPara">The target paragraph.</param>
    /// <param name="startRunIndex">The start run index.</param>
    /// <param name="endRunIndex">The end run index.</param>
    /// <returns>A tuple containing the start and end runs.</returns>
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

    /// <summary>
    ///     Gets the run range when both start and end indices are provided.
    /// </summary>
    /// <param name="runs">The collection of runs.</param>
    /// <param name="startIndex">The start run index.</param>
    /// <param name="endIndex">The end run index.</param>
    /// <returns>A tuple containing the start and end runs.</returns>
    /// <exception cref="ArgumentException">Thrown when run index is out of range.</exception>
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

    /// <summary>
    ///     Gets the run range when only start index is provided.
    /// </summary>
    /// <param name="runs">The collection of runs.</param>
    /// <param name="startIndex">The start run index.</param>
    /// <returns>A tuple containing the same run as both start and end.</returns>
    /// <exception cref="ArgumentException">Thrown when run index is out of range.</exception>
    private static (Run startRun, Run endRun) GetRunRangeWithStartIndex(NodeCollection runs, int startIndex)
    {
        if (startIndex < 0 || startIndex >= runs.Count)
            throw new ArgumentException($"Run index is out of range (paragraph has {runs.Count} Runs)");

        var startRun = runs[startIndex] as Run ??
                       throw new InvalidOperationException("Unable to determine comment range");
        return (startRun, startRun);
    }

    /// <summary>
    ///     Gets the paragraph containing the specified run.
    /// </summary>
    /// <param name="run">The run node.</param>
    /// <returns>The containing paragraph.</returns>
    /// <exception cref="InvalidOperationException">Thrown when paragraph cannot be found.</exception>
    private static WordParagraph GetContainingParagraph(Run run)
    {
        var para = run.ParentNode as WordParagraph ?? run.GetAncestor(NodeType.Paragraph) as WordParagraph;
        if (para == null)
            throw new InvalidOperationException("Unable to find paragraph node containing Run");
        return para;
    }

    /// <summary>
    ///     Creates a new comment with the specified text and author.
    /// </summary>
    /// <param name="doc">The Word document.</param>
    /// <param name="text">The comment text.</param>
    /// <param name="author">The author name.</param>
    /// <param name="authorInitial">The author initials.</param>
    /// <returns>The newly created comment.</returns>
    private static Aspose.Words.Comment CreateComment(Document doc, string text, string author, string? authorInitial)
    {
        var initial = authorInitial ?? (author.Length >= 2 ? author[..2].ToUpper() : author.ToUpper());
        var comment = new Aspose.Words.Comment(doc, author, initial, DateTime.UtcNow);
        comment.Paragraphs.Add(new WordParagraph(doc));
        comment.FirstParagraph.AppendChild(new Run(doc, text));
        return comment;
    }

    /// <summary>
    ///     Inserts the comment and its range markers into the document.
    /// </summary>
    /// <param name="doc">The Word document.</param>
    /// <param name="comment">The comment to insert.</param>
    /// <param name="startRun">The start run.</param>
    /// <param name="endRun">The end run.</param>
    /// <param name="para">The paragraph to insert into.</param>
    private static void InsertCommentNodes(Document doc, Aspose.Words.Comment comment, Run startRun, Run endRun,
        WordParagraph para)
    {
        var rangeStart = new CommentRangeStart(doc, comment.Id);
        var rangeEnd = new CommentRangeEnd(doc, comment.Id);

        InsertRangeStart(rangeStart, startRun, para);
        var endPara = InsertRangeEnd(rangeEnd, endRun);
        InsertComment(comment, endPara);
    }

    /// <summary>
    ///     Inserts the comment range start marker before the start run.
    /// </summary>
    /// <param name="rangeStart">The comment range start marker.</param>
    /// <param name="startRun">The start run.</param>
    /// <param name="para">The paragraph.</param>
    private static void InsertRangeStart(CommentRangeStart rangeStart, Run startRun, WordParagraph para)
    {
        var startPara = para;
        if (startRun.ParentNode != startPara && startRun.ParentNode is WordParagraph parentPara)
            startPara = parentPara;

        var insertBefore = startRun.ParentNode == startPara ? startRun : startPara.FirstChild;
        startPara.InsertBefore(rangeStart, insertBefore);
    }

    /// <summary>
    ///     Inserts the comment range end marker after the end run.
    /// </summary>
    /// <param name="rangeEnd">The comment range end marker.</param>
    /// <param name="endRun">The end run.</param>
    /// <returns>The paragraph containing the range end.</returns>
    /// <exception cref="InvalidOperationException">Thrown when paragraph cannot be found.</exception>
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

    /// <summary>
    ///     Inserts the comment node into the paragraph.
    /// </summary>
    /// <param name="comment">The comment to insert.</param>
    /// <param name="endPara">The paragraph to insert into.</param>
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

    /// <summary>
    ///     Inserts the comment after the range end marker.
    /// </summary>
    /// <param name="comment">The comment to insert.</param>
    /// <param name="endPara">The paragraph to insert into.</param>
    /// <param name="rangeEndNode">The range end marker.</param>
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

    /// <summary>
    ///     Inserts the comment at the end of the paragraph.
    /// </summary>
    /// <param name="comment">The comment to insert.</param>
    /// <param name="endPara">The paragraph to insert into.</param>
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

    /// <summary>
    ///     Extracts and validates parameters for the add comment operation.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted parameters.</returns>
    /// <exception cref="ArgumentException">Thrown when text is not provided.</exception>
    private static AddParameters ExtractAddParameters(OperationParameters parameters)
    {
        var text = parameters.GetOptional<string?>("text");
        var author = parameters.GetOptional("author", "Comment Author");
        var authorInitial = parameters.GetOptional<string?>("authorInitial");
        var paragraphIndex = parameters.GetOptional<int?>("paragraphIndex");
        var startRunIndex = parameters.GetOptional<int?>("startRunIndex");
        var endRunIndex = parameters.GetOptional<int?>("endRunIndex");

        if (string.IsNullOrEmpty(text))
            throw new ArgumentException("text is required for add operation");

        return new AddParameters(text, author, authorInitial, paragraphIndex, startRunIndex, endRunIndex);
    }

    /// <summary>
    ///     Parameters for the add comment operation.
    /// </summary>
    /// <param name="Text">The comment text content.</param>
    /// <param name="Author">The comment author name.</param>
    /// <param name="AuthorInitial">The author initials.</param>
    /// <param name="ParagraphIndex">The paragraph index to attach the comment to.</param>
    /// <param name="StartRunIndex">The start run index for the comment range.</param>
    /// <param name="EndRunIndex">The end run index for the comment range.</param>
    private sealed record AddParameters(
        string Text,
        string Author,
        string? AuthorInitial,
        int? ParagraphIndex,
        int? StartRunIndex,
        int? EndRunIndex);
}
