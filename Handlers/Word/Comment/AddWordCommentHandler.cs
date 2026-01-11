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
        List<WordParagraph> paragraphs = [];
        foreach (var section in doc.Sections.Cast<Section>())
        {
            var bodyParagraphs = section.Body.GetChildNodes(NodeType.Paragraph, false).Cast<WordParagraph>().ToList();
            paragraphs.AddRange(bodyParagraphs);
        }

        WordParagraph? targetPara;
        Run? startRun;
        Run? endRun;

        if (paragraphIndex.HasValue)
        {
            if (paragraphIndex.Value == -1)
            {
                if (paragraphs.Count > 0)
                    targetPara = paragraphs[^1];
                else
                    throw new InvalidOperationException("Document has no paragraphs to add comment to");
            }
            else if (paragraphIndex.Value < 0 || paragraphIndex.Value >= paragraphs.Count)
            {
                throw new ArgumentException(
                    $"Paragraph index {paragraphIndex.Value} is out of range (document has {paragraphs.Count} paragraphs)");
            }
            else
            {
                targetPara = paragraphs[paragraphIndex.Value];
            }
        }
        else
        {
            var builder = new DocumentBuilder(doc);
            builder.MoveToDocumentEnd();

            var newPara = new WordParagraph(doc);
            var newRun = new Run(doc, " ");
            newPara.AppendChild(newRun);
            doc.LastSection.Body.AppendChild(newPara);

            targetPara = newPara;
        }

        if (targetPara == null)
            throw new InvalidOperationException("Unable to determine target paragraph");

        var runs = targetPara.GetChildNodes(NodeType.Run, false);
        if (runs == null || runs.Count == 0)
        {
            var placeholderRun = new Run(doc, " ");
            targetPara.AppendChild(placeholderRun);
            startRun = placeholderRun;
            endRun = placeholderRun;
        }
        else if (startRunIndex.HasValue && endRunIndex.HasValue)
        {
            if (startRunIndex.Value < 0 || startRunIndex.Value >= runs.Count ||
                endRunIndex.Value < 0 || endRunIndex.Value >= runs.Count ||
                startRunIndex.Value > endRunIndex.Value)
                throw new ArgumentException($"Run index is out of range (paragraph has {runs.Count} Runs)");
            startRun = runs[startRunIndex.Value] as Run;
            endRun = runs[endRunIndex.Value] as Run;
        }
        else if (startRunIndex.HasValue)
        {
            if (startRunIndex.Value < 0 || startRunIndex.Value >= runs.Count)
                throw new ArgumentException($"Run index is out of range (paragraph has {runs.Count} Runs)");
            startRun = runs[startRunIndex.Value] as Run;
            endRun = startRun;
        }
        else
        {
            startRun = runs[0] as Run;
            endRun = runs[^1] as Run;
        }

        if (startRun == null || endRun == null)
            throw new InvalidOperationException("Unable to determine comment range");

        var para = startRun.ParentNode as WordParagraph ?? startRun.GetAncestor(NodeType.Paragraph) as WordParagraph;
        if (para == null)
            throw new InvalidOperationException("Unable to find paragraph node containing Run");

        var initial = authorInitial ?? (author.Length >= 2 ? author[..2].ToUpper() : author.ToUpper());
        var comment = new Aspose.Words.Comment(doc, author, initial, DateTime.UtcNow);
        comment.Paragraphs.Add(new WordParagraph(doc));
        comment.FirstParagraph.AppendChild(new Run(doc, text));

        var rangeStart = new CommentRangeStart(doc, comment.Id);
        var rangeEnd = new CommentRangeEnd(doc, comment.Id);

        var startPara = para;
        if (startRun.ParentNode != startPara)
            if (startRun.ParentNode is WordParagraph parentPara)
                startPara = parentPara;

        startPara.InsertBefore(rangeStart, startRun.ParentNode == startPara ? startRun : startPara.FirstChild);

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

        var rangeEndNode = endPara.GetChildNodes(NodeType.CommentRangeEnd, false)
            .Cast<CommentRangeEnd>()
            .FirstOrDefault(re => re.Id == comment.Id);

        if (rangeEndNode != null)
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
        else
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

        doc.EnsureMinimum();

        MarkModified(context);

        var result = "Comment added successfully\n";
        result += $"Author: {author}\n";
        result += $"Content: {text}";

        return result;
    }
}
