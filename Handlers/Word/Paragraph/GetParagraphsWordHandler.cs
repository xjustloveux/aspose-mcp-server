using Aspose.Words;
using Aspose.Words.Drawing;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Results.Word.Paragraph;
using WordComment = Aspose.Words.Comment;
using WordShape = Aspose.Words.Drawing.Shape;

namespace AsposeMcpServer.Handlers.Word.Paragraph;

/// <summary>
///     Represents location information for a paragraph.
/// </summary>
/// <param name="Location">The location type (Body, Comment, TextBox, etc.).</param>
/// <param name="CommentInfo">Comment information if the paragraph is in a comment.</param>
internal record ParagraphLocation(string Location, string? CommentInfo);

/// <summary>
///     Handler for getting paragraphs from Word documents.
/// </summary>
[ResultType(typeof(GetParagraphsWordResult))]
public class GetParagraphsWordHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "get";

    /// <summary>
    ///     Gets all paragraphs from the document with optional filtering.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Optional: sectionIndex, includeEmpty, styleFilter, includeCommentParagraphs, includeTextboxParagraphs
    /// </param>
    /// <returns>JSON string containing paragraph information.</returns>
    public override object Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var getParams = ExtractGetParagraphsParameters(parameters);

        var doc = context.Document;

        var paragraphs = GetBaseParagraphs(doc, getParams.SectionIndex, getParams.IncludeCommentParagraphs);
        paragraphs = ApplyFilters(paragraphs, getParams.IncludeEmpty, getParams.StyleFilter,
            getParams.IncludeTextboxParagraphs);

        var paragraphList = BuildParagraphList(paragraphs);

        var result = new GetParagraphsWordResult
        {
            Count = paragraphs.Count,
            Filters = new ParagraphFilters
            {
                SectionIndex = getParams.SectionIndex,
                IncludeEmpty = getParams.IncludeEmpty,
                StyleFilter = getParams.StyleFilter,
                IncludeCommentParagraphs = getParams.IncludeCommentParagraphs,
                IncludeTextboxParagraphs = getParams.IncludeTextboxParagraphs
            },
            Paragraphs = paragraphList
        };

        return result;
    }

    private static List<Aspose.Words.Paragraph> GetBaseParagraphs(Document doc, int? sectionIndex,
        bool includeCommentParagraphs)
    {
        if (sectionIndex.HasValue)
        {
            if (sectionIndex.Value < 0 || sectionIndex.Value >= doc.Sections.Count)
                throw new ArgumentException($"sectionIndex must be between 0 and {doc.Sections.Count - 1}");
            return doc.Sections[sectionIndex.Value].Body
                .GetChildNodes(NodeType.Paragraph, includeCommentParagraphs).Cast<Aspose.Words.Paragraph>().ToList();
        }

        if (includeCommentParagraphs)
            return doc.GetChildNodes(NodeType.Paragraph, true).Cast<Aspose.Words.Paragraph>().ToList();

        List<Aspose.Words.Paragraph> paragraphs = [];
        foreach (var section in doc.Sections.Cast<Section>())
        {
            var bodyParagraphs = section.Body.GetChildNodes(NodeType.Paragraph, false)
                .Cast<Aspose.Words.Paragraph>()
                .ToList();
            paragraphs.AddRange(bodyParagraphs);
        }

        return paragraphs;
    }

    private static List<Aspose.Words.Paragraph> ApplyFilters(List<Aspose.Words.Paragraph> paragraphs,
        bool includeEmpty, string? styleFilter, bool includeTextboxParagraphs)
    {
        if (!includeEmpty)
            paragraphs = paragraphs.Where(p => !string.IsNullOrWhiteSpace(p.GetText())).ToList();

        if (!string.IsNullOrEmpty(styleFilter))
            paragraphs = paragraphs.Where(p => p.ParagraphFormat.Style?.Name == styleFilter).ToList();

        if (!includeTextboxParagraphs)
            paragraphs = paragraphs.Where(p => !IsInTextBox(p)).ToList();

        return paragraphs;
    }

    private static bool IsInTextBox(Aspose.Words.Paragraph p)
    {
        var shapeAncestor = p.GetAncestor(NodeType.Shape);
        if (shapeAncestor is WordShape { ShapeType: ShapeType.TextBox })
            return true;

        var currentNode = p.ParentNode;
        while (currentNode != null)
        {
            if (currentNode is WordShape { ShapeType: ShapeType.TextBox })
                return true;
            currentNode = currentNode.ParentNode;
        }

        return false;
    }

    private static List<ParagraphInfo> BuildParagraphList(List<Aspose.Words.Paragraph> paragraphs)
    {
        List<ParagraphInfo> paragraphList = [];

        for (var i = 0; i < paragraphs.Count; i++)
        {
            var para = paragraphs[i];
            var text = para.GetText().Trim();
            var paragraphLocation = DetermineLocation(para);

            var paraInfo = new ParagraphInfo
            {
                Index = i,
                Location = paragraphLocation.Location,
                Style = para.ParagraphFormat.Style?.Name,
                Text = text.Length > 100 ? text[..100] + "..." : text,
                TextLength = text.Length,
                CommentInfo = paragraphLocation.CommentInfo
            };

            paragraphList.Add(paraInfo);
        }

        return paragraphList;
    }

    private static ParagraphLocation DetermineLocation(Aspose.Words.Paragraph para)
    {
        if (para.ParentNode == null)
            return new ParagraphLocation("Body", null);

        var commentAncestor = para.GetAncestor(NodeType.Comment);
        if (commentAncestor != null)
        {
            var commentInfo = commentAncestor is WordComment comment
                ? $"ID: {comment.Id}, Author: {comment.Author}"
                : null;
            return new ParagraphLocation("Comment", commentInfo);
        }

        var shapeAncestor = para.GetAncestor(NodeType.Shape);
        if (shapeAncestor != null)
        {
            var location = shapeAncestor is WordShape { ShapeType: ShapeType.TextBox } ? "TextBox" : "Shape";
            return new ParagraphLocation(location, null);
        }

        var bodyAncestor = para.GetAncestor(NodeType.Body);
        if (bodyAncestor == null || para.ParentNode.NodeType != NodeType.Body)
            return new ParagraphLocation(para.ParentNode.NodeType.ToString(), null);

        return new ParagraphLocation("Body", null);
    }

    /// <summary>
    ///     Extracts get paragraphs parameters from operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted get paragraphs parameters.</returns>
    private static GetParagraphsParameters ExtractGetParagraphsParameters(OperationParameters parameters)
    {
        return new GetParagraphsParameters(
            parameters.GetOptional<int?>("sectionIndex"),
            parameters.GetOptional("includeEmpty", true),
            parameters.GetOptional<string?>("styleFilter"),
            parameters.GetOptional("includeCommentParagraphs", true),
            parameters.GetOptional("includeTextboxParagraphs", true)
        );
    }

    /// <summary>
    ///     Record to hold get paragraphs parameters.
    /// </summary>
    /// <param name="SectionIndex">The section index to filter by.</param>
    /// <param name="IncludeEmpty">Whether to include empty paragraphs.</param>
    /// <param name="StyleFilter">The style name to filter by.</param>
    /// <param name="IncludeCommentParagraphs">Whether to include paragraphs in comments.</param>
    /// <param name="IncludeTextboxParagraphs">Whether to include paragraphs in textboxes.</param>
    private sealed record GetParagraphsParameters(
        int? SectionIndex,
        bool IncludeEmpty,
        string? StyleFilter,
        bool IncludeCommentParagraphs,
        bool IncludeTextboxParagraphs);
}
