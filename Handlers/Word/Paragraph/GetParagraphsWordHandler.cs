using System.Text.Json;
using Aspose.Words;
using Aspose.Words.Drawing;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;
using WordComment = Aspose.Words.Comment;
using WordShape = Aspose.Words.Drawing.Shape;

namespace AsposeMcpServer.Handlers.Word.Paragraph;

/// <summary>
///     Handler for getting paragraphs from Word documents.
/// </summary>
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
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var getParams = ExtractGetParagraphsParameters(parameters);

        var doc = context.Document;

        var paragraphs = GetBaseParagraphs(doc, getParams.SectionIndex, getParams.IncludeCommentParagraphs);
        paragraphs = ApplyFilters(paragraphs, getParams.IncludeEmpty, getParams.StyleFilter,
            getParams.IncludeTextboxParagraphs);

        var paragraphList = BuildParagraphList(paragraphs);

        var result = new
        {
            count = paragraphs.Count,
            filters = new
            {
                sectionIndex = getParams.SectionIndex,
                includeEmpty = getParams.IncludeEmpty,
                styleFilter = getParams.StyleFilter,
                includeCommentParagraphs = getParams.IncludeCommentParagraphs,
                includeTextboxParagraphs = getParams.IncludeTextboxParagraphs
            },
            paragraphs = paragraphList
        };

        return JsonSerializer.Serialize(result, JsonDefaults.Indented);
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

    private static List<object> BuildParagraphList(List<Aspose.Words.Paragraph> paragraphs)
    {
        List<object> paragraphList = [];

        for (var i = 0; i < paragraphs.Count; i++)
        {
            var para = paragraphs[i];
            var text = para.GetText().Trim();
            var (location, commentInfo) = DetermineLocation(para);

            var paraInfo = new Dictionary<string, object?>
            {
                ["index"] = i,
                ["location"] = location,
                ["style"] = para.ParagraphFormat.Style?.Name,
                ["text"] = text.Length > 100 ? text[..100] + "..." : text,
                ["textLength"] = text.Length
            };

            if (commentInfo != null)
                paraInfo["commentInfo"] = commentInfo;

            paragraphList.Add(paraInfo);
        }

        return paragraphList;
    }

    private static (string location, string? commentInfo) DetermineLocation(Aspose.Words.Paragraph para)
    {
        if (para.ParentNode == null)
            return ("Body", null);

        var commentAncestor = para.GetAncestor(NodeType.Comment);
        if (commentAncestor != null)
        {
            var commentInfo = commentAncestor is WordComment comment
                ? $"ID: {comment.Id}, Author: {comment.Author}"
                : null;
            return ("Comment", commentInfo);
        }

        var shapeAncestor = para.GetAncestor(NodeType.Shape);
        if (shapeAncestor != null)
        {
            var location = shapeAncestor is WordShape { ShapeType: ShapeType.TextBox } ? "TextBox" : "Shape";
            return (location, null);
        }

        var bodyAncestor = para.GetAncestor(NodeType.Body);
        if (bodyAncestor == null || para.ParentNode.NodeType != NodeType.Body)
            return (para.ParentNode.NodeType.ToString(), null);

        return ("Body", null);
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
    private record GetParagraphsParameters(
        int? SectionIndex,
        bool IncludeEmpty,
        string? StyleFilter,
        bool IncludeCommentParagraphs,
        bool IncludeTextboxParagraphs);
}
