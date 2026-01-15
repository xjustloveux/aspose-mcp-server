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
        var sectionIndex = parameters.GetOptional<int?>("sectionIndex");
        var includeEmpty = parameters.GetOptional("includeEmpty", true);
        var styleFilter = parameters.GetOptional<string?>("styleFilter");
        var includeCommentParagraphs = parameters.GetOptional("includeCommentParagraphs", true);
        var includeTextboxParagraphs = parameters.GetOptional("includeTextboxParagraphs", true);

        var doc = context.Document;

        List<Aspose.Words.Paragraph> paragraphs;
        if (sectionIndex.HasValue)
        {
            if (sectionIndex.Value < 0 || sectionIndex.Value >= doc.Sections.Count)
                throw new ArgumentException($"sectionIndex must be between 0 and {doc.Sections.Count - 1}");
            paragraphs = doc.Sections[sectionIndex.Value].Body
                .GetChildNodes(NodeType.Paragraph, includeCommentParagraphs).Cast<Aspose.Words.Paragraph>().ToList();
        }
        else
        {
            if (includeCommentParagraphs)
            {
                paragraphs = doc.GetChildNodes(NodeType.Paragraph, true).Cast<Aspose.Words.Paragraph>().ToList();
            }
            else
            {
                paragraphs = [];
                foreach (var section in doc.Sections.Cast<Section>())
                {
                    var bodyParagraphs = section.Body.GetChildNodes(NodeType.Paragraph, false)
                        .Cast<Aspose.Words.Paragraph>()
                        .ToList();
                    paragraphs.AddRange(bodyParagraphs);
                }
            }
        }

        if (!includeEmpty) paragraphs = paragraphs.Where(p => !string.IsNullOrWhiteSpace(p.GetText())).ToList();
        if (!string.IsNullOrEmpty(styleFilter))
            paragraphs = paragraphs.Where(p => p.ParagraphFormat.Style?.Name == styleFilter).ToList();

        if (!includeTextboxParagraphs)
            paragraphs = paragraphs.Where(p =>
            {
                var shapeAncestor = p.GetAncestor(NodeType.Shape);
                if (shapeAncestor is WordShape { ShapeType: ShapeType.TextBox }) return false;
                var currentNode = p.ParentNode;
                while (currentNode != null)
                {
                    if (currentNode.NodeType == NodeType.Shape)
                        if (currentNode is WordShape { ShapeType: ShapeType.TextBox })
                            return false;
                    currentNode = currentNode.ParentNode;
                }

                return true;
            }).ToList();

        List<object> paragraphList = [];
        for (var i = 0; i < paragraphs.Count; i++)
        {
            var para = paragraphs[i];
            var text = para.GetText().Trim();
            var location = "Body";
            string? commentInfo = null;

            if (para.ParentNode != null)
            {
                var commentAncestor = para.GetAncestor(NodeType.Comment);
                if (commentAncestor != null)
                {
                    location = "Comment";
                    if (commentAncestor is WordComment comment)
                        commentInfo = $"ID: {comment.Id}, Author: {comment.Author}";
                }
                else
                {
                    var shapeAncestor = para.GetAncestor(NodeType.Shape);
                    if (shapeAncestor != null)
                    {
                        location = shapeAncestor is WordShape { ShapeType: ShapeType.TextBox } ? "TextBox" : "Shape";
                    }
                    else
                    {
                        var bodyAncestor = para.GetAncestor(NodeType.Body);
                        if (bodyAncestor == null || para.ParentNode.NodeType != NodeType.Body)
                            location = para.ParentNode.NodeType.ToString();
                    }
                }
            }

            var paraInfo = new Dictionary<string, object?>
            {
                ["index"] = i,
                ["location"] = location,
                ["style"] = para.ParagraphFormat.Style?.Name,
                ["text"] = text.Length > 100 ? text[..100] + "..." : text,
                ["textLength"] = text.Length
            };

            if (commentInfo != null) paraInfo["commentInfo"] = commentInfo;
            paragraphList.Add(paraInfo);
        }

        var result = new
        {
            count = paragraphs.Count,
            filters = new
                { sectionIndex, includeEmpty, styleFilter, includeCommentParagraphs, includeTextboxParagraphs },
            paragraphs = paragraphList
        };

        return JsonSerializer.Serialize(result, JsonDefaults.Indented);
    }
}
