using Aspose.Words;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Handlers.Word.Comment;
using AsposeMcpServer.Helpers.Word;
using AsposeMcpServer.Tests.Infrastructure;
using WordsComment = Aspose.Words.Comment;

namespace AsposeMcpServer.Tests.Handlers.Word;

/// <summary>
///     Guards three addressing and clearing contracts that used to differ between operations:
///     header and footer clearing ignored the caller's request (RB-12), comment deletion counted
///     replies and used document order while get and reply used top-level comments ordered by date
///     (RB-13), and section-scoped table lookup scanned only the body while the document-wide
///     lookup also walked headers and footers (RB-14).
/// </summary>
public class WordIndexAndClearContractTests : TestBase
{
    /// <summary>
    ///     Builds a document with a header carrying one paragraph of text.
    /// </summary>
    /// <param name="headerText">Text to place in the primary header.</param>
    /// <returns>The document.</returns>
    private static Document DocumentWithHeader(string headerText)
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.Write(headerText);
        builder.MoveToDocumentStart();
        return doc;
    }

    [Fact]
    public void Clear_WithClearExistingFalse_ShouldKeepContent()
    {
        var doc = DocumentWithHeader("KEEP ME");
        var header = doc.FirstSection.HeadersFooters[HeaderFooterType.HeaderPrimary];

        WordHeaderFooterHelper.Clear(header, false, false);

        Assert.Contains("KEEP ME", header.GetText());
    }

    [Fact]
    public void Clear_WithClearExistingTrue_ShouldRemoveEverything()
    {
        var doc = DocumentWithHeader("REMOVE ME");
        var header = doc.FirstSection.HeadersFooters[HeaderFooterType.HeaderPrimary];

        WordHeaderFooterHelper.Clear(header, true, false);

        Assert.DoesNotContain("REMOVE ME", header.GetText());
        Assert.Equal(0, header.GetChildNodes(NodeType.Any, true).Count);
    }

    [Fact]
    public void Clear_WithTextOnly_ShouldEmptyTextButKeepStructure()
    {
        var doc = DocumentWithHeader("REMOVE ME");
        var header = doc.FirstSection.HeadersFooters[HeaderFooterType.HeaderPrimary];

        WordHeaderFooterHelper.Clear(header, true, true);

        Assert.DoesNotContain("REMOVE ME", header.GetText());
        Assert.True(header.GetChildNodes(NodeType.Paragraph, true).Count > 0);
    }

    /// <summary>
    ///     Builds a document with two top-level comments where the later-dated one appears first in
    ///     document order, and gives the first a reply.
    /// </summary>
    /// <returns>The document.</returns>
    private static Document DocumentWithRepliedComments()
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Write("body text");

        var newer = new WordsComment(doc, "Alice", "A", new DateTime(2026, 2, 1));
        newer.SetText("newer top level");
        builder.CurrentParagraph.AppendChild(newer);

        var reply = newer.AddReply("Bob", "B", new DateTime(2026, 2, 2), "a reply");
        Assert.NotNull(reply);

        var older = new WordsComment(doc, "Carol", "C", new DateTime(2026, 1, 1));
        older.SetText("older top level");
        builder.CurrentParagraph.AppendChild(older);

        return doc;
    }

    [Fact]
    public void CommentDelete_ShouldUseTheSameIndexSpaceAsGet()
    {
        var doc = DocumentWithRepliedComments();
        var expected = WordCommentHelper.GetTopLevelComments(doc);
        Assert.Equal(2, expected.Count);
        var targetText = expected[0].GetText();

        var handler = new DeleteWordCommentHandler();
        var context = new OperationContext<Document> { Document = doc };
        var parameters = new OperationParameters();
        parameters.Set("commentIndex", 0);
        handler.Execute(context, parameters);

        var remaining = WordCommentHelper.GetTopLevelComments(doc);
        Assert.Single(remaining);
        Assert.DoesNotContain(targetText.Trim(), remaining[0].GetText());
    }

    [Fact]
    public void SectionTables_ShouldIncludeHeaderTablesLikeTheDocumentWideLookup()
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.StartTable();
        builder.InsertCell();
        builder.Write("header table");
        builder.EndRow();
        builder.EndTable();

        builder.MoveToDocumentStart();
        builder.StartTable();
        builder.InsertCell();
        builder.Write("body table");
        builder.EndRow();
        builder.EndTable();

        var documentWide = WordTableHelper.GetTables(doc, null);
        var sectionScoped = WordTableHelper.GetTables(doc, 0);

        Assert.Equal(2, documentWide.Count);
        Assert.Equal(documentWide.Count, sectionScoped.Count);
        Assert.Equal(
            documentWide.Select(t => t.GetText().Trim()),
            sectionScoped.Select(t => t.GetText().Trim()));
    }
}
