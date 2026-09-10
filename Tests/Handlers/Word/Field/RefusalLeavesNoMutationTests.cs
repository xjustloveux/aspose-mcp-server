using Aspose.Words;
using AsposeMcpServer.Tests.Infrastructure;
using AsposeMcpServer.Tools.Word;

namespace AsposeMcpServer.Tests.Handlers.Word.Field;

/// <summary>
///     R8-W02: three update paths wrote to the document before the check that can refuse.
///     <para>
///         Adding a table of contents wrote the heading and inserted the field, editing a hyperlink
///         assigned the address, and editing a field rewrote its code and its lock state — all
///         before <c>UpdateAllowedFields</c> or <c>UpdateField</c> could refuse because of a
///         disallowed field the update would resolve. The caller was told the operation had failed
///         while the half-applied change stayed behind.
///     </para>
///     <para>
///         These fixtures work through a <em>session</em>, not a file. A refused file operation
///         never reaches its save, so the bytes on disk are unchanged whether the handler mutated
///         the document or not — asserting on them proves nothing about this defect. The session
///         document is the thing that keeps the mutation, and it is what is inspected here.
///     </para>
/// </summary>
public class RefusalLeavesNoMutationTests : WordTestBase
{
    /// <summary>
    ///     Writes a document whose REF field carries a nested INCLUDETEXT, so a document-wide
    ///     update is a refusal.
    /// </summary>
    /// <param name="name">Fixture file name.</param>
    /// <returns>The document path.</returns>
    private string DocumentWithANestedDisallowedField(string name)
    {
        var path = CreateTestFilePath(name);
        var document = new Document();
        var builder = new DocumentBuilder(document);

        builder.Writeln("Body text");

        // A permitted outer field whose code contains a refused one: updating the outer field is
        // what resolves the inner, which is the shape the preflight exists to catch.
        var outer = builder.InsertField("REF bookmark");
        var separator = outer.Separator;
        Assert.NotNull(separator);
        builder.MoveTo(separator);
        builder.InsertField("INCLUDETEXT \"C:\\\\secrets.txt\"");

        document.Save(path);
        return path;
    }

    /// <summary>
    ///     Writes a document whose hyperlink field itself carries a nested disallowed field, which
    ///     is what makes updating that field a refusal. One nested in some other field is not
    ///     something editing this hyperlink would resolve, so there would be nothing to refuse.
    /// </summary>
    /// <param name="name">Fixture file name.</param>
    /// <returns>The document path.</returns>
    private string DocumentWithADisallowedFieldInsideTheHyperlink(string name)
    {
        var path = CreateTestFilePath(name);
        var document = new Document();
        var builder = new DocumentBuilder(document);

        builder.Writeln("Body text");
        var hyperlink = builder.InsertField("HYPERLINK \"https://example.com/\"");
        var separator = hyperlink.Separator;
        Assert.NotNull(separator);
        builder.MoveTo(separator);
        builder.InsertField("INCLUDETEXT \"C:\\\\secrets.txt\"");

        document.Save(path);
        return path;
    }

    /// <summary>What the session document holds, as text plus its field codes.</summary>
    /// <param name="sessionId">The open session.</param>
    /// <returns>A single string describing the document's content.</returns>
    private string SessionState(string sessionId)
    {
        var document = SessionManager.GetDocument<Document>(sessionId);
        var codes = string.Join(" | ", document.Range.Fields
            .Select(f => $"{f.Type}:{f.GetFieldCode()}"));

        return document.GetText() + " || " + codes
               + " || paragraphs=" + document.FirstSection.Body.Paragraphs.Count;
    }

    [Fact]
    public void ARefusedTableOfContents_ShouldLeaveTheSessionDocumentUnchanged()
    {
        var sessionId = OpenSession(DocumentWithANestedDisallowedField("refusal_toc.docx"));
        var before = SessionState(sessionId);

        Assert.ThrowsAny<ArgumentException>(() =>
            new WordReferenceTool(SessionManager).Execute(
                "add_toc", sessionId: sessionId, title: "Contents"));

        Assert.Equal(before, SessionState(sessionId));
    }

    [Fact]
    public void ARefusedHyperlinkEdit_ShouldLeaveTheAddressAlone()
    {
        var sessionId = OpenSession(
            DocumentWithADisallowedFieldInsideTheHyperlink("refusal_hyperlink.docx"));
        var before = SessionState(sessionId);

        Assert.ThrowsAny<ArgumentException>(() =>
            new WordHyperlinkTool(SessionManager).Execute(
                "edit", sessionId: sessionId, hyperlinkIndex: 0,
                url: "https://elsewhere.example/"));

        Assert.Equal(before, SessionState(sessionId));
        Assert.DoesNotContain("elsewhere.example", SessionState(sessionId), StringComparison.Ordinal);
    }

    [Fact]
    public void ARefusedFieldEdit_ShouldLeaveTheFieldCodeAlone()
    {
        var sessionId = OpenSession(DocumentWithANestedDisallowedField("refusal_field_edit.docx"));
        var before = SessionState(sessionId);

        Assert.ThrowsAny<ArgumentException>(() =>
            new WordFieldTool(SessionManager).Execute(
                "edit", sessionId: sessionId, fieldIndex: 0, fieldCode: "NUMPAGES",
                updateField: true));

        Assert.Equal(before, SessionState(sessionId));
        Assert.DoesNotContain("NUMPAGES", SessionState(sessionId), StringComparison.Ordinal);
    }
}
