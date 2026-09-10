using Aspose.Words;
using AsposeMcpServer.Handlers.Word.Field;
using AsposeMcpServer.Helpers.Word;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Word.Field;

/// <summary>
///     R5-T01: what a refused field update leaves behind in a session's own document.
///     <para>
///         The file-mode fixtures show the bytes on disk are unchanged after a refusal, which is
///         true whether or not the in-memory document was touched — a refused request never saves.
///         A session keeps the document, so a partial mutation there survives the refusal and lands
///         in whatever the caller does next. That was untested, and the refusal happens inside
///         Aspose's own field-update traversal, where this server does not control how far it has
///         got.
///     </para>
/// </summary>
public class FieldDenialRollbackTests : WordHandlerTestBase
{
    private readonly UpdateFieldWordHandler _handler = new();

    /// <summary>
    ///     Builds a document whose only field resolves external content, so every update of it is
    ///     refused and nothing legitimate can change.
    /// </summary>
    /// <returns>The document.</returns>
    private static Document CreateDocumentWithOnlyARefusedField()
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Write("before ");
        builder.InsertField("INCLUDETEXT \"nowhere.docx\"");
        builder.Write(" after");
        return doc;
    }

    /// <summary>
    ///     Builds a document whose one allowed field carries a refused field inside its own code,
    ///     which is the case that reaches the callback part-way through a traversal.
    /// </summary>
    /// <returns>The document.</returns>
    private static Document CreateDocumentWithANestedRefusedField()
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.StartBookmark("bm1");
        builder.Writeln("target");
        builder.EndBookmark("bm1");

        var outer = builder.InsertField("REF bm1 \\h", null);
        builder.MoveTo(outer.Separator);
        builder.InsertField(" INCLUDETEXT \"nowhere.docx\" ", null);
        builder.MoveToDocumentEnd();
        builder.Write(" after");
        return doc;
    }

    /// <summary>
    ///     Builds a document whose first field is allowed and would change, followed by one that
    ///     carries a refused field inside its own code.
    ///     <para>
    ///         The order is the point. A denial raised part-way through Aspose's traversal does not
    ///         undo the fields already updated, so this arrangement left the first field's new
    ///         result in a document the session keeps. The fifth-round fixture happened to put the
    ///         refused field first, where there was nothing yet to lose (R5-T01).
    ///     </para>
    /// </summary>
    /// <returns>The document.</returns>
    private static Document CreateDocumentWithAChangingFieldBeforeANestedRefusedOne()
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.StartBookmark("bm1");
        builder.Writeln("target");
        builder.EndBookmark("bm1");

        builder.Write("allowed: ");
        builder.InsertField("REF bm1", null);
        builder.Writeln();

        builder.Write("outer: ");
        var outer = builder.InsertField("REF bm1 \\h", null);
        builder.MoveTo(outer.Separator);
        builder.InsertField(" INCLUDETEXT \"nowhere.docx\" ", null);
        builder.MoveToDocumentEnd();
        return doc;
    }

    /// <summary>
    ///     Describes a document's whole structure, so two states can be compared as a whole rather
    ///     than field by field.
    ///     <para>
    ///         Not the saved bytes: a docx carries a save timestamp, so two saves of an untouched
    ///         document differ. What has to be the same is the node graph, the text and every
    ///         field's code and result.
    ///     </para>
    /// </summary>
    /// <param name="doc">The document to describe.</param>
    /// <returns>Its structure as text.</returns>
    private static string Snapshot(Document doc)
    {
        var fields = string.Join("|", doc.Range.Fields
            .Select(f => $"{f.Type}:{f.GetFieldCode()}=>{f.Result}"));
        return $"{doc.GetChildNodes(NodeType.Any, true).Count}/{doc.GetText()}/{fields}";
    }

    [Fact]
    public void UpdateAll_WhenEveryFieldIsRefused_ShouldLeaveTheSessionDocumentUnchanged()
    {
        var doc = CreateDocumentWithOnlyARefusedField();
        var before = doc.GetText();
        var context = CreateContext(doc);

        var result = Assert.IsType<SuccessResult>(_handler.Execute(context, CreateParameters(
            new Dictionary<string, object?> { { "updateAll", true } })));

        Assert.Contains("Updated 0 field(s)", result.Message, StringComparison.Ordinal);
        Assert.Equal(before, doc.GetText());
        AssertNotModified(context);
    }

    /// <summary>
    ///     The nested case: the refusal is raised from inside the outer field's update, so the
    ///     traversal is interrupted where this server did not choose.
    /// </summary>
    [Fact]
    public void UpdateAll_WhenARefusedFieldIsNestedInAnAllowedOne_ShouldLeaveTheDocumentUnchanged()
    {
        var doc = CreateDocumentWithANestedRefusedField();
        var before = Snapshot(doc);
        var context = CreateContext(doc);

        // The callback raises the refusal from inside Aspose's traversal, so the whole update-all
        // request throws rather than reporting a tally.
        Assert.Throws<ArgumentException>(() => _handler.Execute(context, CreateParameters(
            new Dictionary<string, object?> { { "updateAll", true } })));

        // What R5-T01 asks: the session keeps this document, so the interrupted traversal must not
        // have left anything of itself in it.
        Assert.Equal(before, Snapshot(doc));
        AssertNotModified(context);
    }

    /// <summary>
    ///     The single-field path refuses before touching the document at all, which is the property
    ///     that makes the traversal question above the only one worth asking.
    /// </summary>
    [Fact]
    public void UpdateOne_WhenTheFieldIsRefused_ShouldThrowWithoutTouchingTheDocument()
    {
        var doc = CreateDocumentWithOnlyARefusedField();
        var before = doc.GetText();
        var context = CreateContext(doc);

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, CreateParameters(
            new Dictionary<string, object?> { { "fieldIndex", 0 } })));

        Assert.Equal(before, doc.GetText());
        AssertNotModified(context);
    }

    /// <summary>
    ///     R5-T01: an allowed field that would change, updated before the refusal is raised, must
    ///     not survive in the session's document.
    /// </summary>
    [Fact]
    public void UpdateAll_WhenAnAllowedFieldIsUpdatedBeforeTheRefusal_ShouldLeaveNothingBehind()
    {
        var doc = CreateDocumentWithAChangingFieldBeforeANestedRefusedOne();
        var before = Snapshot(doc);
        var context = CreateContext(doc);

        // The fixture only means anything if that first field really would change.
        Assert.Contains("FieldRef:", before, StringComparison.Ordinal);
        Assert.DoesNotContain("=>target", before, StringComparison.Ordinal);

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, CreateParameters(
            new Dictionary<string, object?> { { "updateAll", true } })));

        Assert.Equal(before, Snapshot(doc));
        AssertNotModified(context);
    }

    /// <summary>
    ///     R7-W01: the single-field path checked only the selected field's own type, so a nested
    ///     field it would have updated was resolved before the denial was raised.
    /// </summary>
    [Fact]
    public void UpdateOne_WhenTheSelectedFieldNestsAnUpdatableAndARefusedField_ShouldChangeNothing()
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.StartBookmark("bm1");
        builder.Writeln("target");
        builder.EndBookmark("bm1");

        var outer = builder.InsertField("REF bm1 \\h", null);
        builder.MoveTo(outer.Separator);
        builder.InsertField(" REF bm1 ", null);
        builder.InsertField(" INCLUDETEXT \"nowhere.docx\" ", null);
        builder.MoveToDocumentEnd();

        var before = Snapshot(doc);

        Assert.Throws<ArgumentException>(() => WordFieldPolicy.UpdateField(doc.Range.Fields[0]));

        Assert.Equal(before, Snapshot(doc));
    }

    /// <summary>
    ///     R7-W01: containment was decided by walking siblings from a field's start to its end,
    ///     which never arrives when the two markers sit in different paragraphs. The preflight then
    ///     reported the nested denial as not contained and updated everything before it.
    /// </summary>
    [Fact]
    public void UpdateAll_WhenTheOuterFieldSpansAParagraphBreak_ShouldStillSeeTheNestedRefusal()
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.StartBookmark("bm2");
        builder.Writeln("first");
        builder.Writeln("second");
        builder.EndBookmark("bm2");

        builder.Write("top: ");
        builder.InsertField("REF bm2", null);
        builder.Writeln();

        var outer = builder.InsertField("REF bm2 \\h", null);
        builder.MoveTo(outer.Separator);
        builder.Writeln("spans");
        builder.InsertField(" INCLUDETEXT \"nowhere.docx\" ", null);
        builder.MoveToDocumentEnd();

        var before = Snapshot(doc);
        var context = CreateContext(doc);

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, CreateParameters(
            new Dictionary<string, object?> { { "updateAll", true } })));

        Assert.Equal(before, Snapshot(doc));
        AssertNotModified(context);
    }

    [Fact]
    public void ThePolicy_ShouldReportARefusalRatherThanAnUpdate()
    {
        var doc = CreateDocumentWithOnlyARefusedField();

        var tally = WordFieldPolicy.UpdateAllowedFields(doc);

        Assert.Equal(0, tally.Updated);
        Assert.Equal(1, tally.Refused);
        Assert.Equal(0, tally.Locked);
    }
}
