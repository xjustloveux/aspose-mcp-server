using Aspose.Words;
using AsposeMcpServer.Tests.Infrastructure;
using AsposeMcpServer.Tools.Word;

namespace AsposeMcpServer.Tests.Handlers.Word.Field;

/// <summary>
///     R9-W01: a refused edit must not leave the field unlocked on the way out.
///     <para>
///         The document-wide preflight skips locked fields, because a locked field is not updated
///         and so cannot resolve anything. A request that unlocks the field and then updates it
///         made that reasoning false between the two steps: the preflight saw a locked field and
///         allowed the request, <c>ApplyLockState</c> unlocked it, and only then did the update
///         see the nested disallowed field and refuse. The caller was told the operation failed
///         while the field stayed unlocked — and, with a field code in the same request, rewritten
///         as well.
///     </para>
/// </summary>
public class LockedFieldRefusalAtomicityTests : WordTestBase
{
    /// <summary>
    ///     Writes a document whose <em>locked</em> REF field carries a nested INCLUDETEXT.
    /// </summary>
    /// <param name="name">Fixture file name.</param>
    /// <returns>The document path.</returns>
    private string DocumentWithALockedFieldAroundADisallowedOne(string name)
    {
        var path = CreateTestFilePath(name);
        var document = new Document();
        var builder = new DocumentBuilder(document);

        builder.Writeln("Body text");

        var outer = builder.InsertField("REF bookmark");
        var separator = outer.Separator;
        Assert.NotNull(separator);
        builder.MoveTo(separator);
        builder.InsertField("INCLUDETEXT \"C:\\\\secrets.txt\"");

        // Locked, so the preflight leaves it out of the fields it considers updatable.
        outer.IsLocked = true;

        document.Save(path);
        return path;
    }

    /// <summary>
    ///     Everything about the session document a refusal must leave untouched: its text, every
    ///     field's type, code and lock state, the paragraph count, and whether it is dirty.
    /// </summary>
    /// <param name="sessionId">The open session.</param>
    /// <returns>A single string describing the document's state.</returns>
    private string SessionState(string sessionId)
    {
        var document = SessionManager.GetDocument<Document>(sessionId);
        var codes = string.Join(" | ", document.Range.Fields
            .Select(f => $"{f.Type}:{f.GetFieldCode()}:locked={f.IsLocked}"));

        return document.GetText() + " || " + codes
               + " || paragraphs=" + document.FirstSection.Body.Paragraphs.Count
               + " || nodes=" + document.GetChildNodes(NodeType.Any, true).Count
               + " || dirty=" + SessionManager.GetSession(sessionId).IsDirty;
    }

    [Fact]
    public void ARefusedUnlockAndUpdate_ShouldLeaveTheFieldLocked()
    {
        var sessionId = OpenSession(
            DocumentWithALockedFieldAroundADisallowedOne("locked_refusal_unlock.docx"));
        var before = SessionState(sessionId);

        Assert.Contains("locked=True", before, StringComparison.Ordinal);

        Assert.ThrowsAny<ArgumentException>(() =>
            new WordFieldTool(SessionManager).Execute(
                "edit", sessionId: sessionId, fieldIndex: 0,
                unlockField: true, updateField: true));

        Assert.Equal(before, SessionState(sessionId));
    }

    [Fact]
    public void ARefusedUnlockWithANewCode_ShouldLeaveBothTheLockAndTheCode()
    {
        var sessionId = OpenSession(
            DocumentWithALockedFieldAroundADisallowedOne("locked_refusal_code.docx"));
        var before = SessionState(sessionId);

        Assert.ThrowsAny<ArgumentException>(() =>
            new WordFieldTool(SessionManager).Execute(
                "edit", sessionId: sessionId, fieldIndex: 0, fieldCode: "NUMPAGES",
                unlockField: true, updateField: true));

        Assert.Equal(before, SessionState(sessionId));
        Assert.DoesNotContain("NUMPAGES", SessionState(sessionId), StringComparison.Ordinal);
    }

    [Fact]
    public void UnlockingWithoutUpdating_ShouldStillWork()
    {
        // The control. Nothing is resolved when the caller does not ask for an update, so there is
        // nothing to refuse and the unlock must go through — otherwise the fix is a new defect.
        var sessionId = OpenSession(
            DocumentWithALockedFieldAroundADisallowedOne("locked_unlock_only.docx"));

        new WordFieldTool(SessionManager).Execute(
            "edit", sessionId: sessionId, fieldIndex: 0,
            unlockField: true, updateField: false);

        Assert.Contains("locked=False", SessionState(sessionId), StringComparison.Ordinal);
    }

    [Fact]
    public void AnOrdinaryUnlockAndUpdate_ShouldStillWork()
    {
        // The second control: with no disallowed field anywhere, the same request has to succeed.
        var path = CreateTestFilePath("locked_plain.docx");
        var document = new Document();
        var builder = new DocumentBuilder(document);
        builder.Writeln("Body text");
        builder.InsertField("NUMPAGES").IsLocked = true;
        document.Save(path);

        var sessionId = OpenSession(path);

        new WordFieldTool(SessionManager).Execute(
            "edit", sessionId: sessionId, fieldIndex: 0,
            unlockField: true, updateField: true);

        Assert.Contains("locked=False", SessionState(sessionId), StringComparison.Ordinal);
    }
}
