using Aspose.Words;
using Aspose.Words.Fields;
using AsposeMcpServer.Tests.Infrastructure;
using AsposeMcpServer.Tools.Word;

namespace AsposeMcpServer.Tests.Handlers.Word.Field;

/// <summary>
///     R8-W01: the edit path takes a whole field code from the caller, and it must go through the
///     same allowlist the add path applies to a field type.
///     <para>
///         The handler removed the existing field-code runs and wrote the caller's text in their
///         place before anything looked at what that text said. The type it checked afterwards was
///         the field's old type, which the edit had already replaced, and with
///         <c>updateField=false</c> nothing looked at all.
///     </para>
///     <para>
///         Measured on the pinned Aspose.Words, in this order. Against the shipped implementation
///         the code was written <em>outside</em> the code region — the document came back as
///         <c>FieldNone</c> with an empty code and the text visible in the body — so the refused
///         command was inert and the reported "Field code updated" was not true either. Correcting
///         that (the builder now writes inside the code region) makes the same input come back as a
///         live <c>FieldIncludeText</c> carrying the caller's path, which is what these fixtures
///         exist to prevent. The fix to the operation and this guard belong together: the first one
///         alone would have created the hole (R8-W01).
///     </para>
/// </summary>
public class EditFieldCodePolicyTests : WordTestBase
{
    private readonly WordFieldTool _tool;

    /// <summary>Builds the tool under test.</summary>
    public EditFieldCodePolicyTests()
    {
        _tool = new WordFieldTool();
    }

    /// <summary>Writes a document carrying one permitted field.</summary>
    /// <param name="name">Fixture file name.</param>
    /// <returns>The document path.</returns>
    private string DocumentWithOneField(string name)
    {
        var path = CreateTestFilePath(name);
        var document = new Document();
        var builder = new DocumentBuilder(document);
        builder.InsertField("PAGE");
        document.Save(path);
        return path;
    }

    /// <summary>Everything about the document that a refused edit must leave untouched.</summary>
    /// <param name="path">The document path.</param>
    /// <returns>Bytes, field codes and field types.</returns>
    private static (byte[] Bytes, string Codes) Snapshot(string path)
    {
        var document = new Document(path);
        var codes = string.Join(" | ", document.Range.Fields
            .Select(f => $"{f.Type}:{f.GetFieldCode()}"));

        // Fully qualified: this namespace has a sibling called File, which is what an unqualified
        // name binds to here.
        return (System.IO.File.ReadAllBytes(path), codes);
    }

    /// <param name="fieldCode">The refused field code.</param>
    /// <param name="updateField">Whether the caller also asked for an update.</param>
    [Theory]
    [InlineData("INCLUDETEXT \"C:\\\\secrets.txt\"", true)]
    [InlineData("INCLUDETEXT \"C:\\\\secrets.txt\"", false)]
    [InlineData("INCLUDEPICTURE \"http://127.0.0.1/x.png\"", false)]
    [InlineData("DDEAUTO Excel Sheet1 R1C1", false)]
    [InlineData("LINK Excel.Sheet.12 \"C:\\\\book.xlsx\"", false)]
    [InlineData("MACROBUTTON Shell \"run me\"", false)]
    [InlineData("IMPORT \"C:\\\\secrets.txt\"", false)]
    public void EditingAFieldCodeToADisallowedCommand_ShouldBeRefusedWithoutTouchingTheDocument(
        string fieldCode, bool updateField)
    {
        var path = DocumentWithOneField($"edit_field_policy_{Guid.NewGuid():N}.docx");
        var before = Snapshot(path);

        var exception = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("edit", path, outputPath: path,
                fieldIndex: 0, fieldCode: fieldCode, updateField: updateField));

        Assert.Contains("not available through this server", exception.Message,
            StringComparison.Ordinal);

        var after = Snapshot(path);
        Assert.Equal(before.Codes, after.Codes);
        Assert.Equal(before.Bytes, after.Bytes);
    }

    [Fact]
    public void APermittedCommandCarryingARefusedOne_ShouldStillBeRefused()
    {
        var path = DocumentWithOneField("edit_field_nested.docx");
        var before = Snapshot(path);

        var exception = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("edit", path, outputPath: path,
                fieldIndex: 0, fieldCode: "IF 1 = 1 {INCLUDETEXT \"C:\\secrets.txt\"} \"\""));

        Assert.Contains("Unsupported field type", exception.Message, StringComparison.Ordinal);
        Assert.Equal(before.Bytes, Snapshot(path).Bytes);
    }

    [Fact]
    public void EditingAFieldCodeToAPermittedCommand_ShouldStillWork()
    {
        // The guard has to leave the feature working, or it is just a disabled operation.
        var path = DocumentWithOneField("edit_field_allowed.docx");

        _tool.Execute("edit", path, outputPath: path,
            fieldIndex: 0, fieldCode: "NUMPAGES", updateField: true);

        var document = new Document(path);
        var field = document.Range.Fields.First();
        Assert.Equal(FieldType.FieldNumPages, field.Type);
        Assert.Contains("NUMPAGES", field.GetFieldCode(), StringComparison.Ordinal);
    }
}
