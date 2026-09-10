using System.Text;
using Aspose.Words;
using Aspose.Words.Fields;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;
using AsposeMcpServer.Tools.Word;

namespace AsposeMcpServer.Tests.Handlers.Word.Field;

/// <summary>
///     No <c>word_field</c> entry point may create or resolve a field that reads a file or fetches
///     a URL (R3-S01).
///     <para>
///         Restricting the header and footer writer closed one door. <c>word_field add</c> inserted
///         whatever code the caller supplied and called <c>field.Update()</c> immediately, so the
///         same <c>INCLUDETEXT</c> read was available through a second entry — and this one did not
///         even need a later refresh to trigger it. The update paths additionally called
///         <c>Document.UpdateFields()</c>, which resolves fields that arrived in the caller's own
///         input file rather than ones this server created.
///     </para>
/// </summary>
public class ActiveFieldPolicyTests : WordTestBase
{
    private readonly WordFieldTool _tool = new();

    /// <summary>
    ///     Writes a document and returns its path.
    /// </summary>
    /// <param name="name">File name under the test directory.</param>
    /// <returns>The path written.</returns>
    private string WriteDocument(string name)
    {
        var path = CreateTestFilePath(name);
        var doc = new Document();
        new DocumentBuilder(doc).Writeln("body");
        doc.Save(path);
        return path;
    }

    /// <summary>
    ///     Creates the canary file the refused fields would read.
    /// </summary>
    /// <returns>Path of the file that must stay unread.</returns>
    private string WriteCanary()
    {
        var path = Path.Combine(TestDir, "not-for-the-caller.txt");
        System.IO.File.WriteAllText(path, "CANARY_STRING_12345", Encoding.UTF8);
        return path;
    }

    /// <param name="fieldType">A field type that reaches outside the document.</param>
    [Theory]
    [InlineData("INCLUDETEXT")]
    [InlineData("INCLUDEPICTURE")]
    [InlineData("DDEAUTO")]
    [InlineData("LINK")]
    public void AddingAnActiveField_ShouldBeRefused(string fieldType)
    {
        var canary = WriteCanary();
        var docPath = WriteDocument("active_field_add.docx");

        var exception = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("add", docPath, outputPath: docPath,
                fieldType: fieldType, fieldArgument: "\"" + canary.Replace("\\", "\\\\") + "\""));

        Assert.Contains(fieldType, exception.Message);

        var written = new Document(docPath);
        Assert.DoesNotContain("CANARY_STRING_12345", written.GetText());
    }

    [Fact]
    public void AddingADocumentedField_ShouldStillWork()
    {
        var docPath = WriteDocument("active_field_ok.docx");

        _tool.Execute("add", docPath, outputPath: docPath, fieldType: "DATE");

        var written = new Document(docPath);
        Assert.Contains(written.Range.Fields,
            f => f.Type == FieldType.FieldDate);
    }

    /// <summary>
    ///     A field that arrived in the input document must not be resolved by a refresh either.
    /// </summary>
    [Fact]
    public void UpdatingADocumentThatAlreadyCarriesAnActiveField_ShouldNotResolveIt()
    {
        var canary = WriteCanary();
        var docPath = CreateTestFilePath("active_field_existing.docx");

        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln("body");
        builder.InsertField(" INCLUDETEXT \"" + canary.Replace("\\", "\\\\") + "\" ", null);
        builder.InsertField(FieldType.FieldDate, true);
        doc.Save(docPath);

        _tool.Execute("update", docPath, outputPath: docPath);

        var written = new Document(docPath);
        Assert.DoesNotContain("CANARY_STRING_12345", written.GetText());
    }

    /// <summary>
    ///     Checking a field's own type only answers for the field named. A field's code can
    ///     contain another field, and updating the outer one updates what is inside it: measured
    ///     against the pinned Aspose.Words, REF, HYPERLINK and TOC each resolved a nested
    ///     INCLUDETEXT and read the canary, while the caller was told the dangerous field had
    ///     been skipped (R4-S01).
    /// </summary>
    /// <param name="outerCode">Field code of an allowed outer field.</param>
    [Theory]
    [InlineData("HYPERLINK \"https://example.invalid\"")]
    [InlineData("REF bm1 \\h")]
    [InlineData("TOC \\o \"1-3\"")]
    public void UpdatingAnAllowedFieldThatContainsADangerousOne_ShouldBeRefused(string outerCode)
    {
        var canary = WriteCanary();
        var docPath = CreateTestFilePath("nested_field.docx");

        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.StartBookmark("bm1");
        builder.Writeln("target");
        builder.EndBookmark("bm1");

        var outer = builder.InsertField(outerCode, null);
        builder.MoveTo(outer.Separator);
        builder.InsertField(" INCLUDETEXT \"" + canary.Replace("\\", "\\\\") + "\" ", null);
        doc.Save(docPath);

        var before = System.IO.File.ReadAllBytes(docPath);

        Assert.Throws<ArgumentException>(() =>
            _tool.Execute("update", docPath, outputPath: docPath));

        // The refusal must leave the document alone and the canary unread.
        Assert.Equal(before, System.IO.File.ReadAllBytes(docPath));
        var written = new Document(docPath);
        Assert.DoesNotContain("CANARY_STRING_12345", written.GetText());
    }

    /// <summary>
    ///     The same nesting reached through the single-field entry point, which updates one named
    ///     field rather than every allowed one.
    /// </summary>
    [Fact]
    public void UpdatingOneFieldThatContainsADangerousOne_ShouldBeRefused()
    {
        var canary = WriteCanary();
        var docPath = CreateTestFilePath("nested_field_single.docx");

        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        var outer = builder.InsertField("HYPERLINK \"https://example.invalid\"", null);
        builder.MoveTo(outer.Separator);
        builder.InsertField(" INCLUDETEXT \"" + canary.Replace("\\", "\\\\") + "\" ", null);
        doc.Save(docPath);

        Assert.Throws<ArgumentException>(() =>
            _tool.Execute("update", docPath, outputPath: docPath, fieldIndex: 0));

        var written = new Document(docPath);
        Assert.DoesNotContain("CANARY_STRING_12345", written.GetText());
    }

    /// <summary>
    ///     A locked field this server refuses anyway was counted both as locked and as refused, so
    ///     the reported number of updated fields went below zero (R4-C01).
    /// </summary>
    [Fact]
    public void UpdatingADocumentWhoseOnlyFieldIsLockedAndDangerous_ShouldNotReportANegativeCount()
    {
        var canary = WriteCanary();
        var docPath = CreateTestFilePath("locked_dangerous.docx");

        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        var field = builder.InsertField(" INCLUDETEXT \"" + canary.Replace("\\", "\\\\") + "\" ", null);
        field.IsLocked = true;
        doc.Save(docPath);

        var result = _tool.Execute("update", docPath, outputPath: docPath);
        var message = GetResultData<SuccessResult>(result).Message;

        Assert.Contains("Updated 0 field(s)", message);
        Assert.DoesNotContain("-1", message);

        var written = new Document(docPath);
        Assert.DoesNotContain("CANARY_STRING_12345", written.GetText());
    }

    /// <summary>
    ///     Counts stay exclusive when the document mixes all three cases.
    /// </summary>
    [Fact]
    public void UpdatingAMixedDocument_ShouldCountEachFieldOnce()
    {
        var canary = WriteCanary();
        var docPath = CreateTestFilePath("mixed_fields.docx");

        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.InsertField(FieldType.FieldDate, true);
        var locked = builder.InsertField(FieldType.FieldPage, true);
        locked.IsLocked = true;
        builder.InsertField(" INCLUDETEXT \"" + canary.Replace("\\", "\\\\") + "\" ", null);
        doc.Save(docPath);

        var result = _tool.Execute("update", docPath, outputPath: docPath);
        var message = GetResultData<SuccessResult>(result).Message;

        Assert.Contains("Updated 1 field(s)", message);
        Assert.Contains("1 locked", message);
        Assert.Contains("1 field(s) that resolve external content", message);
    }
}
