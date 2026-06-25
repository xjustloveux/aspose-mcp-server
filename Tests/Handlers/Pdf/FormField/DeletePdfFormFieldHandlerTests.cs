using System.Text;
using Aspose.Pdf;
using Aspose.Pdf.Forms;
using AsposeMcpServer.Handlers.Pdf.FormField;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Results.Pdf.FormField;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Pdf.FormField;

public class DeletePdfFormFieldHandlerTests : PdfHandlerTestBase
{
    private readonly DeletePdfFormFieldHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Delete()
    {
        Assert.Equal("delete", _handler.Operation);
    }

    #endregion

    #region Helper Methods

    private static Document CreateDocumentWithFormFields(int count)
    {
        var doc = new Document();
        var page = doc.Pages.Add();

        for (var i = 0; i < count; i++)
        {
            var field = new TextBoxField(page, new Rectangle(100, 700 - i * 30, 300, 720 - i * 30))
            {
                PartialName = $"field{i}"
            };
            doc.Form.Add(field);
        }

        return doc;
    }

    #endregion

    #region Basic Delete Operations

    [Fact]
    public void Execute_DeletesFormField()
    {
        var doc = CreateDocumentWithFormFields(2);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "fieldName", "field0" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("Deleted", result.Message);
        AssertModified(context);
    }

    [Fact]
    public void Execute_ReturnsFieldName()
    {
        var doc = CreateDocumentWithFormFields(1);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "fieldName", "field0" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("field0", result.Message);
    }

    [Fact]
    public void Execute_ReducesFormFieldCount()
    {
        var doc = CreateDocumentWithFormFields(3);
        var initialCount = doc.Form.Count;
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "fieldName", "field0" }
        });

        _handler.Execute(context, parameters);

        Assert.Equal(initialCount - 1, doc.Form.Count);
    }

    [Fact]
    public void Execute_RemovesCorrectField()
    {
        var doc = CreateDocumentWithFormFields(2);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "fieldName", "field0" }
        });

        _handler.Execute(context, parameters);

        Assert.Throws<ArgumentException>(() => doc.Form["field0"]);
        Assert.NotNull(doc.Form["field1"]);
    }

    [Fact]
    public void Execute_GetThenDeleteByReportedName_RemovesThatField()
    {
        // Round-trip: the name 'get' reports for a field must be the one delete(fieldName) removes.
        // The theorized nested-field divergence (PartialName != FullName) was investigated with hand-built
        // nested AcroForm PDFs and is NOT reachable here: document.Form (what get enumerates) yields only
        // top-level fields, where PartialName == FullName, so delete's Form.Delete(reportedName) round-trips.
        // Nested leaves are never surfaced by get and are safely rejected by delete (no wrong-element delete).
        var doc = new Document();
        var page = doc.Pages.Add();
        var child = new TextBoxField(page, new Rectangle(100, 700, 300, 720));
        doc.Form.Add(child, "address.first", 1);

        var getRes = (GetFormFieldsResult)new GetPdfFormFieldsHandler()
            .Execute(CreateContext(doc), CreateEmptyParameters());
        var reportedName = getRes.Items.Single().Name;

        new DeletePdfFormFieldHandler().Execute(CreateContext(doc), CreateParameters(new Dictionary<string, object?>
        {
            { "fieldName", reportedName }
        }));

        Assert.Empty(doc.Form);
    }

    [Fact]
    public void Execute_NestedAcroForm_GetReportsTopLevelAndDeleteRoundTrips()
    {
        // Adversarial fixture for the audited get->mutate addressing class: a TRUE nested AcroForm field
        // (parent "address" -> terminal child "first"; child FullName "address.first", PartialName "first").
        // Proves the round-trip is SAFE: get enumerates document.Form = top-level fields only, so it reports
        // "address" (never the leaf "first"); deleting that reported name round-trips; and a delete targeting
        // the leaf PartialName is safely rejected (no wrong-element deletion). Guards against a future change
        // to get that surfaces nested leaves, which would reintroduce the PartialName-vs-FullName divergence.
        var bytes = BuildNestedAcroFormPdf();

        using (var stream = new MemoryStream(bytes))
        using (var getDoc = new Document(stream))
        {
            var getRes = (GetFormFieldsResult)new GetPdfFormFieldsHandler()
                .Execute(CreateContext(getDoc), CreateEmptyParameters());
            var reportedName = Assert.Single(getRes.Items).Name;
            Assert.Equal("address", reportedName);
        }

        using (var stream = new MemoryStream(bytes))
        using (var leafDoc = new Document(stream))
        {
            var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(CreateContext(leafDoc),
                CreateParameters(new Dictionary<string, object?> { { "fieldName", "first" } })));
            Assert.Contains("not found", ex.Message);
        }

        using (var stream = new MemoryStream(bytes))
        using (var delDoc = new Document(stream))
        {
            _handler.Execute(CreateContext(delDoc),
                CreateParameters(new Dictionary<string, object?> { { "fieldName", "address" } }));
            Assert.Empty(delDoc.Form);
        }
    }

    private static byte[] BuildNestedAcroFormPdf()
    {
        // Minimal PDF with a 2-level AcroForm hierarchy: field "address" (non-terminal) -> field "first"
        // (terminal text field merged with its widget). Child FullName = "address.first", PartialName = "first".
        var enc = Encoding.Latin1;
        string[] objects =
        [
            "<< /Type /Catalog /Pages 2 0 R /AcroForm << /Fields [4 0 R] >> >>",
            "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Annots [5 0 R] >>",
            "<< /T (address) /Kids [5 0 R] >>",
            "<< /Type /Annot /Subtype /Widget /FT /Tx /T (first) /Parent 4 0 R /Rect [100 700 300 720] >>"
        ];
        var sb = new StringBuilder();
        sb.Append("%PDF-1.7\n");
        var offsets = new int[objects.Length];
        for (var i = 0; i < objects.Length; i++)
        {
            offsets[i] = enc.GetByteCount(sb.ToString());
            sb.Append($"{i + 1} 0 obj\n{objects[i]}\nendobj\n");
        }

        var xrefOffset = enc.GetByteCount(sb.ToString());
        sb.Append("xref\n");
        sb.Append($"0 {objects.Length + 1}\n");
        sb.Append("0000000000 65535 f \n");
        foreach (var off in offsets)
            sb.Append($"{off:D10} 00000 n \n");
        sb.Append($"trailer\n<< /Size {objects.Length + 1} /Root 1 0 R >>\nstartxref\n{xrefOffset}\n%%EOF");
        return enc.GetBytes(sb.ToString());
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutFieldName_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithFormFields(1);
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("fieldName", ex.Message);
    }

    [Fact]
    public void Execute_WithNonExistentFieldName_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithFormFields(1);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "fieldName", "nonExistent" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("not found", ex.Message);
    }

    [Fact]
    public void Execute_NoFormFields_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithPages(1);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "fieldName", "anyField" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("not found", ex.Message);
    }

    #endregion
}
