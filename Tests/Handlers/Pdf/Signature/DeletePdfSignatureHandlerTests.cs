using Aspose.Pdf;
using Aspose.Pdf.Forms;
using AsposeMcpServer.Handlers.Pdf.Signature;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Results.Pdf.Signature;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Pdf.Signature;

public class DeletePdfSignatureHandlerTests : PdfHandlerTestBase
{
    private readonly DeletePdfSignatureHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Delete()
    {
        Assert.Equal("delete", _handler.Operation);
    }

    #endregion

    #region Helper Methods

    /// <summary>
    ///     Creates a PDF document with a signature field that is properly indexed.
    ///     Uses PartialName which is the field identifier used by Form.Delete.
    /// </summary>
    private static Document CreatePdfWithSignatureFieldPersisted(string name)
    {
        var document = new Document();
        var page = document.Pages.Add();
        var signatureField = new SignatureField(page, new Rectangle(100, 600, 300, 700))
        {
            PartialName = name
        };
        document.Form.Add(signatureField);
        return document;
    }

    #endregion

    #region Basic Delete Signature Operations

    [SkippableFact]
    public void Execute_DeletesSignatureByName()
    {
        SkipInEvaluationMode(AsposeLibraryType.Pdf);
        var document = CreatePdfWithSignatureFieldPersisted("TestSig");
        var initialCount = document.Form.Fields.OfType<SignatureField>().Count();
        var context = CreateContext(document);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "signatureName", "TestSig" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        if (!IsEvaluationMode(AsposeLibraryType.Pdf))
        {
            var remainingCount = document.Form.Fields.OfType<SignatureField>().Count();
            Assert.Equal(initialCount - 1, remainingCount);
            Assert.DoesNotContain(document.Form.Fields.OfType<SignatureField>(),
                f => f.PartialName == "TestSig");
        }

        AssertModified(context);
    }

    [SkippableFact]
    public void Execute_DeletesSignatureByIndex()
    {
        SkipInEvaluationMode(AsposeLibraryType.Pdf);
        var document = CreatePdfWithSignatureFieldPersisted("TestSig");
        var initialCount = document.Form.Fields.OfType<SignatureField>().Count();
        var context = CreateContext(document);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "signatureIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        if (!IsEvaluationMode(AsposeLibraryType.Pdf))
        {
            var remainingCount = document.Form.Fields.OfType<SignatureField>().Count();
            Assert.Equal(initialCount - 1, remainingCount);
        }

        AssertModified(context);
    }

    [SkippableFact]
    public void Execute_GetThenDeleteByReportedIndex_DeletesThatSignatureField()
    {
        SkipInEvaluationMode(AsposeLibraryType.Pdf);
        // True single-document round-trip: a field reported by 'get' at index N must be the one
        // delete(signatureIndex=N) removes — both enumerate all signature fields in document order. Running
        // get then delete on the SAME document also confirms get leaves it usable (does not dispose it).
        var doc = new Document();
        var page = doc.Pages.Add();
        doc.Form.Add(new SignatureField(page, new Rectangle(100, 600, 300, 700)) { PartialName = "SigA" });
        doc.Form.Add(new SignatureField(page, new Rectangle(100, 400, 300, 500)) { PartialName = "SigB" });

        var getRes = (GetSignaturesResult)new GetPdfSignaturesHandler()
            .Execute(CreateContext(doc), CreateEmptyParameters());
        var sigBIndex = getRes.Items.Single(s => s.Name == "SigB").Index;

        _handler.Execute(CreateContext(doc), CreateParameters(new Dictionary<string, object?>
        {
            { "signatureIndex", sigBIndex }
        }));

        var remaining = doc.Form.Fields.OfType<SignatureField>().ToList();
        Assert.Single(remaining);
        Assert.Equal("SigA", remaining[0].PartialName);
    }

    [Fact]
    public void Execute_WithoutNameOrIndex_ThrowsArgumentException()
    {
        var document = CreateEmptyDocument();
        var context = CreateContext(document);
        var parameters = CreateEmptyParameters();

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithNoSignatureFields_ThrowsArgumentException()
    {
        var document = CreateEmptyDocument();
        var context = CreateContext(document);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "signatureName", "NonExistent" }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [SkippableFact]
    public void Execute_WithInvalidSignatureName_ThrowsArgumentException()
    {
        SkipInEvaluationMode(AsposeLibraryType.Pdf);
        var document = CreatePdfWithSignatureFieldPersisted("TestSig");
        var context = CreateContext(document);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "signatureName", "WrongName" }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [SkippableFact]
    public void Execute_WithInvalidSignatureIndex_ThrowsArgumentException()
    {
        SkipInEvaluationMode(AsposeLibraryType.Pdf);
        var document = CreatePdfWithSignatureFieldPersisted("TestSig");
        var context = CreateContext(document);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "signatureIndex", 999 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion
}
