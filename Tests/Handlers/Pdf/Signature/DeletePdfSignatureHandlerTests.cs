using Aspose.Pdf;
using Aspose.Pdf.Forms;
using AsposeMcpServer.Handlers.Pdf.Signature;
using AsposeMcpServer.Results.Common;
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
        var context = CreateContext(document);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "signatureName", "TestSig" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("deleted", result.Message, StringComparison.OrdinalIgnoreCase);
        AssertModified(context);
    }

    [SkippableFact]
    public void Execute_DeletesSignatureByIndex()
    {
        SkipInEvaluationMode(AsposeLibraryType.Pdf);
        var document = CreatePdfWithSignatureFieldPersisted("TestSig");
        var context = CreateContext(document);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "signatureIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("deleted", result.Message, StringComparison.OrdinalIgnoreCase);
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
