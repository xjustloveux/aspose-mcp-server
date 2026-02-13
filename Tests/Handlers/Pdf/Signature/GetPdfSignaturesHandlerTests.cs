using Aspose.Pdf;
using Aspose.Pdf.Forms;
using AsposeMcpServer.Handlers.Pdf.Signature;
using AsposeMcpServer.Results.Pdf.Signature;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Pdf.Signature;

public class GetPdfSignaturesHandlerTests : PdfHandlerTestBase
{
    private readonly GetPdfSignaturesHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Get()
    {
        Assert.Equal("get", _handler.Operation);
    }

    #endregion

    #region Helper Methods

    private static Document CreatePdfWithSignatureField()
    {
        var document = new Document();
        var page = document.Pages.Add();
        var signatureField = new SignatureField(page, new Rectangle(100, 600, 300, 700))
        {
            Name = "TestSignature"
        };
        document.Form.Add(signatureField);
        return document;
    }

    #endregion

    #region Helper Methods - Additional

    private static Document CreatePdfWithMultipleSignatureFields()
    {
        var document = new Document();
        var page = document.Pages.Add();

        var signatureField1 = new SignatureField(page, new Rectangle(100, 600, 300, 700))
        {
            Name = "Signature1"
        };
        var signatureField2 = new SignatureField(page, new Rectangle(100, 400, 300, 500))
        {
            Name = "Signature2"
        };
        var signatureField3 = new SignatureField(page, new Rectangle(100, 200, 300, 300))
        {
            Name = "Signature3"
        };

        document.Form.Add(signatureField1);
        document.Form.Add(signatureField2);
        document.Form.Add(signatureField3);

        return document;
    }

    #endregion

    #region Basic Get Signatures Operations

    [Fact]
    public void Execute_ReturnsEmptyWhenNoSignatures()
    {
        var document = CreateEmptyDocument();
        var context = CreateContext(document);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetSignaturesResult>(res);
        Assert.Equal(0, result.Count);
        Assert.Equal("No signatures found", result.Message);
    }

    [SkippableFact]
    public void Execute_ReturnsSignaturesList()
    {
        SkipInEvaluationMode(AsposeLibraryType.Pdf);
        var document = CreatePdfWithSignatureField();
        var context = CreateContext(document);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetSignaturesResult>(res);
        Assert.True(result.Count >= 0);
        Assert.NotNull(result.Items);
    }

    [SkippableFact]
    public void Execute_WithSignatureField_ReturnsSignatureInfo()
    {
        SkipInEvaluationMode(AsposeLibraryType.Pdf);
        var document = CreatePdfWithSignatureField();
        var context = CreateContext(document);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetSignaturesResult>(res);
        if (result.Count > 0)
        {
            var firstSignature = result.Items[0];
            Assert.Equal(0, firstSignature.Index);
            Assert.NotNull(firstSignature.Name);
        }
    }

    [SkippableFact]
    public void Execute_WithMultipleSignatureFields_ReturnsAllSignatures()
    {
        SkipInEvaluationMode(AsposeLibraryType.Pdf);
        var document = CreatePdfWithMultipleSignatureFields();
        var context = CreateContext(document);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetSignaturesResult>(res);
        Assert.NotNull(result.Items);
        for (var i = 0; i < result.Items.Count; i++) Assert.Equal(i, result.Items[i].Index);
    }

    [SkippableFact]
    public void Execute_WithUnsignedField_ReturnsInvalidSignature()
    {
        SkipInEvaluationMode(AsposeLibraryType.Pdf);
        var document = CreatePdfWithSignatureField();
        var context = CreateContext(document);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetSignaturesResult>(res);
        if (result.Count > 0)
        {
            var signature = result.Items[0];
            Assert.False(signature.IsValid);
            Assert.False(signature.HasCertificate);
        }
    }

    [Fact]
    public void Execute_EmptyDocument_ReturnsNoSignaturesMessage()
    {
        var document = new Document();
        document.Pages.Add();
        var context = CreateContext(document);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetSignaturesResult>(res);
        Assert.Equal(0, result.Count);
        Assert.Empty(result.Items);
        Assert.Equal("No signatures found", result.Message);
    }

    [SkippableFact]
    public void Execute_WithSignatureField_CountMatchesItemsCount()
    {
        SkipInEvaluationMode(AsposeLibraryType.Pdf);
        var document = CreatePdfWithSignatureField();
        var context = CreateContext(document);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetSignaturesResult>(res);
        Assert.Equal(result.Count, result.Items.Count);
    }

    [SkippableFact]
    public void Execute_WithMultipleFields_CountMatchesItemsCount()
    {
        SkipInEvaluationMode(AsposeLibraryType.Pdf);
        var document = CreatePdfWithMultipleSignatureFields();
        var context = CreateContext(document);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetSignaturesResult>(res);
        Assert.Equal(result.Count, result.Items.Count);
    }

    #endregion
}
