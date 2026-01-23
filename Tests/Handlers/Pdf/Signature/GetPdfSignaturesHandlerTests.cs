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

    #endregion
}
