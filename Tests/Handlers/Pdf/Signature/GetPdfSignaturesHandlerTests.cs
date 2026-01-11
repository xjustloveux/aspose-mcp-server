using Aspose.Pdf;
using Aspose.Pdf.Forms;
using AsposeMcpServer.Handlers.Pdf.Signature;
using AsposeMcpServer.Tests.Helpers;

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

        var result = _handler.Execute(context, parameters);

        Assert.Contains("\"count\": 0", result);
        Assert.Contains("No signatures found", result);
    }

    [SkippableFact]
    public void Execute_ReturnsSignaturesList()
    {
        SkipInEvaluationMode(AsposeLibraryType.Pdf);
        var document = CreatePdfWithSignatureField();
        var context = CreateContext(document);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("count", result);
        Assert.Contains("items", result);
    }

    #endregion
}
