using AsposeMcpServer.Handlers.Word.DigitalSignature;
using AsposeMcpServer.Results.Word.DigitalSignature;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Word.DigitalSignature;

/// <summary>
///     Tests for VerifyWordDigitalSignatureHandler.
/// </summary>
public class VerifyWordDigitalSignatureHandlerTests : WordHandlerTestBase
{
    private readonly VerifyWordDigitalSignatureHandler _handler = new();

    [Fact]
    public void Operation_ShouldBeVerify()
    {
        Assert.Equal("verify", _handler.Operation);
    }

    [Fact]
    public void Execute_WithUnsignedDocument_ShouldReturnNoSignatures()
    {
        var doc = CreateEmptyDocument();
        var docPath = Path.Combine(TestDir, "test_verify_unsigned.docx");
        doc.Save(docPath);

        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", docPath }
        });

        var result = _handler.Execute(context, parameters);

        var verifyResult = Assert.IsType<VerifySignaturesResult>(result);
        Assert.Equal(0, verifyResult.TotalCount);
        Assert.Equal(0, verifyResult.ValidCount);
        Assert.False(verifyResult.AllValid);
        Assert.Contains("No digital signatures", verifyResult.Message);
    }

    [Fact]
    public void Execute_WithMissingPath_ShouldThrowArgumentException()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }
}
