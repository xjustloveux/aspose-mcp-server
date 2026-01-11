using AsposeMcpServer.Handlers.Pdf.Signature;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Pdf.Signature;

public class SignPdfHandlerTests : PdfHandlerTestBase
{
    private readonly SignPdfHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Sign()
    {
        Assert.Equal("sign", _handler.Operation);
    }

    #endregion

    #region Basic Sign Operations

    [Fact]
    public void Execute_WithNonExistentCertificate_ThrowsFileNotFoundException()
    {
        var document = CreateEmptyDocument();
        var context = CreateContext(document);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "certificatePath", "C:/nonexistent/certificate.pfx" },
            { "password", "testpassword" }
        });

        Assert.Throws<FileNotFoundException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithInvalidPageIndex_ThrowsArgumentException()
    {
        var document = CreateEmptyDocument();
        var context = CreateContext(document);
        var tempCert = CreateTempFile(".pfx", Array.Empty<byte>());

        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "certificatePath", tempCert },
            { "password", "test" },
            { "pageIndex", 999 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithPageIndexZero_ThrowsArgumentException()
    {
        var document = CreateEmptyDocument();
        var context = CreateContext(document);
        var tempCert = CreateTempFile(".pfx", Array.Empty<byte>());

        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "certificatePath", tempCert },
            { "password", "test" },
            { "pageIndex", 0 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion
}
