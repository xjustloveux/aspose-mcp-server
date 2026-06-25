using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using Aspose.Pdf;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Handlers.Pdf.Signature;
using AsposeMcpServer.Tests.Infrastructure;

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

    #region Helper Methods

    private string CreateSelfSignedPfx(string password)
    {
        using var rsa = RSA.Create(2048);
        var request = new CertificateRequest("CN=AsposeMcpTest", rsa, HashAlgorithmName.SHA256,
            RSASignaturePadding.Pkcs1);
        request.CertificateExtensions.Add(new X509KeyUsageExtension(
            X509KeyUsageFlags.DigitalSignature | X509KeyUsageFlags.NonRepudiation, true));
        using var certificate = request.CreateSelfSigned(
            DateTimeOffset.UtcNow.AddDays(-1), DateTimeOffset.UtcNow.AddYears(1));
        var pfxBytes = certificate.Export(X509ContentType.Pfx, password);
        return CreateTempFile(".pfx", pfxBytes);
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

    #region Successful Sign (Disposal + File Output)

    [SkippableFact]
    public void Execute_SessionMode_DoesNotDisposeTheDocument()
    {
        SkipInEvaluationMode(AsposeLibraryType.Pdf);
        // In session mode the Document is reused across operations. sign binds a PdfFileSignature to it and
        // takes the MarkModified branch (no direct Save), so the handler must NOT dispose the session-owned
        // Document when it returns — otherwise every later operation on that session throws ObjectDisposed.
        const string password = "test-password";
        var certPath = CreateSelfSignedPfx(password);
        var document = CreateEmptyDocument();
        var context = new OperationContext<Document>
        {
            Document = document,
            SessionId = "test-session"
        };
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "certificatePath", certPath },
            { "password", password }
        });

        _handler.Execute(context, parameters);

        var ex = Record.Exception(() => _ = document.Pages.Count);
        Assert.Null(ex);
    }

    [SkippableFact]
    public void Execute_FileMode_WritesSignedFileAndLeavesItReadable()
    {
        SkipInEvaluationMode(AsposeLibraryType.Pdf);
        // File mode persists via pdfSign.Save(outputPath). Dropping the 'using' on PdfFileSignature must not
        // leave the output file locked or unwritten — the saved PDF must exist and be re-openable afterwards.
        // PdfFileSignature.Save re-reads the document's backing stream, so the source must be a real PDF on
        // disk (an in-memory `new Document()` has no header for Save to re-read); this mirrors real file mode.
        const string password = "test-password";
        var certPath = CreateSelfSignedPfx(password);
        var sourcePath = Path.Combine(TestDir, $"source_{Guid.NewGuid()}.pdf");
        using (var seed = CreateEmptyDocument())
        {
            seed.Save(sourcePath);
        }

        using var document = new Document(sourcePath);
        var outputPath = Path.Combine(TestDir, $"signed_{Guid.NewGuid()}.pdf");
        var context = CreateContext(document, outputPath);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "certificatePath", certPath },
            { "password", password }
        });

        _handler.Execute(context, parameters);

        Assert.True(File.Exists(outputPath));
        using var signed = new Document(outputPath);
        Assert.True(signed.Pages.Count >= 1);
    }

    #endregion
}
