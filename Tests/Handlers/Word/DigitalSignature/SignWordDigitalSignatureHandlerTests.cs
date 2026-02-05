using AsposeMcpServer.Handlers.Word.DigitalSignature;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Word.DigitalSignature;

/// <summary>
///     Tests for SignWordDigitalSignatureHandler.
/// </summary>
public class SignWordDigitalSignatureHandlerTests : WordHandlerTestBase
{
    private readonly SignWordDigitalSignatureHandler _handler = new();

    [Fact]
    public void Operation_ShouldBeSign()
    {
        Assert.Equal("sign", _handler.Operation);
    }

    [Fact]
    public void Execute_WithMissingPath_ShouldThrowArgumentException()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "outputPath", "output.docx" },
            { "certificatePath", "cert.pfx" },
            { "certificatePassword", "pass" }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithMissingOutputPath_ShouldThrowArgumentException()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", "input.docx" },
            { "certificatePath", "cert.pfx" },
            { "certificatePassword", "pass" }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithMissingCertificatePath_ShouldThrowArgumentException()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", "input.docx" },
            { "outputPath", "output.docx" },
            { "certificatePassword", "pass" }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithNonExistentCertificate_ShouldThrowFileNotFoundException()
    {
        var doc = CreateEmptyDocument();
        var docPath = Path.Combine(TestDir, "test_sign.docx");
        doc.Save(docPath);

        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", docPath },
            { "outputPath", Path.Combine(TestDir, "signed.docx") },
            { "certificatePath", Path.Combine(TestDir, "nonexistent.pfx") },
            { "certificatePassword", "pass" }
        });

        Assert.Throws<FileNotFoundException>(() => _handler.Execute(context, parameters));
    }
}
