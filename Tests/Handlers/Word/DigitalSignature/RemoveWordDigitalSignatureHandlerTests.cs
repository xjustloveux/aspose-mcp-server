using AsposeMcpServer.Handlers.Word.DigitalSignature;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Word.DigitalSignature;

/// <summary>
///     Tests for RemoveWordDigitalSignatureHandler.
/// </summary>
public class RemoveWordDigitalSignatureHandlerTests : WordHandlerTestBase
{
    private readonly RemoveWordDigitalSignatureHandler _handler = new();

    [Fact]
    public void Operation_ShouldBeRemove()
    {
        Assert.Equal("remove", _handler.Operation);
    }

    [Fact]
    public void Execute_WithUnsignedDocument_ShouldSucceed()
    {
        var doc = CreateDocumentWithText("Test content");
        var docPath = Path.Combine(TestDir, "test_remove_unsigned.docx");
        doc.Save(docPath);

        var outputPath = Path.Combine(TestDir, "test_remove_output.docx");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", docPath },
            { "outputPath", outputPath }
        });

        var result = _handler.Execute(context, parameters);

        Assert.NotNull(result);
        Assert.True(System.IO.File.Exists(outputPath));
    }

    [Fact]
    public void Execute_WithMissingPath_ShouldThrowArgumentException()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "outputPath", "output.docx" }
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
            { "path", "input.docx" }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }
}
