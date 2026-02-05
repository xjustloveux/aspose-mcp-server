using AsposeMcpServer.Handlers.Word.DigitalSignature;
using AsposeMcpServer.Results.Word.DigitalSignature;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Word.DigitalSignature;

/// <summary>
///     Tests for ListWordDigitalSignatureHandler.
/// </summary>
public class ListWordDigitalSignatureHandlerTests : WordHandlerTestBase
{
    private readonly ListWordDigitalSignatureHandler _handler = new();

    [Fact]
    public void Operation_ShouldBeList()
    {
        Assert.Equal("list", _handler.Operation);
    }

    [Fact]
    public void Execute_WithUnsignedDocument_ShouldReturnEmptyList()
    {
        var doc = CreateEmptyDocument();
        var docPath = Path.Combine(TestDir, "test_list_unsigned.docx");
        doc.Save(docPath);

        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", docPath }
        });

        var result = _handler.Execute(context, parameters);

        var listResult = Assert.IsType<GetSignaturesResult>(result);
        Assert.Equal(0, listResult.Count);
        Assert.Empty(listResult.Signatures);
        Assert.Contains("0 digital signature", listResult.Message);
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
