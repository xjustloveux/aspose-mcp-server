using AsposeMcpServer.Handlers.Pdf.FileOperations;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Pdf.FileOperations;

public class EncryptPdfFileHandlerTests : PdfHandlerTestBase
{
    private readonly EncryptPdfFileHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Encrypt()
    {
        Assert.Equal("encrypt", _handler.Operation);
    }

    #endregion

    #region Basic Encrypt Operations

    [Fact]
    public void Execute_EncryptsPdf()
    {
        var doc = CreateDocumentWithText("Confidential content");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "userPassword", "user123" },
            { "ownerPassword", "owner456" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("encrypted", result.Message, StringComparison.OrdinalIgnoreCase);
        AssertModified(context);
        Assert.True(doc.IsEncrypted, "Document should be encrypted after operation");
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutUserPassword_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithText("Test content");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "ownerPassword", "owner456" }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithoutOwnerPassword_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithText("Test content");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "userPassword", "user123" }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithoutAnyPassword_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithText("Test content");
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion
}
