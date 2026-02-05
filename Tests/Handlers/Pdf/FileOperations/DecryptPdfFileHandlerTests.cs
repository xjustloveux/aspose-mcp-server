using Aspose.Pdf;
using AsposeMcpServer.Handlers.Pdf.FileOperations;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Pdf.FileOperations;

/// <summary>
///     Unit tests for DecryptPdfFileHandler class.
/// </summary>
public class DecryptPdfFileHandlerTests : PdfHandlerTestBase
{
    private readonly DecryptPdfFileHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Decrypt()
    {
        Assert.Equal("decrypt", _handler.Operation);
    }

    #endregion

    #region Basic Decrypt Operations

    [Fact]
    public void Execute_DecryptsEncryptedPdf()
    {
        var doc = CreateDocumentWithText("Confidential content");
        doc.Encrypt("user123", "owner456",
            Permissions.PrintDocument | Permissions.ModifyContent,
            CryptoAlgorithm.AESx256);
        Assert.True(doc.IsEncrypted, "Document should be encrypted before decrypt");

        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);
        Assert.Contains("decrypt", result.Message, StringComparison.OrdinalIgnoreCase);
        AssertModified(context);
        Assert.False(doc.IsEncrypted, "Document should not be encrypted after decrypt");
    }

    [Fact]
    public void Execute_OnUnencryptedPdf_StillSucceeds()
    {
        var doc = CreateDocumentWithText("Normal content");
        Assert.False(doc.IsEncrypted);

        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);
        Assert.Contains("decrypt", result.Message, StringComparison.OrdinalIgnoreCase);
        AssertModified(context);
    }

    #endregion
}
