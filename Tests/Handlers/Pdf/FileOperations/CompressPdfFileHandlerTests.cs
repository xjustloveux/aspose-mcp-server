using AsposeMcpServer.Handlers.Pdf.FileOperations;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Pdf.FileOperations;

public class CompressPdfFileHandlerTests : PdfHandlerTestBase
{
    private readonly CompressPdfFileHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Compress()
    {
        Assert.Equal("compress", _handler.Operation);
    }

    #endregion

    #region Basic Compress Operations

    [Fact]
    public void Execute_CompressesPdf()
    {
        var doc = CreateDocumentWithText("Test content");
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("compressed", result.ToLower());
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithCompressImages_CompressesPdf()
    {
        var doc = CreateDocumentWithText("Test content");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "compressImages", true }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("compressed", result.ToLower());
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithCompressFonts_CompressesPdf()
    {
        var doc = CreateDocumentWithText("Test content");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "compressFonts", true }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("compressed", result.ToLower());
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithRemoveUnusedObjects_CompressesPdf()
    {
        var doc = CreateDocumentWithText("Test content");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "removeUnusedObjects", true }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("compressed", result.ToLower());
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithAllOptionsFalse_CompressesPdf()
    {
        var doc = CreateDocumentWithText("Test content");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "compressImages", false },
            { "compressFonts", false },
            { "removeUnusedObjects", false }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("compressed", result.ToLower());
        AssertModified(context);
    }

    #endregion
}
