using AsposeMcpServer.Handlers.Pdf.FileOperations;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

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

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("compressed", result.Message, StringComparison.OrdinalIgnoreCase);
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

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("compressed", result.Message, StringComparison.OrdinalIgnoreCase);
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

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("compressed", result.Message, StringComparison.OrdinalIgnoreCase);
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

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("compressed", result.Message, StringComparison.OrdinalIgnoreCase);
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

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("compressed", result.Message, StringComparison.OrdinalIgnoreCase);
        AssertModified(context);
    }

    #endregion
}
