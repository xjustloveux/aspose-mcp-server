using Aspose.Pdf;
using AsposeMcpServer.Handlers.Pdf.FileOperations;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Pdf.FileOperations;

public class CreatePdfFileHandlerTests : PdfHandlerTestBase
{
    private readonly CreatePdfFileHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Create()
    {
        Assert.Equal("create", _handler.Operation);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutOutputPath_ThrowsArgumentException()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion

    #region Basic Create Operations

    [Fact]
    public void Execute_CreatesPdfDocument()
    {
        var outputPath = Path.Combine(TestDir, "test.pdf");
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "outputPath", outputPath }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("created", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.True(File.Exists(outputPath));
        var fileInfo = new FileInfo(outputPath);
        Assert.True(fileInfo.Length > 0, "Created PDF should have content");
    }

    [Fact]
    public void Execute_CreatesValidPdf()
    {
        var outputPath = Path.Combine(TestDir, "valid.pdf");
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "outputPath", outputPath }
        });

        _handler.Execute(context, parameters);

        using var createdDoc = new Document(outputPath);
        Assert.NotNull(createdDoc);
        Assert.Single(createdDoc.Pages);
    }

    #endregion
}
