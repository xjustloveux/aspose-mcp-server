using Aspose.Pdf;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Handlers.Pdf.Info;
using AsposeMcpServer.Results.Pdf.Info;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Pdf.Info;

public class GetPdfStatisticsHandlerTests : PdfHandlerTestBase
{
    private readonly GetPdfStatisticsHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_GetStatistics()
    {
        Assert.Equal("get_statistics", _handler.Operation);
    }

    #endregion

    #region Basic Statistics Operations

    [Fact]
    public void Execute_WithSession_ReturnsStatistics()
    {
        var doc = CreateDocumentWithPages(3);
        var context = CreateContextWithSession(doc, "test-session");
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetPdfStatisticsResult>(res);
        Assert.Equal(3, result.TotalPages);
    }

    [Fact]
    public void Execute_WithSession_ReturnsEncryptedStatus()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContextWithSession(doc, "test-session");
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetPdfStatisticsResult>(res);
        Assert.False(result.IsEncrypted);
    }

    [Fact]
    public void Execute_WithSession_ReturnsLinearizedStatus()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContextWithSession(doc, "test-session");
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetPdfStatisticsResult>(res);
        Assert.False(result.IsLinearized);
    }

    [Fact]
    public void Execute_WithSession_ReturnsAnnotationsCount()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContextWithSession(doc, "test-session");
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetPdfStatisticsResult>(res);
        Assert.True(result.TotalAnnotations >= 0);
    }

    [Fact]
    public void Execute_WithSession_ReturnsParagraphsCount()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContextWithSession(doc, "test-session");
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetPdfStatisticsResult>(res);
        Assert.True(result.TotalParagraphs >= 0);
    }

    [Fact]
    public void Execute_WithSession_ReturnsBookmarksCount()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContextWithSession(doc, "test-session");
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetPdfStatisticsResult>(res);
        Assert.True(result.Bookmarks >= 0);
    }

    [Fact]
    public void Execute_WithSession_ReturnsFormFieldsCount()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContextWithSession(doc, "test-session");
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetPdfStatisticsResult>(res);
        Assert.True(result.FormFields >= 0);
    }

    #endregion

    #region File Path Mode

    [Fact]
    public void Execute_WithSourcePath_ReturnsFileSize()
    {
        var pdfPath = CreateTempPdfFile();
        var doc = new Document(pdfPath);
        var context = CreateContextWithSourcePath(doc, pdfPath);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetPdfStatisticsResult>(res);
        Assert.NotNull(result.FileSizeBytes);
        Assert.NotNull(result.FileSizeKb);
    }

    [Fact]
    public void Execute_WithoutSessionOrPath_ThrowsArgumentException()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("path is required", ex.Message);
    }

    #endregion

    #region Helper Methods

    private string CreateTempPdfFile()
    {
        var tempPath = Path.Combine(TestDir, $"test_{Guid.NewGuid()}.pdf");
        var doc = new Document();
        doc.Pages.Add();
        doc.Save(tempPath);
        return tempPath;
    }

    private static OperationContext<Document> CreateContextWithSession(Document doc, string sessionId)
    {
        return new OperationContext<Document>
        {
            Document = doc,
            SessionId = sessionId
        };
    }

    private static OperationContext<Document> CreateContextWithSourcePath(Document doc, string sourcePath)
    {
        return new OperationContext<Document>
        {
            Document = doc,
            SourcePath = sourcePath
        };
    }

    #endregion
}
