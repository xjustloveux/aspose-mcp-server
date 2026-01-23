using Aspose.Pdf;
using AsposeMcpServer.Results.Pdf.Signature;
using AsposeMcpServer.Tests.Infrastructure;
using AsposeMcpServer.Tools.Pdf;

namespace AsposeMcpServer.Tests.Tools.Pdf;

/// <summary>
///     Integration tests for PdfSignatureTool.
///     Focuses on session management, file I/O, and operation routing.
///     Detailed parameter validation and business logic tests are in Handler tests.
/// </summary>
public class PdfSignatureToolTests : PdfTestBase
{
    private readonly PdfSignatureTool _tool;

    public PdfSignatureToolTests()
    {
        _tool = new PdfSignatureTool(SessionManager);
    }

    private string CreateTestPdf(string fileName)
    {
        var filePath = CreateTestFilePath(fileName);
        using var document = new Document();
        document.Pages.Add();
        document.Save(filePath);
        return filePath;
    }

    private string CreateMultiPagePdf(string fileName, int pageCount)
    {
        var filePath = CreateTestFilePath(fileName);
        using var document = new Document();
        for (var i = 0; i < pageCount; i++)
            document.Pages.Add();
        document.Save(filePath);
        return filePath;
    }

    #region File I/O Smoke Tests

    [Fact]
    public void Get_WithNoSignatures_ShouldReturnEmptyResult()
    {
        var pdfPath = CreateTestPdf("test_get_empty.pdf");
        var result = _tool.Execute("get", pdfPath);
        var data = GetResultData<GetSignaturesResult>(result);

        Assert.Equal(0, data.Count);
        Assert.Equal("No signatures found", data.Message);
    }

    [Fact]
    public void Get_ShouldReturnValidJson()
    {
        var pdfPath = CreateTestPdf("test_get_json.pdf");
        var result = _tool.Execute("get", pdfPath);
        var data = GetResultData<GetSignaturesResult>(result);
        Assert.True(data.Count >= 0);
        Assert.NotNull(data.Items);
    }

    #endregion

    #region Operation Routing

    [Theory]
    [InlineData("GET")]
    [InlineData("Get")]
    [InlineData("get")]
    public void Operation_ShouldBeCaseInsensitive(string operation)
    {
        var pdfPath = CreateTestPdf($"test_case_{operation}.pdf");
        var result = _tool.Execute(operation, pdfPath);
        var data = GetResultData<GetSignaturesResult>(result);
        Assert.Equal(0, data.Count);
    }

    [Fact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        var pdfPath = CreateTestPdf("test_unknown_op.pdf");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("unknown", pdfPath));
        Assert.StartsWith("Unknown operation: unknown", ex.Message);
    }

    [Fact]
    public void Execute_WithNoPathOrSessionId_ShouldThrowException()
    {
        Assert.ThrowsAny<Exception>(() => _tool.Execute("get"));
    }

    #endregion

    #region Session Management

    [Fact]
    public void Get_WithSessionId_ShouldReturnResult()
    {
        var pdfPath = CreateTestPdf("test_session_get.pdf");
        var sessionId = OpenSession(pdfPath);
        var result = _tool.Execute("get", sessionId: sessionId);
        var data = GetResultData<GetSignaturesResult>(result);
        Assert.Equal(0, data.Count);
        var document = SessionManager.GetDocument<Document>(sessionId);
        Assert.NotNull(document);
    }

    [Fact]
    public void Execute_WithInvalidSessionId_ShouldThrowKeyNotFoundException()
    {
        Assert.Throws<KeyNotFoundException>(() => _tool.Execute("get", sessionId: "invalid_session"));
    }

    [Fact]
    public void Execute_WithBothPathAndSessionId_ShouldPreferSessionId()
    {
        var pdfPath1 = CreateTestPdf("test_path_sig.pdf");
        var pdfPath2 = CreateMultiPagePdf("test_session_sig.pdf", 2);
        var sessionId = OpenSession(pdfPath2);
        var result = _tool.Execute("get", pdfPath1, sessionId);
        var data = GetResultData<GetSignaturesResult>(result);
        Assert.Equal(0, data.Count);
    }

    #endregion
}
