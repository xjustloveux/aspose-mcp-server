using Aspose.Pdf;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Results.Pdf.Properties;
using AsposeMcpServer.Tests.Infrastructure;
using AsposeMcpServer.Tools.Pdf;

namespace AsposeMcpServer.Tests.Tools.Pdf;

/// <summary>
///     Integration tests for PdfPropertiesTool.
///     Focuses on session management, file I/O, and operation routing.
///     Detailed parameter validation and business logic tests are in Handler tests.
/// </summary>
public class PdfPropertiesToolTests : PdfTestBase
{
    private readonly PdfPropertiesTool _tool;

    public PdfPropertiesToolTests()
    {
        _tool = new PdfPropertiesTool(SessionManager);
    }

    private string CreateTestPdf(string fileName)
    {
        var filePath = CreateTestFilePath(fileName);
        using var document = new Document();
        document.Pages.Add();
        document.Save(filePath);
        return filePath;
    }

    #region File I/O Smoke Tests

    [Fact]
    public void Get_ShouldReturnProperties()
    {
        var pdfPath = CreateTestPdf("test_get.pdf");
        var result = _tool.Execute("get", pdfPath);
        var data = GetResultData<GetPropertiesPdfResult>(result);
        Assert.NotNull(data);
        Assert.True(data.TotalPages > 0);
    }

    [Fact]
    public void Set_WithTitleAndAuthor_ShouldSetProperties()
    {
        var pdfPath = CreateTestPdf("test_set.pdf");
        var outputPath = CreateTestFilePath("test_set_output.pdf");
        var result = _tool.Execute("set", pdfPath, outputPath: outputPath,
            title: "Test PDF", author: "Test Author", subject: "Test Subject");
        Assert.True(File.Exists(outputPath));
        var data = GetResultData<SuccessResult>(result);
        Assert.StartsWith("Document properties updated", data.Message);
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
        var data = GetResultData<GetPropertiesPdfResult>(result);
        Assert.NotNull(data);
        Assert.True(data.TotalPages > 0);
    }

    [Fact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        var pdfPath = CreateTestPdf("test_unknown_op.pdf");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("unknown", pdfPath));
        Assert.Contains("Unknown operation", ex.Message);
    }

    #endregion

    #region Session Management

    [Fact]
    public void Get_WithSessionId_ShouldGetFromSession()
    {
        var pdfPath = CreateTestPdf("test_session_get.pdf");
        var sessionId = OpenSession(pdfPath);
        var result = _tool.Execute("get", sessionId: sessionId);
        var data = GetResultData<GetPropertiesPdfResult>(result);
        Assert.NotNull(data);
        Assert.True(data.TotalPages > 0);
        var output = GetResultOutput<GetPropertiesPdfResult>(result);
        Assert.True(output.IsSession);
    }

    [Fact]
    public void Set_WithSessionId_ShouldSetInSession()
    {
        var pdfPath = CreateTestPdf("test_session_set.pdf");
        var sessionId = OpenSession(pdfPath);
        var result = _tool.Execute("set", sessionId: sessionId,
            title: "Session Title", author: "Session Author");
        var data = GetResultData<SuccessResult>(result);
        Assert.StartsWith("Document properties updated", data.Message);
        var output = GetResultOutput<SuccessResult>(result);
        Assert.True(output.IsSession);
        var doc = SessionManager.GetDocument<Document>(sessionId);
        Assert.Equal("Session Title", doc.Info.Title);
        Assert.Equal("Session Author", doc.Info.Author);
    }

    [Fact]
    public void Execute_WithInvalidSessionId_ShouldThrowKeyNotFoundException()
    {
        Assert.Throws<KeyNotFoundException>(() => _tool.Execute("get", sessionId: "invalid_session"));
    }

    [Fact]
    public void Execute_WithBothPathAndSessionId_ShouldPreferSessionId()
    {
        _ = CreateTestPdf("test_path_props.pdf");
        var pdfPath2 = CreateTestPdf("test_session_props.pdf");
        var sessionId = OpenSession(pdfPath2);
        _tool.Execute("set", sessionId: sessionId, title: "Session Doc Title");
        var doc = SessionManager.GetDocument<Document>(sessionId);
        Assert.Equal("Session Doc Title", doc.Info.Title);
    }

    #endregion
}
