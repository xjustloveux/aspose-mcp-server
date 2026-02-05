using Aspose.Pdf;
using Aspose.Pdf.Text;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Results.Pdf.Toc;
using AsposeMcpServer.Tests.Infrastructure;
using AsposeMcpServer.Tools.Pdf;

namespace AsposeMcpServer.Tests.Tools.Pdf;

/// <summary>
///     Integration tests for PdfTocTool.
///     Focuses on session management, file I/O, and operation routing.
///     Detailed parameter validation and business logic tests are in Handler tests.
/// </summary>
public class PdfTocToolTests : PdfTestBase
{
    private readonly PdfTocTool _tool;

    public PdfTocToolTests()
    {
        _tool = new PdfTocTool(SessionManager);
    }

    private string CreateTestPdf(string fileName, int pageCount = 3)
    {
        var filePath = CreateTestFilePath(fileName);
        using var document = new Document();
        for (var i = 0; i < pageCount; i++)
        {
            var page = document.Pages.Add();
            page.Paragraphs.Add(new TextFragment($"Page {i + 1} content"));
        }

        document.Save(filePath);
        return filePath;
    }

    #region File I/O Smoke Tests

    [Fact]
    public void Execute_Generate_ReturnsResult()
    {
        var pdfPath = CreateTestPdf("test_generate.pdf");
        var outputPath = CreateTestFilePath("test_generate_output.pdf");
        var result = _tool.Execute("generate", pdfPath, outputPath: outputPath);
        var data = GetResultData<SuccessResult>(result);
        Assert.Contains("Generated TOC", data.Message);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void Execute_Get_ReturnsResult()
    {
        var pdfPath = CreateTestPdf("test_get.pdf");
        var result = _tool.Execute("get", pdfPath);
        var data = GetResultData<GetTocPdfResult>(result);
        Assert.NotNull(data.Entries);
    }

    [Fact]
    public void Execute_Remove_ReturnsResult()
    {
        var pdfPath = CreateTestPdf("test_remove.pdf");
        var outputPath = CreateTestFilePath("test_remove_output.pdf");
        var result = _tool.Execute("remove", pdfPath, outputPath: outputPath);
        var data = GetResultData<SuccessResult>(result);
        Assert.NotNull(data.Message);
    }

    [Fact]
    public void Execute_Generate_WithCustomTitle_ReturnsResult()
    {
        var pdfPath = CreateTestPdf("test_generate_title.pdf");
        var outputPath = CreateTestFilePath("test_generate_title_output.pdf");
        var result = _tool.Execute("generate", pdfPath, outputPath: outputPath,
            title: "Custom Contents");
        var data = GetResultData<SuccessResult>(result);
        Assert.Contains("Generated TOC", data.Message);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void Execute_Generate_WithDepthAndTocPage_ReturnsResult()
    {
        var pdfPath = CreateTestPdf("test_generate_opts.pdf");
        var outputPath = CreateTestFilePath("test_generate_opts_output.pdf");
        var result = _tool.Execute("generate", pdfPath, outputPath: outputPath,
            depth: 2, tocPage: 1);
        var data = GetResultData<SuccessResult>(result);
        Assert.Contains("Generated TOC", data.Message);
    }

    #endregion

    #region Operation Routing

    [Theory]
    [InlineData("GENERATE")]
    [InlineData("Generate")]
    [InlineData("generate")]
    public void Operation_ShouldBeCaseInsensitive(string operation)
    {
        var pdfPath = CreateTestPdf($"test_case_{operation}.pdf");
        var outputPath = CreateTestFilePath($"test_case_{operation}_output.pdf");
        var result = _tool.Execute(operation, pdfPath, outputPath: outputPath);
        var data = GetResultData<SuccessResult>(result);
        Assert.Contains("Generated TOC", data.Message);
    }

    [Fact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        var pdfPath = CreateTestPdf("test_unknown_op.pdf");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("unknown", pdfPath));
        Assert.Contains("Unknown operation", ex.Message);
    }

    [Fact]
    public void Execute_WithNoPathOrSessionId_ShouldThrowException()
    {
        Assert.ThrowsAny<Exception>(() => _tool.Execute("get"));
    }

    #endregion

    #region Session Management

    [Fact]
    public void Get_WithSessionId_ShouldGetFromSession()
    {
        var pdfPath = CreateTestPdf("test_session_get.pdf");
        var sessionId = OpenSession(pdfPath);
        var result = _tool.Execute("get", sessionId: sessionId);
        var data = GetResultData<GetTocPdfResult>(result);
        Assert.NotNull(data.Entries);
    }

    [Fact]
    public void Generate_WithSessionId_ShouldGenerateInSession()
    {
        var pdfPath = CreateTestPdf("test_session_generate.pdf");
        var sessionId = OpenSession(pdfPath);
        var result = _tool.Execute("generate", sessionId: sessionId);
        var data = GetResultData<SuccessResult>(result);
        Assert.Contains("Generated TOC", data.Message);
        var output = GetResultOutput<SuccessResult>(result);
        Assert.True(output.IsSession);
    }

    [Fact]
    public void Remove_WithSessionId_ShouldRemoveFromSession()
    {
        var pdfPath = CreateTestPdf("test_session_remove.pdf");
        var sessionId = OpenSession(pdfPath);
        var result = _tool.Execute("remove", sessionId: sessionId);
        var data = GetResultData<SuccessResult>(result);
        Assert.NotNull(data.Message);
        var output = GetResultOutput<SuccessResult>(result);
        Assert.True(output.IsSession);
    }

    [Fact]
    public void Execute_WithInvalidSessionId_ShouldThrowKeyNotFoundException()
    {
        Assert.Throws<KeyNotFoundException>(() =>
            _tool.Execute("get", sessionId: "invalid_session"));
    }

    #endregion
}
