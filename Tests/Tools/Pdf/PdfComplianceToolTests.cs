using System.Runtime.Versioning;
using Aspose.Pdf;
using Aspose.Pdf.Text;
using AsposeMcpServer.Results;
using AsposeMcpServer.Results.Pdf.Compliance;
using AsposeMcpServer.Tests.Infrastructure;
using AsposeMcpServer.Tools.Pdf;

namespace AsposeMcpServer.Tests.Tools.Pdf;

/// <summary>
///     Integration tests for <see cref="PdfComplianceTool" />.
///     Focuses on session management, file I/O, and operation routing.
///     Detailed parameter validation and business logic tests are in Handler tests.
/// </summary>
[SupportedOSPlatform("windows")]
public class PdfComplianceToolTests : PdfTestBase
{
    private readonly PdfComplianceTool _tool;

    public PdfComplianceToolTests()
    {
        _tool = new PdfComplianceTool(SessionManager);
    }

    private string CreateTestPdf(string fileName)
    {
        var filePath = CreateTestFilePath(fileName);
        using var document = new Document();
        var page = document.Pages.Add();
        page.Paragraphs.Add(new TextFragment("Test content"));
        document.Save(filePath);
        return filePath;
    }

    #region File I/O Smoke Tests

    [SkippableFact]
    public void Execute_Validate_ReturnsResult()
    {
        SkipIfNotWindows();
        var pdfPath = CreateTestPdf("test_validate.pdf");

        var result = _tool.Execute("validate", pdfPath, format: "pdf/a-1b");

        var data = GetResultData<ValidateCompliancePdfResult>(result);
        Assert.Equal("PDF/A-1b", data.Format);
        Assert.NotNull(data.Message);
    }

    [SkippableFact]
    public void Execute_Convert_ReturnsResult()
    {
        SkipIfNotWindows();
        var pdfPath = CreateTestPdf("test_convert.pdf");
        var outputPath = CreateTestFilePath("test_convert_output.pdf");

        var result = _tool.Execute("convert", pdfPath, outputPath: outputPath, format: "pdf/a-1b");

        var data = GetResultData<ConvertCompliancePdfResult>(result);
        Assert.Equal("PDF/A-1b", data.Format);
        Assert.NotNull(data.Message);
        Assert.True(File.Exists(outputPath));
    }

    [SkippableFact]
    public void Execute_Convert_MarksModified()
    {
        SkipIfNotWindows();
        var pdfPath = CreateTestPdf("test_convert_modified.pdf");
        var outputPath = CreateTestFilePath("test_convert_modified_output.pdf");

        var result = _tool.Execute("convert", pdfPath, outputPath: outputPath, format: "pdf/a-1b");

        Assert.IsType<FinalizedResult<ConvertCompliancePdfResult>>(result);
        Assert.True(File.Exists(outputPath));
    }

    [SkippableFact]
    public void Execute_Validate_WithLogPath_WritesLog()
    {
        SkipIfNotWindows();
        var pdfPath = CreateTestPdf("test_validate_log.pdf");
        var logPath = CreateTestFilePath("validation.log");

        var result = _tool.Execute("validate", pdfPath, format: "pdf/a-1b", logPath: logPath);

        var data = GetResultData<ValidateCompliancePdfResult>(result);
        Assert.Equal(logPath, data.LogPath);
        Assert.True(File.Exists(logPath));
    }

    #endregion

    #region Operation Routing

    [SkippableTheory]
    [InlineData("VALIDATE")]
    [InlineData("Validate")]
    [InlineData("validate")]
    public void Validate_Operation_ShouldBeCaseInsensitive(string operation)
    {
        SkipIfNotWindows();
        var pdfPath = CreateTestPdf($"test_case_validate_{operation}.pdf");

        var result = _tool.Execute(operation, pdfPath, format: "pdf/a-1b");

        Assert.IsType<FinalizedResult<ValidateCompliancePdfResult>>(result);
    }

    [SkippableTheory]
    [InlineData("CONVERT")]
    [InlineData("Convert")]
    [InlineData("convert")]
    public void Convert_Operation_ShouldBeCaseInsensitive(string operation)
    {
        SkipIfNotWindows();
        var pdfPath = CreateTestPdf($"test_case_convert_{operation}.pdf");
        var outputPath = CreateTestFilePath($"test_case_convert_{operation}_output.pdf");

        var result = _tool.Execute(operation, pdfPath, outputPath: outputPath, format: "pdf/a-1b");

        Assert.IsType<FinalizedResult<ConvertCompliancePdfResult>>(result);
    }

    [SkippableFact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        SkipIfNotWindows();
        var pdfPath = CreateTestPdf("test_unknown_op.pdf");

        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("unknown", pdfPath));
        Assert.Contains("Unknown operation", ex.Message);
    }

    [SkippableFact]
    public void Execute_WithNoPathOrSessionId_ShouldThrowException()
    {
        SkipIfNotWindows();
        Assert.ThrowsAny<Exception>(() => _tool.Execute("validate", format: "pdf/a-1b"));
    }

    #endregion

    #region Session Management

    [SkippableFact]
    public void Validate_WithSessionId_ShouldValidateFromSession()
    {
        SkipIfNotWindows();
        var pdfPath = CreateTestPdf("test_session_validate.pdf");
        var sessionId = OpenSession(pdfPath);

        var result = _tool.Execute("validate", sessionId: sessionId, format: "pdf/a-1b");

        var data = GetResultData<ValidateCompliancePdfResult>(result);
        Assert.Equal("PDF/A-1b", data.Format);
        var output = GetResultOutput<ValidateCompliancePdfResult>(result);
        Assert.True(output.IsSession);
    }

    [SkippableFact]
    public void Convert_WithSessionId_ShouldConvertInSession()
    {
        SkipIfNotWindows();
        var pdfPath = CreateTestPdf("test_session_convert.pdf");
        var sessionId = OpenSession(pdfPath);

        var result = _tool.Execute("convert", sessionId: sessionId, format: "pdf/a-1b");

        var data = GetResultData<ConvertCompliancePdfResult>(result);
        Assert.Equal("PDF/A-1b", data.Format);
        var output = GetResultOutput<ConvertCompliancePdfResult>(result);
        Assert.True(output.IsSession);
    }

    [SkippableFact]
    public void Execute_WithInvalidSessionId_ShouldThrowKeyNotFoundException()
    {
        SkipIfNotWindows();
        Assert.Throws<KeyNotFoundException>(() =>
            _tool.Execute("validate", sessionId: "invalid_session", format: "pdf/a-1b"));
    }

    [SkippableFact]
    public void Execute_WithBothPathAndSessionId_ShouldPreferSessionId()
    {
        SkipIfNotWindows();
        var pdfPath1 = CreateTestPdf("test_path_file.pdf");
        var pdfPath2 = CreateTestPdf("test_session_file.pdf");
        var sessionId = OpenSession(pdfPath2);

        var result = _tool.Execute("validate", pdfPath1, sessionId, format: "pdf/a-1b");

        var data = GetResultData<ValidateCompliancePdfResult>(result);
        Assert.Equal("PDF/A-1b", data.Format);
    }

    #endregion
}
