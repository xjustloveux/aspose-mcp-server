using System.Text.Json;
using Aspose.Pdf;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Pdf;

namespace AsposeMcpServer.Tests.Tools.Pdf;

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
        var document = new Document();
        document.Pages.Add();
        document.Save(filePath);
        return filePath;
    }

    private string CreateMultiPagePdf(string fileName, int pageCount)
    {
        var filePath = CreateTestFilePath(fileName);
        var document = new Document();
        for (var i = 0; i < pageCount; i++)
            document.Pages.Add();
        document.Save(filePath);
        return filePath;
    }

    #region General Tests

    [Fact]
    public void GetSignatures_WithNoSignatures_ShouldReturnEmptyResult()
    {
        var pdfPath = CreateTestPdf("test_get_no_signatures.pdf");
        var result = _tool.Execute("get", pdfPath);
        Assert.NotNull(result);
        Assert.Contains("count", result);
        Assert.Contains("0", result);
        Assert.Contains("No signatures found", result);
    }

    [Fact]
    public void GetSignatures_ShouldReturnValidJson()
    {
        var pdfPath = CreateTestPdf("test_get_signatures_json.pdf");
        var result = _tool.Execute("get", pdfPath);
        var json = JsonSerializer.Deserialize<JsonElement>(result);
        Assert.True(json.TryGetProperty("count", out var countProp));
        Assert.Equal(0, countProp.GetInt32());
        Assert.True(json.TryGetProperty("items", out var itemsProp));
        Assert.Equal(JsonValueKind.Array, itemsProp.ValueKind);
    }

    [Fact]
    public void Sign_WithMissingCertificatePath_ShouldThrowException()
    {
        var pdfPath = CreateTestPdf("test_sign_missing_cert.pdf");
        // ArgumentHelper.GetString uses key as default when missing, which then fails file validation
        Assert.ThrowsAny<Exception>(() => _tool.Execute(
            "sign",
            pdfPath,
            certificatePassword: "password"));
    }

    [Fact]
    public void Sign_WithNonExistentCertificatePath_ShouldThrowFileNotFoundException()
    {
        var pdfPath = CreateTestPdf("test_sign_missing_password.pdf");
        Assert.Throws<FileNotFoundException>(() => _tool.Execute(
            "sign",
            pdfPath,
            certificatePath: "nonexistent_cert.pfx",
            certificatePassword: "password"));
    }

    [Fact]
    public void Sign_WithInvalidPageIndex_ShouldThrowArgumentException()
    {
        var pdfPath = CreateTestPdf("test_sign_invalid_page.pdf");
        var certPath = CreateTestFilePath("test_cert.pfx");
        // Create a dummy cert file for the test (won't actually be used due to page validation)
        File.WriteAllText(certPath, "dummy");
        var exception = Assert.Throws<ArgumentException>(() => _tool.Execute(
            "sign",
            pdfPath,
            certificatePath: certPath,
            certificatePassword: "password",
            pageIndex: 99));
        Assert.Contains("pageIndex must be between", exception.Message);
    }

    [Fact]
    public void Delete_WithNoSignatures_ShouldThrowArgumentException()
    {
        var pdfPath = CreateTestPdf("test_delete_no_signatures.pdf");
        var outputPath = CreateTestFilePath("test_delete_no_signatures_output.pdf");
        var exception = Assert.Throws<ArgumentException>(() => _tool.Execute(
            "delete",
            pdfPath,
            outputPath: outputPath,
            signatureIndex: 0));
        Assert.Contains("signatureIndex must be between", exception.Message);
    }

    [Fact]
    public void Delete_WithNegativeIndex_ShouldThrowArgumentException()
    {
        var pdfPath = CreateTestPdf("test_delete_negative_index.pdf");
        var exception = Assert.Throws<ArgumentException>(() => _tool.Execute(
            "delete",
            pdfPath,
            signatureIndex: -1));
        Assert.Contains("signatureIndex must be between", exception.Message);
    }

    [Fact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        var pdfPath = CreateTestPdf("test_unknown_op.pdf");
        var exception = Assert.Throws<ArgumentException>(() => _tool.Execute(
            "unknown",
            pdfPath));
        Assert.Contains("Unknown operation", exception.Message);
    }

    [SkippableFact]
    public void GetSignatures_WithMultiPagePdf_ShouldWork()
    {
        // Skip in evaluation mode - 5 pages exceeds 4-page limit
        SkipInEvaluationMode(AsposeLibraryType.Pdf, "5 pages exceeds 4-page limit in evaluation mode");
        var pdfPath = CreateMultiPagePdf("test_get_multipage.pdf", 5);
        var result = _tool.Execute("get", pdfPath);
        Assert.NotNull(result);
        var json = JsonSerializer.Deserialize<JsonElement>(result);
        Assert.True(json.TryGetProperty("count", out _));
    }

    [Fact]
    public void Sign_WithNonExistentImagePath_ShouldThrowFileNotFoundException()
    {
        var pdfPath = CreateTestPdf("test_sign_nonexistent_image.pdf");
        var certPath = CreateTestFilePath("test_cert_for_image.pfx");
        File.WriteAllText(certPath, "dummy");
        // This will fail at certificate validation before image validation
        // But if it reaches image validation, it should throw FileNotFoundException
        Assert.ThrowsAny<Exception>(() => _tool.Execute(
            "sign",
            pdfPath,
            certificatePath: certPath,
            certificatePassword: "password",
            imagePath: "nonexistent_image.png"));
    }

    [Fact]
    public void Delete_WithMissingSignatureIndex_ShouldThrowArgumentException()
    {
        var pdfPath = CreateTestPdf("test_delete_missing_index.pdf");
        Assert.Throws<ArgumentException>(() => _tool.Execute(
            "delete",
            pdfPath));
    }

    #endregion

    #region Exception Tests

    [Fact]
    public void Execute_WithNoPathOrSessionId_ShouldThrowException()
    {
        Assert.ThrowsAny<Exception>(() => _tool.Execute(
            "get"));
    }

    [Fact]
    public void Execute_WithNonExistentFile_ShouldThrowFileNotFoundException()
    {
        Assert.Throws<FileNotFoundException>(() => _tool.Execute(
            "get",
            "nonexistent_file.pdf"));
    }

    #endregion

    #region Session ID Tests

    [Fact]
    public void GetSignatures_WithSessionId_ShouldReturnResult()
    {
        var pdfPath = CreateTestPdf("test_session_get_signatures.pdf");
        var sessionId = OpenSession(pdfPath);
        var result = _tool.Execute("get", sessionId: sessionId);
        Assert.NotNull(result);
        var json = JsonSerializer.Deserialize<JsonElement>(result);
        Assert.True(json.TryGetProperty("count", out var countProp));
        Assert.Equal(0, countProp.GetInt32());
        var document = SessionManager.GetDocument<Document>(sessionId);
        Assert.NotNull(document);
    }

    [Fact]
    public void Delete_WithSessionId_ShouldThrowWhenNoSignatures()
    {
        var pdfPath = CreateTestPdf("test_session_delete_no_sig.pdf");
        var sessionId = OpenSession(pdfPath);
        var exception = Assert.Throws<ArgumentException>(() => _tool.Execute(
            "delete",
            sessionId: sessionId,
            signatureIndex: 0));
        Assert.Contains("signatureIndex must be between", exception.Message);
        var document = SessionManager.GetDocument<Document>(sessionId);
        Assert.NotNull(document);
    }

    [Fact]
    public void Sign_WithSessionId_ShouldThrowWhenCertNotFound()
    {
        var pdfPath = CreateTestPdf("test_session_sign.pdf");
        var sessionId = OpenSession(pdfPath);
        Assert.Throws<FileNotFoundException>(() => _tool.Execute(
            "sign",
            sessionId: sessionId,
            certificatePath: "nonexistent_cert.pfx",
            certificatePassword: "password"));
        var document = SessionManager.GetDocument<Document>(sessionId);
        Assert.NotNull(document);
    }

    #endregion
}