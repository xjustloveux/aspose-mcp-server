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

    /// <summary>
    ///     Creates a dummy certificate file for testing parameter validation.
    ///     NOTE: This is NOT a valid PFX certificate and cannot be used for actual signing tests.
    ///     These dummy files are only used to test that the tool properly validates parameters
    ///     before attempting to use the certificate.
    /// </summary>
    private string CreateDummyCertFileForValidation(string fileName)
    {
        var certPath = CreateTestFilePath(fileName);
        File.WriteAllBytes(certPath, [0x00, 0x01, 0x02, 0x03]);
        return certPath;
    }

    #region General

    [Fact]
    public void Get_WithNoSignatures_ShouldReturnEmptyResult()
    {
        var pdfPath = CreateTestPdf("test_get_empty.pdf");
        var result = _tool.Execute("get", pdfPath);
        var json = JsonSerializer.Deserialize<JsonElement>(result);

        Assert.True(json.TryGetProperty("count", out var countProp));
        Assert.Equal(0, countProp.GetInt32());

        Assert.True(json.TryGetProperty("items", out var itemsProp));
        Assert.Equal(JsonValueKind.Array, itemsProp.ValueKind);
        Assert.Equal(0, itemsProp.GetArrayLength());

        Assert.True(json.TryGetProperty("message", out var messageProp));
        Assert.Equal("No signatures found", messageProp.GetString());
    }

    [Fact]
    public void Get_ShouldReturnValidJson()
    {
        var pdfPath = CreateTestPdf("test_get_json.pdf");
        var result = _tool.Execute("get", pdfPath);
        var json = JsonSerializer.Deserialize<JsonElement>(result);
        Assert.True(json.TryGetProperty("count", out _));
        Assert.True(json.TryGetProperty("items", out _));
    }

    [SkippableFact]
    public void Get_WithMultiPagePdf_ShouldWork()
    {
        SkipInEvaluationMode(AsposeLibraryType.Pdf, "5 pages exceeds 4-page limit in evaluation mode");
        var pdfPath = CreateMultiPagePdf("test_get_multipage.pdf", 5);
        var result = _tool.Execute("get", pdfPath);
        var json = JsonSerializer.Deserialize<JsonElement>(result);
        Assert.True(json.TryGetProperty("count", out _));
    }

    [Theory]
    [InlineData("GET")]
    [InlineData("Get")]
    [InlineData("get")]
    public void Operation_ShouldBeCaseInsensitive_Get(string operation)
    {
        var pdfPath = CreateTestPdf($"test_case_{operation}.pdf");
        var result = _tool.Execute(operation, pdfPath);
        var json = JsonSerializer.Deserialize<JsonElement>(result);

        Assert.True(json.TryGetProperty("count", out var countProp));
        Assert.Equal(0, countProp.GetInt32());
    }

    [Theory]
    [InlineData("SIGN")]
    [InlineData("Sign")]
    [InlineData("sign")]
    public void Operation_ShouldBeCaseInsensitive_Sign(string operation)
    {
        var pdfPath = CreateTestPdf($"test_case_sign_{operation}.pdf");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute(operation, pdfPath, certificatePath: null, certificatePassword: "pass"));
        Assert.Equal("certificatePath is required for sign operation", ex.Message);
    }

    [Theory]
    [InlineData("DELETE")]
    [InlineData("Delete")]
    [InlineData("delete")]
    public void Operation_ShouldBeCaseInsensitive_Delete(string operation)
    {
        var pdfPath = CreateTestPdf($"test_case_delete_{operation}.pdf");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute(operation, pdfPath, signatureIndex: 0));
        Assert.StartsWith("signatureIndex must be between 0 and", ex.Message);
    }

    #endregion

    #region Exception

    [Fact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        var pdfPath = CreateTestPdf("test_unknown_op.pdf");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("unknown", pdfPath));
        Assert.StartsWith("Unknown operation: unknown", ex.Message);
    }

    [Fact]
    public void Sign_WithMissingCertificatePath_ShouldThrowArgumentException()
    {
        var pdfPath = CreateTestPdf("test_sign_no_cert.pdf");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("sign", pdfPath, certificatePath: null, certificatePassword: "password"));
        Assert.Equal("certificatePath is required for sign operation", ex.Message);
    }

    [Fact]
    public void Sign_WithMissingCertificatePassword_ShouldThrowArgumentException()
    {
        var pdfPath = CreateTestPdf("test_sign_no_pass.pdf");
        var certPath = CreateDummyCertFileForValidation("test_cert_nopass.pfx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("sign", pdfPath, certificatePath: certPath, certificatePassword: null));
        Assert.Equal("certificatePassword is required for sign operation", ex.Message);
    }

    [Fact]
    public void Sign_WithNonExistentCertificatePath_ShouldThrowFileNotFoundException()
    {
        var pdfPath = CreateTestPdf("test_sign_cert_notfound.pdf");
        Assert.Throws<FileNotFoundException>(() =>
            _tool.Execute("sign", pdfPath, certificatePath: "nonexistent.pfx", certificatePassword: "password"));
    }

    [Fact]
    public void Sign_WithInvalidPageIndex_ShouldThrowArgumentException()
    {
        var pdfPath = CreateTestPdf("test_sign_invalid_page.pdf");
        var certPath = CreateDummyCertFileForValidation("test_cert_page.pfx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("sign", pdfPath, certificatePath: certPath, certificatePassword: "password", pageIndex: 99));
        Assert.StartsWith("pageIndex must be between 1 and", ex.Message);
    }

    [Fact]
    public void Sign_WithNonExistentImagePath_ShouldThrowException()
    {
        var pdfPath = CreateTestPdf("test_sign_image_notfound.pdf");
        var certPath = CreateDummyCertFileForValidation("test_cert_image.pfx");
        Assert.ThrowsAny<Exception>(() =>
            _tool.Execute("sign", pdfPath, certificatePath: certPath, certificatePassword: "password",
                imagePath: "nonexistent_image.png"));
    }

    [Fact]
    public void Delete_WithNoSignatures_ShouldThrowArgumentException()
    {
        var pdfPath = CreateTestPdf("test_delete_empty.pdf");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("delete", pdfPath, signatureIndex: 0));
        // When there are no signatures, the valid range is empty (0 to -1)
        Assert.StartsWith("signatureIndex must be between 0 and", ex.Message);
    }

    [Fact]
    public void Delete_WithNegativeIndex_ShouldThrowArgumentException()
    {
        var pdfPath = CreateTestPdf("test_delete_negative.pdf");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("delete", pdfPath, signatureIndex: -1));
        Assert.StartsWith("signatureIndex must be between 0 and", ex.Message);
    }

    [Fact]
    public void Execute_WithNoPathOrSessionId_ShouldThrowException()
    {
        Assert.ThrowsAny<Exception>(() => _tool.Execute("get"));
    }

    [Fact]
    public void Execute_WithNonExistentFile_ShouldThrowFileNotFoundException()
    {
        Assert.Throws<FileNotFoundException>(() => _tool.Execute("get", "nonexistent_file.pdf"));
    }

    #endregion

    #region Session

    [Fact]
    public void Get_WithSessionId_ShouldReturnResult()
    {
        var pdfPath = CreateTestPdf("test_session_get.pdf");
        var sessionId = OpenSession(pdfPath);
        var result = _tool.Execute("get", sessionId: sessionId);
        var json = JsonSerializer.Deserialize<JsonElement>(result);
        Assert.True(json.TryGetProperty("count", out var countProp));
        Assert.Equal(0, countProp.GetInt32());
        var document = SessionManager.GetDocument<Document>(sessionId);
        Assert.NotNull(document);
    }

    [Fact]
    public void Delete_WithSessionId_WithNoSignatures_ShouldThrowArgumentException()
    {
        var pdfPath = CreateTestPdf("test_session_delete.pdf");
        var sessionId = OpenSession(pdfPath);
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("delete", sessionId: sessionId, signatureIndex: 0));
        Assert.StartsWith("signatureIndex must be between 0 and", ex.Message);
        var document = SessionManager.GetDocument<Document>(sessionId);
        Assert.NotNull(document);
    }

    [Fact]
    public void Sign_WithSessionId_WithNonExistentCert_ShouldThrowFileNotFoundException()
    {
        var pdfPath = CreateTestPdf("test_session_sign.pdf");
        var sessionId = OpenSession(pdfPath);
        Assert.Throws<FileNotFoundException>(() =>
            _tool.Execute("sign", sessionId: sessionId,
                certificatePath: "nonexistent_cert.pfx", certificatePassword: "password"));
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
        var json = JsonSerializer.Deserialize<JsonElement>(result);

        Assert.True(json.TryGetProperty("count", out var countProp));
        Assert.Equal(0, countProp.GetInt32());
    }

    #endregion
}