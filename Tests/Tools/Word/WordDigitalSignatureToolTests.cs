using AsposeMcpServer.Results;
using AsposeMcpServer.Results.Word.DigitalSignature;
using AsposeMcpServer.Tests.Infrastructure;
using AsposeMcpServer.Tools.Word;

namespace AsposeMcpServer.Tests.Tools.Word;

/// <summary>
///     Integration tests for WordDigitalSignatureTool.
///     Note: Sign operation requires a real PFX certificate, so only verify/list/remove are tested here.
/// </summary>
public class WordDigitalSignatureToolTests : WordTestBase
{
    private readonly WordDigitalSignatureTool _tool;

    public WordDigitalSignatureToolTests()
    {
        _tool = new WordDigitalSignatureTool(SessionManager);
    }

    #region File I/O Smoke Tests

    [Fact]
    public void VerifySignatures_WithUnsignedDocument_ShouldReturnNoSignatures()
    {
        var docPath = CreateWordDocumentWithContent("test_sig_verify.docx", "Test content");
        var result = _tool.Execute("verify", docPath);
        var data = GetResultData<VerifySignaturesResult>(result);
        Assert.Equal(0, data.TotalCount);
        Assert.False(data.AllValid);
    }

    [Fact]
    public void ListSignatures_WithUnsignedDocument_ShouldReturnEmptyList()
    {
        var docPath = CreateWordDocumentWithContent("test_sig_list.docx", "Test content");
        var result = _tool.Execute("list", docPath);
        var data = GetResultData<GetSignaturesResult>(result);
        Assert.Equal(0, data.Count);
        Assert.Empty(data.Signatures);
    }

    [Fact]
    public void RemoveSignatures_WithUnsignedDocument_ShouldSucceed()
    {
        var docPath = CreateWordDocumentWithContent("test_sig_remove.docx", "Test content");
        var outputPath = CreateTestFilePath("test_sig_remove_output.docx");
        var result = _tool.Execute("remove", docPath, outputPath);
        Assert.NotNull(result);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void Sign_WithNonExistentCertificate_ShouldThrowException()
    {
        var docPath = CreateWordDocumentWithContent("test_sig_sign_fail.docx", "Test content");
        var outputPath = CreateTestFilePath("test_sig_sign_output.docx");
        Assert.ThrowsAny<Exception>(() =>
            _tool.Execute("sign", docPath, outputPath,
                Path.Combine(TestDir, "nonexistent.pfx"),
                "pass"));
    }

    #endregion

    #region Operation Routing

    [Theory]
    [InlineData("VERIFY")]
    [InlineData("Verify")]
    [InlineData("verify")]
    public void Operation_ShouldBeCaseInsensitive(string operation)
    {
        var docPath = CreateWordDocumentWithContent($"test_sig_case_{operation}.docx", "Test");
        var result = _tool.Execute(operation, docPath);
        Assert.IsType<FinalizedResult<VerifySignaturesResult>>(result);
    }

    [Fact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocumentWithContent("test_sig_unknown.docx", "Test");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("unknown_operation", docPath));
        Assert.Contains("Unknown operation", ex.Message);
    }

    [Fact]
    public void Execute_WithNoPath_ShouldThrowException()
    {
        Assert.ThrowsAny<Exception>(() => _tool.Execute("verify"));
    }

    #endregion
}
