using System.Runtime.Versioning;
using Aspose.Slides;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;
using AsposeMcpServer.Tools.PowerPoint;

namespace AsposeMcpServer.Tests.Tools.PowerPoint;

/// <summary>
///     Integration tests for PptFileOperationsTool.
///     Focuses on session management, file I/O, and operation routing.
///     Detailed parameter validation and business logic tests are in Handler tests.
/// </summary>
[SupportedOSPlatform("windows")]
public class PptFileOperationsToolTests : PptTestBase
{
    private readonly PptFileOperationsTool _tool;

    public PptFileOperationsToolTests()
    {
        _tool = new PptFileOperationsTool(SessionManager);
    }

    #region File I/O Smoke Tests

    [Fact]
    public void Create_ShouldCreateNewPresentation()
    {
        var outputPath = CreateTestFilePath("test_create.pptx");
        var result = _tool.Execute("create", path: outputPath);
        var data = GetResultData<SuccessResult>(result);
        Assert.StartsWith("PowerPoint presentation created successfully", data.Message);
        Assert.True(File.Exists(outputPath));
        using var presentation = new Presentation(outputPath);
        Assert.True(presentation.Slides.Count > 0);
    }

    [Fact]
    public void Convert_ToPdf_ShouldConvert()
    {
        var pptPath = CreatePresentation("test_convert_pdf.pptx");
        var outputPath = CreateTestFilePath("test_convert_output.pdf");
        var result = _tool.Execute("convert", inputPath: pptPath, outputPath: outputPath, format: "pdf");
        var data = GetResultData<SuccessResult>(result);
        Assert.StartsWith("Presentation from", data.Message);
        Assert.Contains("converted to PDF format", data.Message);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void Merge_ShouldMergePresentations()
    {
        var ppt1Path = CreatePresentation("test_merge1.pptx", 2);
        var ppt2Path = CreatePresentation("test_merge2.pptx", 2);
        var outputPath = CreateTestFilePath("test_merge_output.pptx");
        var result = _tool.Execute("merge", outputPath: outputPath, inputPaths: [ppt1Path, ppt2Path],
            keepSourceFormatting: true);
        var data = GetResultData<SuccessResult>(result);
        Assert.StartsWith("Merged 2 presentations", data.Message);
        Assert.True(File.Exists(outputPath));
        using var presentation = new Presentation(outputPath);
        Assert.Equal(4, presentation.Slides.Count);
    }

    [Fact]
    public void Split_ShouldSplitIntoMultipleFiles()
    {
        var pptPath = CreatePresentation("test_split.pptx", 3);
        var outputDir = Path.Combine(TestDir, "split_output");
        Directory.CreateDirectory(outputDir);
        var result = _tool.Execute("split", inputPath: pptPath, outputDirectory: outputDir);
        var data = GetResultData<SuccessResult>(result);
        Assert.StartsWith("Split presentation into 3 file(s)", data.Message);
        var files = Directory.GetFiles(outputDir, "*.pptx");
        Assert.Equal(3, files.Length);
    }

    #endregion

    #region Operation Routing

    [Theory]
    [InlineData("CREATE")]
    [InlineData("Create")]
    [InlineData("create")]
    public void Operation_ShouldBeCaseInsensitive(string operation)
    {
        var outputPath = CreateTestFilePath($"test_case_create_{operation}.pptx");
        var result = _tool.Execute(operation, path: outputPath);
        var data = GetResultData<SuccessResult>(result);
        Assert.StartsWith("PowerPoint presentation created successfully", data.Message);
    }

    [Fact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("unknown", path: "test.pptx"));
        Assert.Contains("Unknown operation", ex.Message);
    }

    #endregion

    #region Session Management

    [Fact]
    public void Convert_WithSessionId_ShouldConvertFromSession()
    {
        var pptPath = CreatePresentation("test_session_convert.pptx", 2);
        var sessionId = OpenSession(pptPath);
        var outputPath = CreateTestFilePath("test_session_convert_output.pdf");
        var result = _tool.Execute("convert", sessionId, outputPath: outputPath, format: "pdf");
        var data = GetResultData<SuccessResult>(result);
        Assert.StartsWith("Presentation from", data.Message);
        Assert.Contains("converted to PDF format", data.Message);
        Assert.True(File.Exists(outputPath));
        var output = GetResultOutput<SuccessResult>(result);
        Assert.True(output.IsSession);
    }

    [Fact]
    public void Split_WithSessionId_ShouldSplitFromSession()
    {
        var pptPath = CreatePresentation("test_session_split.pptx", 3);
        var sessionId = OpenSession(pptPath);
        var outputDir = Path.Combine(TestDir, "session_split_output");
        Directory.CreateDirectory(outputDir);
        var result = _tool.Execute("split", sessionId, outputDirectory: outputDir);
        var data = GetResultData<SuccessResult>(result);
        Assert.StartsWith("Split presentation into 3 file(s)", data.Message);
        var output = GetResultOutput<SuccessResult>(result);
        Assert.True(output.IsSession);
    }

    [Fact]
    public void Convert_WithInvalidSessionId_ShouldThrowKeyNotFoundException()
    {
        var outputPath = CreateTestFilePath("test_invalid_session.pdf");
        Assert.Throws<KeyNotFoundException>(() =>
            _tool.Execute("convert", "invalid_session", outputPath: outputPath, format: "pdf"));
    }

    [Fact]
    public void Split_WithBothPathAndSessionId_ShouldPreferSessionId()
    {
        var pptPath1 = CreatePresentation("test_path_split.pptx");
        var pptPath2 = CreatePresentation("test_session_split2.pptx", 3);
        var sessionId = OpenSession(pptPath2);
        var outputDir = Path.Combine(TestDir, "prefer_session_split");
        Directory.CreateDirectory(outputDir);
        var result = _tool.Execute("split", sessionId, inputPath: pptPath1, outputDirectory: outputDir);
        var data = GetResultData<SuccessResult>(result);
        Assert.Contains("3 file(s)", data.Message);
    }

    #endregion
}
