using Aspose.Email;
using AsposeMcpServer.Results;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Results.Email.FileOperations;
using AsposeMcpServer.Tests.Infrastructure;
using AsposeMcpServer.Tools.Email;

namespace AsposeMcpServer.Tests.Tools.Email;

/// <summary>
///     Integration tests for <see cref="EmailFileTool" />.
///     Focuses on operation routing, file I/O, and parameter forwarding.
///     Detailed parameter validation and business logic tests are in Handler tests.
/// </summary>
public class EmailFileToolTests : EmailTestBase
{
    private readonly EmailFileTool _tool = new();

    #region Create Operation

    [Fact]
    public void Create_ShouldCreateEmlFile()
    {
        var outputPath = CreateTestFilePath("test_create.eml");

        var result = _tool.Execute("create", outputPath: outputPath,
            subject: "Test Subject", body: "Test Body",
            from: "sender@example.com", to: "recipient@example.com");

        var data = GetResultData<SuccessResult>(result);
        Assert.Contains(outputPath, data.Message);
        Assert.True(File.Exists(outputPath));

        var loaded = MailMessage.Load(outputPath);
        Assert.Contains("Test Subject", loaded.Subject);
    }

    [Fact]
    public void Create_WithHtmlBody_ShouldCreateHtmlEmail()
    {
        var outputPath = CreateTestFilePath("test_create_html.eml");

        var result = _tool.Execute("create", outputPath: outputPath,
            subject: "HTML", body: "<h1>Hello</h1>", isHtml: true);

        var data = GetResultData<SuccessResult>(result);
        Assert.Contains(outputPath, data.Message);
        Assert.True(File.Exists(outputPath));

        var loaded = MailMessage.Load(outputPath);
        Assert.Contains("<h1>Hello</h1>", loaded.HtmlBody);
    }

    [SkippableFact]
    public void Create_ShouldCreateMsgFile()
    {
        SkipInEvaluationMode(AsposeLibraryType.Email);
        var outputPath = CreateTestFilePath("test_create.msg");

        var result = _tool.Execute("create", outputPath: outputPath,
            subject: "MSG Test", from: "sender@example.com");

        Assert.IsType<FinalizedResult<SuccessResult>>(result);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void Create_ShouldCreateMhtmlFile()
    {
        var outputPath = CreateTestFilePath("test_create.mhtml");

        var result = _tool.Execute("create", outputPath: outputPath,
            subject: "MHTML Test");

        Assert.IsType<FinalizedResult<SuccessResult>>(result);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void Create_WithMinimalParams_ShouldSucceed()
    {
        var outputPath = CreateTestFilePath("test_create_minimal.eml");

        var result = _tool.Execute("create", outputPath: outputPath);

        Assert.IsType<FinalizedResult<SuccessResult>>(result);
        Assert.True(File.Exists(outputPath));
    }

    #endregion

    #region GetInfo Operation

    [Fact]
    public void GetInfo_ShouldReturnEmailFileInfo()
    {
        var emlPath = CreateEmlFile("test_info.eml", "Info Subject", "Info Body");

        var result = _tool.Execute("get_info", emlPath);

        var data = GetResultData<EmailFileInfo>(result);
        Assert.Contains("Info Subject", data.Subject!);
        Assert.Equal("sender@example.com", data.From);
        Assert.Contains("recipient@example.com", data.To!);
        Assert.Equal("EML", data.Format);
        Assert.False(data.HasAttachments);
        Assert.Equal(0, data.AttachmentCount);
    }

    [SkippableFact]
    public void GetInfo_MsgFile_ShouldReturnMsgFormat()
    {
        SkipInEvaluationMode(AsposeLibraryType.Email);
        var msgPath = CreateMsgFile("test_info.msg", "MSG Info");

        var result = _tool.Execute("get_info", msgPath);

        var data = GetResultData<EmailFileInfo>(result);
        Assert.Equal("MSG", data.Format);
    }

    [Fact]
    public void GetInfo_WithNonExistentFile_ShouldThrowFileNotFoundException()
    {
        var fakePath = CreateTestFilePath("nonexistent.eml");

        Assert.Throws<FileNotFoundException>(() =>
            _tool.Execute("get_info", fakePath));
    }

    #endregion

    #region Save Operation

    [Fact]
    public void Save_ShouldSaveToNewPath()
    {
        var emlPath = CreateEmlFile("test_save_source.eml", "Save Test");
        var outputPath = CreateTestFilePath("test_save_output.eml");

        var result = _tool.Execute("save", emlPath, outputPath);

        var data = GetResultData<SuccessResult>(result);
        Assert.Contains(outputPath, data.Message);
        Assert.True(File.Exists(outputPath));

        var loaded = MailMessage.Load(outputPath);
        Assert.Contains("Save Test", loaded.Subject);
    }

    [Fact]
    public void Save_ToMhtml_ShouldConvertFormat()
    {
        var emlPath = CreateEmlFile("test_save_to_mhtml.eml");
        var outputPath = CreateTestFilePath("test_save_output.mhtml");

        var result = _tool.Execute("save", emlPath, outputPath);

        Assert.IsType<FinalizedResult<SuccessResult>>(result);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void Save_WithNonExistentSource_ShouldThrowFileNotFoundException()
    {
        var fakePath = CreateTestFilePath("nonexistent_save.eml");
        var outputPath = CreateTestFilePath("save_output.eml");

        Assert.Throws<FileNotFoundException>(() =>
            _tool.Execute("save", fakePath, outputPath));
    }

    #endregion

    #region Convert Operation

    [Fact]
    public void Convert_EmlToMhtml_ShouldConvert()
    {
        var emlPath = CreateEmlFile("test_convert_source.eml", "Convert Test");
        var outputPath = CreateTestFilePath("test_convert_output.mhtml");

        var result = _tool.Execute("convert", emlPath, outputPath);

        var data = GetResultData<EmailConversionResult>(result);
        Assert.Equal("EML", data.SourceFormat);
        Assert.Equal("MHTML", data.TargetFormat);
        Assert.Equal(emlPath, data.SourcePath);
        Assert.Equal(outputPath, data.OutputPath);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void Convert_EmlToHtml_ShouldConvert()
    {
        var emlPath = CreateEmlFile("test_convert_to_html.eml");
        var outputPath = CreateTestFilePath("test_convert_output.html");

        var result = _tool.Execute("convert", emlPath, outputPath);

        var data = GetResultData<EmailConversionResult>(result);
        Assert.Equal("HTML", data.TargetFormat);
        Assert.True(File.Exists(outputPath));
    }

    [SkippableFact]
    public void Convert_EmlToMsg_ShouldConvert()
    {
        SkipInEvaluationMode(AsposeLibraryType.Email);
        var emlPath = CreateEmlFile("test_convert_to_msg.eml");
        var outputPath = CreateTestFilePath("test_convert_output.msg");

        var result = _tool.Execute("convert", emlPath, outputPath);

        var data = GetResultData<EmailConversionResult>(result);
        Assert.Equal("MSG", data.TargetFormat);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void Convert_WithNonExistentSource_ShouldThrowFileNotFoundException()
    {
        var fakePath = CreateTestFilePath("nonexistent_convert.eml");
        var outputPath = CreateTestFilePath("convert_output.mhtml");

        Assert.Throws<FileNotFoundException>(() =>
            _tool.Execute("convert", fakePath, outputPath));
    }

    #endregion

    #region DetectFormat Operation

    [Fact]
    public void DetectFormat_EmlFile_ShouldDetect()
    {
        var emlPath = CreateEmlFile("test_detect.eml");

        var result = _tool.Execute("detect_format", emlPath);

        var data = GetResultData<DetectFormatEmailResult>(result);
        Assert.NotNull(data.Format);
        Assert.NotNull(data.Extension);
        Assert.Contains("Detected format", data.Message);
    }

    [SkippableFact]
    public void DetectFormat_MsgFile_ShouldDetect()
    {
        SkipInEvaluationMode(AsposeLibraryType.Email);
        var msgPath = CreateMsgFile("test_detect.msg");

        var result = _tool.Execute("detect_format", msgPath);

        var data = GetResultData<DetectFormatEmailResult>(result);
        Assert.NotNull(data.Format);
        Assert.NotNull(data.Extension);
    }

    [Fact]
    public void DetectFormat_WithNonExistentFile_ShouldThrowFileNotFoundException()
    {
        var fakePath = CreateTestFilePath("nonexistent_detect.eml");

        Assert.Throws<FileNotFoundException>(() =>
            _tool.Execute("detect_format", fakePath));
    }

    #endregion

    #region Operation Routing

    [Fact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        var emlPath = CreateEmlFile("test_unknown.eml");

        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("unknown_operation", emlPath));
        Assert.Contains("Unknown operation", ex.Message);
    }

    [Theory]
    [InlineData("CREATE")]
    [InlineData("Create")]
    [InlineData("create")]
    public void Execute_OperationShouldBeCaseInsensitive(string operation)
    {
        var outputPath = CreateTestFilePath($"test_case_{operation}.eml");

        var result = _tool.Execute(operation, outputPath: outputPath,
            subject: "Case Test");

        Assert.IsType<FinalizedResult<SuccessResult>>(result);
        Assert.True(File.Exists(outputPath));
    }

    #endregion

    #region Missing Required Parameters

    [Fact]
    public void Create_WithoutOutputPath_ShouldThrowArgumentException()
    {
        Assert.Throws<ArgumentException>(() =>
            _tool.Execute("create", subject: "No Output"));
    }

    [Fact]
    public void GetInfo_WithoutPath_ShouldThrowArgumentException()
    {
        Assert.Throws<ArgumentException>(() =>
            _tool.Execute("get_info"));
    }

    [Fact]
    public void Save_WithoutPath_ShouldThrowArgumentException()
    {
        var outputPath = CreateTestFilePath("save_no_path.eml");

        Assert.Throws<ArgumentException>(() =>
            _tool.Execute("save", outputPath: outputPath));
    }

    [Fact]
    public void Save_WithoutOutputPath_ShouldThrowArgumentException()
    {
        var emlPath = CreateEmlFile("save_no_output.eml");

        Assert.Throws<ArgumentException>(() =>
            _tool.Execute("save", emlPath));
    }

    [Fact]
    public void Convert_WithoutPath_ShouldThrowArgumentException()
    {
        var outputPath = CreateTestFilePath("convert_no_path.mhtml");

        Assert.Throws<ArgumentException>(() =>
            _tool.Execute("convert", outputPath: outputPath));
    }

    [Fact]
    public void Convert_WithoutOutputPath_ShouldThrowArgumentException()
    {
        var emlPath = CreateEmlFile("convert_no_output.eml");

        Assert.Throws<ArgumentException>(() =>
            _tool.Execute("convert", emlPath));
    }

    [Fact]
    public void DetectFormat_WithoutPath_ShouldThrowArgumentException()
    {
        Assert.Throws<ArgumentException>(() =>
            _tool.Execute("detect_format"));
    }

    #endregion
}
