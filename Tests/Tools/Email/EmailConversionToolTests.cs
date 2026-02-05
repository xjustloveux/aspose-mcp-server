using Aspose.Email;
using AsposeMcpServer.Results.Email.Conversion;
using AsposeMcpServer.Tests.Infrastructure;
using AsposeMcpServer.Tools.Email;

namespace AsposeMcpServer.Tests.Tools.Email;

/// <summary>
///     Integration tests for <see cref="EmailConversionTool" />.
///     Focuses on operation routing, file I/O, and end-to-end conversion operations.
///     Detailed parameter validation and business logic tests are in Handler tests.
/// </summary>
public class EmailConversionToolTests : EmailTestBase
{
    private readonly EmailConversionTool _tool = new();

    #region Operation Routing

    [Fact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        var path = CreateEmlFile("test_unknown_op.eml");
        var outputPath = CreateTestFilePath("test_unknown_op.html");

        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("unknown_operation", path, outputPath));
        Assert.Contains("Unknown operation", ex.Message);
    }

    #endregion

    #region Convert Operations

    [Fact]
    public void Convert_EmlToHtml_ShouldConvert()
    {
        var path = CreateEmlFile("test_convert_to_html.eml", "HTML Convert Test", "Body for HTML");
        var outputPath = CreateTestFilePath("test_convert_to_html.html");

        var result = _tool.Execute("convert", path, outputPath);

        Assert.NotNull(result);
        var data = GetResultData<EmailConversionResult>(result);
        Assert.Equal("EML", data.SourceFormat);
        Assert.Equal("HTML", data.TargetFormat);
        Assert.True(File.Exists(outputPath));
        Assert.NotNull(data.FileSize);
        Assert.True(data.FileSize > 0);
    }

    [Fact]
    public void Convert_EmlToMhtml_ShouldConvert()
    {
        var path = CreateEmlFile("test_convert_to_mhtml.eml", "MHTML Test", "Body for MHTML");
        var outputPath = CreateTestFilePath("test_convert_to_mhtml.mhtml");

        var result = _tool.Execute("convert", path, outputPath);

        var data = GetResultData<EmailConversionResult>(result);
        Assert.Equal("EML", data.SourceFormat);
        Assert.Equal("MHTML", data.TargetFormat);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void Convert_EmlToMht_ShouldConvert()
    {
        var path = CreateEmlFile("test_convert_to_mht.eml", "MHT Test", "Body for MHT");
        var outputPath = CreateTestFilePath("test_convert_to_mht.mht");

        var result = _tool.Execute("convert", path, outputPath);

        var data = GetResultData<EmailConversionResult>(result);
        Assert.Equal("EML", data.SourceFormat);
        Assert.Equal("MHT", data.TargetFormat);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void Convert_EmlToHtm_ShouldConvert()
    {
        var path = CreateEmlFile("test_convert_to_htm.eml", "HTM Test", "Body for HTM");
        var outputPath = CreateTestFilePath("test_convert_to_htm.htm");

        var result = _tool.Execute("convert", path, outputPath);

        var data = GetResultData<EmailConversionResult>(result);
        Assert.Equal("EML", data.SourceFormat);
        Assert.Equal("HTM", data.TargetFormat);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void Convert_EmlToEml_ShouldConvert()
    {
        var path = CreateEmlFile("test_convert_eml_eml.eml", "EML to EML", "Body for EML");
        var outputPath = CreateTestFilePath("test_convert_eml_eml_out.eml");

        var result = _tool.Execute("convert", path, outputPath);

        var data = GetResultData<EmailConversionResult>(result);
        Assert.Equal("EML", data.SourceFormat);
        Assert.Equal("EML", data.TargetFormat);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void Convert_EmlToEmlx_ShouldConvert()
    {
        var path = CreateEmlFile("test_convert_to_emlx.eml", "EMLX Test", "Body for EMLX");
        var outputPath = CreateTestFilePath("test_convert_to_emlx.emlx");

        var result = _tool.Execute("convert", path, outputPath);

        var data = GetResultData<EmailConversionResult>(result);
        Assert.Equal("EML", data.SourceFormat);
        Assert.Equal("EMLX", data.TargetFormat);
        Assert.True(File.Exists(outputPath));
    }

    [SkippableFact]
    public void Convert_EmlToMsg_ShouldConvert()
    {
        SkipInEvaluationMode(AsposeLibraryType.Email);
        var path = CreateEmlFile("test_convert_to_msg.eml", "MSG Convert", "Body for MSG");
        var outputPath = CreateTestFilePath("test_convert_to_msg.msg");

        var result = _tool.Execute("convert", path, outputPath);

        var data = GetResultData<EmailConversionResult>(result);
        Assert.Equal("EML", data.SourceFormat);
        Assert.Equal("MSG", data.TargetFormat);
        Assert.True(File.Exists(outputPath));
    }

    [SkippableFact]
    public void Convert_MsgToEml_ShouldConvert()
    {
        SkipInEvaluationMode(AsposeLibraryType.Email);
        var path = CreateMsgFile("test_msg_to_eml.msg", "MSG to EML", "Body from MSG");
        var outputPath = CreateTestFilePath("test_msg_to_eml.eml");

        var result = _tool.Execute("convert", path, outputPath);

        var data = GetResultData<EmailConversionResult>(result);
        Assert.Equal("MSG", data.SourceFormat);
        Assert.Equal("EML", data.TargetFormat);
        Assert.True(File.Exists(outputPath));
    }

    [SkippableFact]
    public void Convert_MsgToHtml_ShouldConvert()
    {
        SkipInEvaluationMode(AsposeLibraryType.Email);
        var path = CreateMsgFile("test_msg_to_html.msg", "MSG to HTML", "Body from MSG");
        var outputPath = CreateTestFilePath("test_msg_to_html.html");

        var result = _tool.Execute("convert", path, outputPath);

        var data = GetResultData<EmailConversionResult>(result);
        Assert.Equal("MSG", data.SourceFormat);
        Assert.Equal("HTML", data.TargetFormat);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void Convert_PreservesEmailSubject()
    {
        var path = CreateEmlFile("test_preserve_subject.eml", "Preserved Subject", "Preserved Body");
        var outputPath = CreateTestFilePath("test_preserve_subject.msg");

        _tool.Execute("convert", path, outputPath);

        var loaded = MailMessage.Load(outputPath);
        Assert.Contains("Preserved Subject", loaded.Subject);
    }

    [Fact]
    public void Convert_ResultContainsSourceAndOutputPaths()
    {
        var path = CreateEmlFile("test_paths.eml");
        var outputPath = CreateTestFilePath("test_paths.html");

        var result = _tool.Execute("convert", path, outputPath);

        var data = GetResultData<EmailConversionResult>(result);
        Assert.Equal(path, data.SourcePath);
        Assert.Equal(outputPath, data.OutputPath);
    }

    [Fact]
    public void Convert_ResultContainsMessage()
    {
        var path = CreateEmlFile("test_msg_field.eml");
        var outputPath = CreateTestFilePath("test_msg_field.html");

        var result = _tool.Execute("convert", path, outputPath);

        var data = GetResultData<EmailConversionResult>(result);
        Assert.Contains("converted", data.Message);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Convert_WithUnsupportedFormat_ShouldThrowArgumentException()
    {
        var path = CreateEmlFile("test_unsupported_fmt.eml");
        var outputPath = CreateTestFilePath("test_unsupported_fmt.xyz");

        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("convert", path, outputPath));
        Assert.Contains("Unsupported target format", ex.Message);
    }

    [Fact]
    public void Convert_WithNonExistentFile_ShouldThrowFileNotFoundException()
    {
        var path = CreateTestFilePath("nonexistent.eml");
        var outputPath = CreateTestFilePath("output.html");

        Assert.Throws<FileNotFoundException>(() =>
            _tool.Execute("convert", path, outputPath));
    }

    [Fact]
    public void Convert_WithMissingPath_ShouldThrowArgumentException()
    {
        var outputPath = CreateTestFilePath("output.html");

        Assert.Throws<ArgumentException>(() =>
            _tool.Execute("convert", outputPath: outputPath));
    }

    [Fact]
    public void Convert_WithMissingOutputPath_ShouldThrowArgumentException()
    {
        var path = CreateEmlFile("test_no_output.eml");

        Assert.Throws<ArgumentException>(() =>
            _tool.Execute("convert", path));
    }

    #endregion
}
