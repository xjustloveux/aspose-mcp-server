using Aspose.Email;
using AsposeMcpServer.Handlers.Email.Conversion;
using AsposeMcpServer.Results.Email.Conversion;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Email.Conversion;

/// <summary>
///     Tests for <see cref="ConvertEmailHandler" />.
///     Verifies email format conversion between EML, EMLX, MSG, MHT/MHTML, and HTML.
/// </summary>
public class ConvertEmailHandlerTests : HandlerTestBase<object>
{
    private readonly ConvertEmailHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Convert()
    {
        Assert.Equal("convert", _handler.Operation);
    }

    #endregion

    #region Helper Methods

    /// <summary>
    ///     Creates a test EML file with default content.
    /// </summary>
    private string CreateTestEmlFile(string fileName, string subject = "Test Subject", string body = "Test Body")
    {
        var filePath = CreateTestFilePath(fileName);
        var message = new MailMessage
        {
            From = "sender@example.com",
            Subject = subject,
            Body = body
        };
        message.To.Add("recipient@example.com");
        message.Save(filePath, SaveOptions.DefaultEml);
        return filePath;
    }

    #endregion

    #region Basic Operations

    [Fact]
    public void Execute_EmlToHtml_ConvertsSuccessfully()
    {
        var path = CreateTestEmlFile("test_convert_source.eml");
        var outputPath = CreateTestFilePath("test_convert_output.html");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", path },
            { "outputPath", outputPath }
        });

        var result = _handler.Execute(context, parameters);

        var conversionResult = Assert.IsType<EmailConversionResult>(result);
        Assert.Equal(path, conversionResult.SourcePath);
        Assert.Equal(outputPath, conversionResult.OutputPath);
        Assert.Equal("EML", conversionResult.SourceFormat);
        Assert.Equal("HTML", conversionResult.TargetFormat);
        Assert.NotNull(conversionResult.FileSize);
        Assert.True(conversionResult.FileSize > 0);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void Execute_EmlToMhtml_ConvertsSuccessfully()
    {
        var path = CreateTestEmlFile("test_eml_to_mhtml.eml");
        var outputPath = CreateTestFilePath("test_eml_to_mhtml.mhtml");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", path },
            { "outputPath", outputPath }
        });

        var result = _handler.Execute(context, parameters);

        var conversionResult = Assert.IsType<EmailConversionResult>(result);
        Assert.Equal("EML", conversionResult.SourceFormat);
        Assert.Equal("MHTML", conversionResult.TargetFormat);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void Execute_EmlToMht_ConvertsSuccessfully()
    {
        var path = CreateTestEmlFile("test_eml_to_mht.eml");
        var outputPath = CreateTestFilePath("test_eml_to_mht.mht");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", path },
            { "outputPath", outputPath }
        });

        var result = _handler.Execute(context, parameters);

        var conversionResult = Assert.IsType<EmailConversionResult>(result);
        Assert.Equal("EML", conversionResult.SourceFormat);
        Assert.Equal("MHT", conversionResult.TargetFormat);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void Execute_EmlToEml_ConvertsSuccessfully()
    {
        var path = CreateTestEmlFile("test_eml_to_eml.eml");
        var outputPath = CreateTestFilePath("test_eml_to_eml_copy.eml");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", path },
            { "outputPath", outputPath }
        });

        var result = _handler.Execute(context, parameters);

        var conversionResult = Assert.IsType<EmailConversionResult>(result);
        Assert.Equal("EML", conversionResult.SourceFormat);
        Assert.Equal("EML", conversionResult.TargetFormat);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void Execute_EmlToEmlx_ConvertsSuccessfully()
    {
        var path = CreateTestEmlFile("test_eml_to_emlx.eml");
        var outputPath = CreateTestFilePath("test_eml_to_emlx.emlx");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", path },
            { "outputPath", outputPath }
        });

        var result = _handler.Execute(context, parameters);

        var conversionResult = Assert.IsType<EmailConversionResult>(result);
        Assert.Equal("EML", conversionResult.SourceFormat);
        Assert.Equal("EMLX", conversionResult.TargetFormat);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void Execute_EmlToHtm_ConvertsSuccessfully()
    {
        var path = CreateTestEmlFile("test_eml_to_htm.eml");
        var outputPath = CreateTestFilePath("test_eml_to_htm.htm");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", path },
            { "outputPath", outputPath }
        });

        var result = _handler.Execute(context, parameters);

        var conversionResult = Assert.IsType<EmailConversionResult>(result);
        Assert.Equal("EML", conversionResult.SourceFormat);
        Assert.Equal("HTM", conversionResult.TargetFormat);
        Assert.True(File.Exists(outputPath));
    }

    [SkippableFact]
    public void Execute_EmlToMsg_ConvertsSuccessfully()
    {
        SkipInEvaluationMode(AsposeLibraryType.Email);
        var path = CreateTestEmlFile("test_eml_to_msg.eml");
        var outputPath = CreateTestFilePath("test_eml_to_msg.msg");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", path },
            { "outputPath", outputPath }
        });

        var result = _handler.Execute(context, parameters);

        var conversionResult = Assert.IsType<EmailConversionResult>(result);
        Assert.Equal("EML", conversionResult.SourceFormat);
        Assert.Equal("MSG", conversionResult.TargetFormat);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void Execute_ResultContainsFileSize()
    {
        var path = CreateTestEmlFile("test_filesize.eml");
        var outputPath = CreateTestFilePath("test_filesize.html");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", path },
            { "outputPath", outputPath }
        });

        var result = _handler.Execute(context, parameters);

        var conversionResult = Assert.IsType<EmailConversionResult>(result);
        Assert.NotNull(conversionResult.FileSize);
        Assert.True(conversionResult.FileSize > 0);
    }

    [Fact]
    public void Execute_ResultContainsMessage()
    {
        var path = CreateTestEmlFile("test_message.eml");
        var outputPath = CreateTestFilePath("test_message.html");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", path },
            { "outputPath", outputPath }
        });

        var result = _handler.Execute(context, parameters);

        var conversionResult = Assert.IsType<EmailConversionResult>(result);
        Assert.Contains("converted", conversionResult.Message);
        Assert.Contains(".eml", conversionResult.Message);
        Assert.Contains(".html", conversionResult.Message);
    }

    [Fact]
    public void Execute_PreservesEmailContent()
    {
        var path = CreateTestEmlFile("test_preserve.eml", "Preserved Subject", "Preserved Body Content");
        var outputPath = CreateTestFilePath("test_preserve.msg");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", path },
            { "outputPath", outputPath }
        });

        _handler.Execute(context, parameters);

        var loaded = MailMessage.Load(outputPath);
        Assert.Contains("Preserved Subject", loaded.Subject);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithMissingPath_ThrowsArgumentException()
    {
        var outputPath = CreateTestFilePath("output.html");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "outputPath", outputPath }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithMissingOutputPath_ThrowsArgumentException()
    {
        var path = CreateTestEmlFile("test_no_output.eml");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", path }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithNonExistentFile_ThrowsFileNotFoundException()
    {
        var path = CreateTestFilePath("nonexistent.eml");
        var outputPath = CreateTestFilePath("output.html");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", path },
            { "outputPath", outputPath }
        });

        Assert.Throws<FileNotFoundException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithUnsupportedFormat_ThrowsArgumentException()
    {
        var path = CreateTestEmlFile("test_unsupported.eml");
        var outputPath = CreateTestFilePath("test_unsupported.xyz");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", path },
            { "outputPath", outputPath }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Theory]
    [InlineData(".pdf")]
    [InlineData(".doc")]
    [InlineData(".txt")]
    [InlineData(".rtf")]
    public void Execute_WithVariousUnsupportedFormats_ThrowsArgumentException(string extension)
    {
        var path = CreateTestEmlFile($"test_unsupported_{extension.TrimStart('.')}.eml");
        var outputPath = CreateTestFilePath($"test_unsupported{extension}");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", path },
            { "outputPath", outputPath }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("Unsupported target format", ex.Message);
        Assert.Contains(extension, ex.Message);
    }

    #endregion
}
