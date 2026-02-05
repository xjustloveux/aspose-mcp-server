using Aspose.Email;
using AsposeMcpServer.Handlers.Email.FileOperations;
using AsposeMcpServer.Results.Email.FileOperations;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Email.FileOperations;

/// <summary>
///     Tests for <see cref="ConvertEmailFileHandler" />.
///     Verifies email format conversion between EML, MSG, MHTML, and HTML.
/// </summary>
public class ConvertEmailFileHandlerTests : HandlerTestBase<object>
{
    private readonly ConvertEmailFileHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Convert()
    {
        Assert.Equal("convert", _handler.Operation);
    }

    #endregion

    #region Helper Methods

    /// <summary>
    ///     Creates an EML file for testing.
    /// </summary>
    /// <param name="fileName">The output file name.</param>
    /// <param name="subject">The email subject.</param>
    /// <returns>The full path to the created file.</returns>
    private string CreateEmlFile(string fileName, string subject = "Convert Test")
    {
        var filePath = CreateTestFilePath(fileName);
        var message = new MailMessage
        {
            From = "sender@example.com",
            Subject = subject,
            Body = "Test Body"
        };
        message.To.Add("recipient@example.com");
        message.Save(filePath, SaveOptions.DefaultEml);
        return filePath;
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithNonExistentSource_ThrowsFileNotFoundException()
    {
        var outputPath = CreateTestFilePath("output_err.mhtml");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", Path.Combine(TestDir, "nonexistent.eml") },
            { "outputPath", outputPath }
        });

        Assert.Throws<FileNotFoundException>(() => _handler.Execute(context, parameters));
    }

    #endregion

    #region Basic Operations

    [Fact]
    public void Execute_ConvertsEmlToMhtml()
    {
        var sourcePath = CreateEmlFile("source_convert.eml");
        var outputPath = CreateTestFilePath("output.mhtml");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", sourcePath },
            { "outputPath", outputPath }
        });

        var result = _handler.Execute(context, parameters);

        Assert.IsType<EmailConversionResult>(result);
        var conversion = (EmailConversionResult)result;
        Assert.Equal(sourcePath, conversion.SourcePath);
        Assert.Equal(outputPath, conversion.OutputPath);
        Assert.Equal("EML", conversion.SourceFormat);
        Assert.Equal("MHTML", conversion.TargetFormat);
        Assert.Contains("EML", conversion.Message);
        Assert.Contains("MHTML", conversion.Message);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void Execute_ConvertsEmlToHtml()
    {
        var sourcePath = CreateEmlFile("source_to_html.eml");
        var outputPath = CreateTestFilePath("output.html");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", sourcePath },
            { "outputPath", outputPath }
        });

        var result = _handler.Execute(context, parameters);

        Assert.IsType<EmailConversionResult>(result);
        var conversion = (EmailConversionResult)result;
        Assert.Equal("EML", conversion.SourceFormat);
        Assert.Equal("HTML", conversion.TargetFormat);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void Execute_ConvertsEmlToHtm()
    {
        var sourcePath = CreateEmlFile("source_to_htm.eml");
        var outputPath = CreateTestFilePath("output.htm");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", sourcePath },
            { "outputPath", outputPath }
        });

        var result = _handler.Execute(context, parameters);

        Assert.IsType<EmailConversionResult>(result);
        var conversion = (EmailConversionResult)result;
        Assert.Equal("HTML", conversion.TargetFormat);
        Assert.True(File.Exists(outputPath));
    }

    [SkippableFact]
    public void Execute_ConvertsEmlToMsg()
    {
        SkipInEvaluationMode(AsposeLibraryType.Email);
        var sourcePath = CreateEmlFile("source_to_msg.eml");
        var outputPath = CreateTestFilePath("output.msg");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", sourcePath },
            { "outputPath", outputPath }
        });

        var result = _handler.Execute(context, parameters);

        Assert.IsType<EmailConversionResult>(result);
        var conversion = (EmailConversionResult)result;
        Assert.Equal("EML", conversion.SourceFormat);
        Assert.Equal("MSG", conversion.TargetFormat);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void Execute_ConvertsEmlToMht()
    {
        var sourcePath = CreateEmlFile("source_to_mht.eml");
        var outputPath = CreateTestFilePath("output.mht");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", sourcePath },
            { "outputPath", outputPath }
        });

        var result = _handler.Execute(context, parameters);

        Assert.IsType<EmailConversionResult>(result);
        var conversion = (EmailConversionResult)result;
        Assert.Equal("MHTML", conversion.TargetFormat);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void Execute_ConvertsEmlToEml()
    {
        var sourcePath = CreateEmlFile("source_to_eml.eml");
        var outputPath = CreateTestFilePath("output_copy.eml");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", sourcePath },
            { "outputPath", outputPath }
        });

        var result = _handler.Execute(context, parameters);

        Assert.IsType<EmailConversionResult>(result);
        var conversion = (EmailConversionResult)result;
        Assert.Equal("EML", conversion.SourceFormat);
        Assert.Equal("EML", conversion.TargetFormat);
        Assert.True(File.Exists(outputPath));

        var loaded = MailMessage.Load(outputPath);
        Assert.Equal("Convert Test", loaded.Subject);
    }

    [Fact]
    public void Execute_UnknownExtension_DefaultsToEml()
    {
        var sourcePath = CreateEmlFile("source_to_unknown.eml");
        var outputPath = CreateTestFilePath("output.xyz");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", sourcePath },
            { "outputPath", outputPath }
        });

        var result = _handler.Execute(context, parameters);

        Assert.IsType<EmailConversionResult>(result);
        var conversion = (EmailConversionResult)result;
        Assert.Equal("EML", conversion.TargetFormat);
        Assert.True(File.Exists(outputPath));
    }

    #endregion

    #region Parameter Validation

    [Fact]
    public void Execute_WithoutPath_ThrowsArgumentException()
    {
        var outputPath = CreateTestFilePath("output_no_path.mhtml");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "outputPath", outputPath }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithoutOutputPath_ThrowsArgumentException()
    {
        var sourcePath = CreateEmlFile("source_no_output.eml");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", sourcePath }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithBothMissing_ThrowsArgumentException()
    {
        var context = CreateContext(new object());
        var parameters = CreateEmptyParameters();

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion
}
