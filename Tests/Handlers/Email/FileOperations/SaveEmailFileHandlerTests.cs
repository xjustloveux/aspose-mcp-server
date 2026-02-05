using Aspose.Email;
using AsposeMcpServer.Handlers.Email.FileOperations;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Email.FileOperations;

/// <summary>
///     Tests for <see cref="SaveEmailFileHandler" />.
///     Verifies loading an email from one path and saving it to another.
/// </summary>
public class SaveEmailFileHandlerTests : HandlerTestBase<object>
{
    private readonly SaveEmailFileHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Save()
    {
        Assert.Equal("save", _handler.Operation);
    }

    #endregion

    #region Helper Methods

    /// <summary>
    ///     Creates an EML file for testing.
    /// </summary>
    /// <param name="fileName">The output file name.</param>
    /// <param name="subject">The email subject.</param>
    /// <returns>The full path to the created file.</returns>
    private string CreateEmlFile(string fileName, string subject = "Test Subject")
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
        var outputPath = CreateTestFilePath("output_missing.eml");
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

    [SkippableFact]
    public void Execute_SavesEmlToEml()
    {
        SkipInEvaluationMode(AsposeLibraryType.Email, "Evaluation mode appends watermark to subject");
        var sourcePath = CreateEmlFile("source.eml", "Save Test");
        var outputPath = CreateTestFilePath("output.eml");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", sourcePath },
            { "outputPath", outputPath }
        });

        var result = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(result);
        var success = (SuccessResult)result;
        Assert.Contains(outputPath, success.Message);
        Assert.True(File.Exists(outputPath));

        var loaded = MailMessage.Load(outputPath);
        Assert.Equal("Save Test", loaded.Subject);
    }

    [Fact]
    public void Execute_SavesEmlToMhtml()
    {
        var sourcePath = CreateEmlFile("source_mhtml.eml", "MHTML Save");
        var outputPath = CreateTestFilePath("output.mhtml");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", sourcePath },
            { "outputPath", outputPath }
        });

        var result = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(result);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void Execute_SavesEmlToHtml()
    {
        var sourcePath = CreateEmlFile("source_html.eml", "HTML Save");
        var outputPath = CreateTestFilePath("output.html");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", sourcePath },
            { "outputPath", outputPath }
        });

        var result = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(result);
        Assert.True(File.Exists(outputPath));
    }

    [SkippableFact]
    public void Execute_SavesEmlToMsg()
    {
        SkipInEvaluationMode(AsposeLibraryType.Email);
        var sourcePath = CreateEmlFile("source_msg.eml", "MSG Save");
        var outputPath = CreateTestFilePath("output.msg");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", sourcePath },
            { "outputPath", outputPath }
        });

        var result = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(result);
        Assert.True(File.Exists(outputPath));
    }

    [SkippableFact]
    public void Execute_PreservesEmailContent()
    {
        SkipInEvaluationMode(AsposeLibraryType.Email, "Evaluation mode appends watermark to subject");
        var sourcePath = CreateEmlFile("source_preserve.eml", "Preserved Subject");
        var outputPath = CreateTestFilePath("output_preserve.eml");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", sourcePath },
            { "outputPath", outputPath }
        });

        _handler.Execute(context, parameters);

        var loaded = MailMessage.Load(outputPath);
        Assert.Equal("Preserved Subject", loaded.Subject);
        Assert.Equal("sender@example.com", loaded.From.Address);
        Assert.Contains(loaded.To, addr => addr.Address == "recipient@example.com");
    }

    #endregion

    #region Parameter Validation

    [Fact]
    public void Execute_WithoutPath_ThrowsArgumentException()
    {
        var outputPath = CreateTestFilePath("output_no_path.eml");
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
