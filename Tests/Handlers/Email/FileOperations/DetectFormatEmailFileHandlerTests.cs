using Aspose.Email;
using AsposeMcpServer.Handlers.Email.FileOperations;
using AsposeMcpServer.Results.Email.FileOperations;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Email.FileOperations;

/// <summary>
///     Tests for <see cref="DetectFormatEmailFileHandler" />.
///     Verifies format detection for various email file types.
/// </summary>
public class DetectFormatEmailFileHandlerTests : HandlerTestBase<object>
{
    private readonly DetectFormatEmailFileHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_DetectFormat()
    {
        Assert.Equal("detect_format", _handler.Operation);
    }

    #endregion

    #region Helper Methods

    /// <summary>
    ///     Creates an EML file for testing.
    /// </summary>
    /// <param name="fileName">The output file name.</param>
    /// <returns>The full path to the created file.</returns>
    private string CreateEmlFile(string fileName)
    {
        var filePath = CreateTestFilePath(fileName);
        var message = new MailMessage
        {
            From = "sender@example.com",
            Subject = "Detect Format Test",
            Body = "Test Body"
        };
        message.To.Add("recipient@example.com");
        message.Save(filePath, SaveOptions.DefaultEml);
        return filePath;
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithNonExistentFile_ThrowsFileNotFoundException()
    {
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", Path.Combine(TestDir, "nonexistent.eml") }
        });

        Assert.Throws<FileNotFoundException>(() => _handler.Execute(context, parameters));
    }

    #endregion

    #region Parameter Validation

    [Fact]
    public void Execute_WithoutPath_ThrowsArgumentException()
    {
        var context = CreateContext(new object());
        var parameters = CreateEmptyParameters();

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion

    #region Basic Operations

    [Fact]
    public void Execute_DetectsEmlFormat()
    {
        var path = CreateEmlFile("test_detect.eml");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", path }
        });

        var result = _handler.Execute(context, parameters);

        Assert.IsType<DetectFormatEmailResult>(result);
        var detectResult = (DetectFormatEmailResult)result;
        Assert.NotNull(detectResult.Format);
        Assert.NotNull(detectResult.Extension);
        Assert.Contains("Detected format", detectResult.Message);
    }

    [Fact]
    public void Execute_ReturnsExtensionWithDot()
    {
        var path = CreateEmlFile("test_detect_ext.eml");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", path }
        });

        var result = _handler.Execute(context, parameters);

        Assert.IsType<DetectFormatEmailResult>(result);
        var detectResult = (DetectFormatEmailResult)result;
        Assert.StartsWith(".", detectResult.Extension);
    }

    [SkippableFact]
    public void Execute_DetectsMsgFormat()
    {
        SkipInEvaluationMode(AsposeLibraryType.Email);
        var emlPath = CreateEmlFile("source_for_msg.eml");
        var msgPath = CreateTestFilePath("test_detect.msg");

        var message = MailMessage.Load(emlPath);
        message.Save(msgPath, SaveOptions.DefaultMsgUnicode);

        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", msgPath }
        });

        var result = _handler.Execute(context, parameters);

        Assert.IsType<DetectFormatEmailResult>(result);
        var detectResult = (DetectFormatEmailResult)result;
        Assert.NotNull(detectResult.Format);
        Assert.NotNull(detectResult.Extension);
    }

    [Fact]
    public void Execute_DetectsMhtFormat()
    {
        var emlPath = CreateEmlFile("source_for_mht.eml");
        var mhtPath = CreateTestFilePath("test_detect.mht");

        var message = MailMessage.Load(emlPath);
        message.Save(mhtPath, SaveOptions.DefaultMhtml);

        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", mhtPath }
        });

        var result = _handler.Execute(context, parameters);

        Assert.IsType<DetectFormatEmailResult>(result);
        var detectResult = (DetectFormatEmailResult)result;
        Assert.NotNull(detectResult.Format);
        Assert.Contains("Detected format", detectResult.Message);
    }

    [Fact]
    public void Execute_MessageContainsFormatAndExtension()
    {
        var path = CreateEmlFile("test_message.eml");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", path }
        });

        var result = _handler.Execute(context, parameters);

        Assert.IsType<DetectFormatEmailResult>(result);
        var detectResult = (DetectFormatEmailResult)result;
        Assert.Contains(detectResult.Format, detectResult.Message);
        Assert.Contains(detectResult.Extension, detectResult.Message);
    }

    #endregion
}
