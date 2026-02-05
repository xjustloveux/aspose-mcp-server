using Aspose.Email;
using AsposeMcpServer.Handlers.Email.FileOperations;
using AsposeMcpServer.Results.Email.FileOperations;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Email.FileOperations;

/// <summary>
///     Tests for <see cref="LoadEmailFileHandler" />.
///     Verifies loading email files and returning correct metadata.
/// </summary>
public class LoadEmailFileHandlerTests : HandlerTestBase<object>
{
    private readonly LoadEmailFileHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_GetInfo()
    {
        Assert.Equal("get_info", _handler.Operation);
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

    #region Helper Methods

    /// <summary>
    ///     Creates an EML file for testing.
    /// </summary>
    /// <param name="fileName">The output file name.</param>
    /// <param name="subject">The email subject.</param>
    /// <param name="from">The sender address.</param>
    /// <param name="to">The recipient address.</param>
    /// <returns>The full path to the created file.</returns>
    private string CreateEmlFile(string fileName, string subject = "Test Subject",
        string from = "sender@example.com", string to = "recipient@example.com")
    {
        var filePath = CreateTestFilePath(fileName);
        var message = new MailMessage
        {
            From = from,
            Subject = subject,
            Body = "Test Body"
        };
        message.To.Add(to);
        message.Save(filePath, SaveOptions.DefaultEml);
        return filePath;
    }

    /// <summary>
    ///     Creates an EML file with attachments for testing.
    /// </summary>
    /// <param name="fileName">The output file name.</param>
    /// <param name="attachmentCount">The number of attachments to add.</param>
    /// <returns>The full path to the created file.</returns>
    private string CreateEmlWithAttachments(string fileName, int attachmentCount)
    {
        var filePath = CreateTestFilePath(fileName);
        var message = new MailMessage
        {
            From = "sender@example.com",
            Subject = "Email With Attachments",
            Body = "Test Body"
        };
        message.To.Add("recipient@example.com");

        for (var i = 0; i < attachmentCount; i++)
        {
            var attachFile = CreateTempFile(".txt", $"Attachment content {i}");
            message.Attachments.Add(new Aspose.Email.Attachment(attachFile));
        }

        message.Save(filePath, SaveOptions.DefaultEml);
        return filePath;
    }

    #endregion

    #region Basic Operations

    [Fact]
    public void Execute_ReturnsEmailFileInfo()
    {
        var path = CreateEmlFile("test_load.eml");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", path }
        });

        var result = _handler.Execute(context, parameters);

        Assert.IsType<EmailFileInfo>(result);
        var info = (EmailFileInfo)result;
        Assert.Equal("Test Subject", info.Subject);
        Assert.Equal("sender@example.com", info.From);
        Assert.Contains("recipient@example.com", info.To!);
        Assert.Equal("EML", info.Format);
        Assert.False(info.HasAttachments);
        Assert.Equal(0, info.AttachmentCount);
        Assert.Contains(path, info.Message);
    }

    [Fact]
    public void Execute_WithAttachments_ReportsAttachmentInfo()
    {
        var path = CreateEmlWithAttachments("test_with_attachments.eml", 2);
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", path }
        });

        var result = _handler.Execute(context, parameters);

        Assert.IsType<EmailFileInfo>(result);
        var info = (EmailFileInfo)result;
        Assert.True(info.HasAttachments);
        Assert.Equal(2, info.AttachmentCount);
    }

    [SkippableFact]
    public void Execute_MsgFile_DetectsFormatAsMSG()
    {
        SkipInEvaluationMode(AsposeLibraryType.Email);
        var emlPath = CreateEmlFile("source.eml");
        var msgPath = CreateTestFilePath("test_load.msg");

        var message = MailMessage.Load(emlPath);
        message.Save(msgPath, SaveOptions.DefaultMsgUnicode);

        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", msgPath }
        });

        var result = _handler.Execute(context, parameters);

        Assert.IsType<EmailFileInfo>(result);
        var info = (EmailFileInfo)result;
        Assert.Equal("MSG", info.Format);
    }

    [Fact]
    public void Execute_MhtmlFile_DetectsFormatAsMHTML()
    {
        var emlPath = CreateEmlFile("source_mhtml.eml");
        var mhtmlPath = CreateTestFilePath("test_load.mhtml");

        var message = MailMessage.Load(emlPath);
        message.Save(mhtmlPath, SaveOptions.DefaultMhtml);

        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", mhtmlPath }
        });

        var result = _handler.Execute(context, parameters);

        Assert.IsType<EmailFileInfo>(result);
        var info = (EmailFileInfo)result;
        Assert.Equal("MHTML", info.Format);
    }

    [Fact]
    public void Execute_MhtFile_DetectsFormatAsMHTML()
    {
        var emlPath = CreateEmlFile("source_mht.eml");
        var mhtPath = CreateTestFilePath("test_load.mht");

        var message = MailMessage.Load(emlPath);
        message.Save(mhtPath, SaveOptions.DefaultMhtml);

        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", mhtPath }
        });

        var result = _handler.Execute(context, parameters);

        Assert.IsType<EmailFileInfo>(result);
        var info = (EmailFileInfo)result;
        Assert.Equal("MHTML", info.Format);
    }

    [Fact]
    public void Execute_HtmlFile_DetectsFormatAsHTML()
    {
        var emlPath = CreateEmlFile("source_html.eml");
        var htmlPath = CreateTestFilePath("test_load.html");

        var message = MailMessage.Load(emlPath);
        message.Save(htmlPath, SaveOptions.DefaultHtml);

        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", htmlPath }
        });

        var result = _handler.Execute(context, parameters);

        Assert.IsType<EmailFileInfo>(result);
        var info = (EmailFileInfo)result;
        Assert.Equal("HTML", info.Format);
    }

    [Fact]
    public void Execute_UnknownExtension_DetectsFormatAsUnknown()
    {
        var emlPath = CreateEmlFile("source_unknown.eml");
        var unknownPath = CreateTestFilePath("test_load.xyz");

        var message = MailMessage.Load(emlPath);
        message.Save(unknownPath, SaveOptions.DefaultEml);

        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", unknownPath }
        });

        var result = _handler.Execute(context, parameters);

        Assert.IsType<EmailFileInfo>(result);
        var info = (EmailFileInfo)result;
        Assert.Equal("Unknown", info.Format);
    }

    #endregion
}
