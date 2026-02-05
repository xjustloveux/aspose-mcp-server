using Aspose.Email;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Results.Email.Content;
using AsposeMcpServer.Tests.Infrastructure;
using AsposeMcpServer.Tools.Email;

namespace AsposeMcpServer.Tests.Tools.Email;

/// <summary>
///     Integration tests for <see cref="EmailContentTool" />.
///     Focuses on operation routing, file I/O, and end-to-end content operations.
///     Detailed parameter validation and business logic tests are in Handler tests.
/// </summary>
public class EmailContentToolTests : EmailTestBase
{
    private readonly EmailContentTool _tool = new();

    #region Get Headers

    [Fact]
    public void GetHeaders_ShouldReturnHeaders()
    {
        var path = CreateEmlFile("test_get_headers.eml");
        var result = _tool.Execute("get_headers", path);

        Assert.NotNull(result);
        var data = GetResultData<EmailHeadersResult>(result);
        Assert.True(data.Count > 0);
        Assert.Equal(data.Headers.Count, data.Count);
    }

    #endregion

    #region Get Recipients

    [Fact]
    public void GetRecipients_ShouldReturnRecipients()
    {
        var path = CreateEmlFile("test_get_recipients.eml");
        var result = _tool.Execute("get_recipients", path);

        Assert.NotNull(result);
        var data = GetResultData<EmailRecipientsResult>(result);
        Assert.Equal("sender@example.com", data.From);
        Assert.NotEmpty(data.To);
        Assert.Contains("recipient@example.com", data.To);
    }

    #endregion

    #region Helper Methods

    /// <summary>
    ///     Creates a test EML file with an HTML body.
    /// </summary>
    private string CreateHtmlEmlFile(string fileName, string subject, string htmlBody)
    {
        var filePath = CreateTestFilePath(fileName);
        var message = new MailMessage
        {
            From = "sender@example.com",
            Subject = subject,
            HtmlBody = htmlBody
        };
        message.To.Add("recipient@example.com");
        message.Save(filePath, SaveOptions.DefaultEml);
        return filePath;
    }

    #endregion

    #region Get Body

    [Fact]
    public void GetBody_ShouldReturnBody()
    {
        var path = CreateEmlFile("test_get_body.eml", "Test Subject", "Body content here");
        var result = _tool.Execute("get_body", path);

        Assert.NotNull(result);
        var data = GetResultData<EmailBodyResult>(result);
        Assert.Contains("Body content here", data.Body);
    }

    [Fact]
    public void GetBody_WithHtmlEmail_ShouldReturnHtmlBody()
    {
        var path = CreateHtmlEmlFile("test_get_html_body.eml", "HTML Test", "<p>HTML body</p>");
        var result = _tool.Execute("get_body", path);

        var data = GetResultData<EmailBodyResult>(result);
        Assert.True(data.IsHtml);
        Assert.Contains("<p>HTML body</p>", data.HtmlBody);
    }

    #endregion

    #region Set Body

    [Fact]
    public void SetBody_ShouldSetPlainTextBody()
    {
        var path = CreateEmlFile("test_set_body.eml");
        var result = _tool.Execute("set_body", path, body: "New body text");

        Assert.NotNull(result);
        var data = GetResultData<SuccessResult>(result);
        Assert.Contains("body updated", data.Message);

        var loaded = MailMessage.Load(path);
        Assert.Contains("New body text", loaded.Body);
    }

    [Fact]
    public void SetBody_WithHtml_ShouldSetHtmlBody()
    {
        var path = CreateEmlFile("test_set_html_body.eml");
        var result = _tool.Execute("set_body", path, body: "<h1>Title</h1>", isHtml: true);

        var data = GetResultData<SuccessResult>(result);
        Assert.Contains("HTML body updated", data.Message);

        var loaded = MailMessage.Load(path);
        Assert.Contains("<h1>Title</h1>", loaded.HtmlBody);
    }

    [Fact]
    public void SetBody_WithOutputPath_ShouldSaveToDifferentFile()
    {
        var path = CreateEmlFile("test_set_body_source.eml");
        var outputPath = CreateTestFilePath("test_set_body_output.eml");
        var result = _tool.Execute("set_body", path, outputPath, "Output body");

        Assert.NotNull(result);
        Assert.True(File.Exists(outputPath));
    }

    #endregion

    #region Set Headers

    [Fact]
    public void SetHeaders_ShouldSetHeader()
    {
        var path = CreateEmlFile("test_set_headers.eml");
        var result = _tool.Execute("set_headers", path, name: "X-Custom", value: "TestValue");

        Assert.NotNull(result);
        var data = GetResultData<SuccessResult>(result);
        Assert.Contains("X-Custom", data.Message);

        var loaded = MailMessage.Load(path);
        Assert.Equal("TestValue", loaded.Headers["X-Custom"]);
    }

    [Fact]
    public void SetHeaders_WithOutputPath_ShouldSaveToDifferentFile()
    {
        var path = CreateEmlFile("test_set_headers_source.eml");
        var outputPath = CreateTestFilePath("test_set_headers_output.eml");
        var result = _tool.Execute("set_headers", path, outputPath, name: "X-Out", value: "OutVal");

        Assert.NotNull(result);
        Assert.True(File.Exists(outputPath));
    }

    #endregion

    #region Get Subject

    [Fact]
    public void GetSubject_ShouldReturnSubject()
    {
        var path = CreateEmlFile("test_get_subject.eml", "My Subject");
        var result = _tool.Execute("get_subject", path);

        Assert.NotNull(result);
        var data = GetResultData<SuccessResult>(result);
        Assert.Contains("My Subject", data.Message);
    }

    [Fact]
    public void GetSubject_WithEmptySubject_ShouldReturnEmptyString()
    {
        var path = CreateEmlFile("test_empty_subject.eml", "");
        var result = _tool.Execute("get_subject", path);

        var data = GetResultData<SuccessResult>(result);
        Assert.Equal("", data.Message);
    }

    #endregion

    #region Set Subject

    [Fact]
    public void SetSubject_ShouldSetSubject()
    {
        var path = CreateEmlFile("test_set_subject.eml");
        var result = _tool.Execute("set_subject", path, subject: "Updated Subject");

        Assert.NotNull(result);
        var data = GetResultData<SuccessResult>(result);
        Assert.Contains("Updated Subject", data.Message);

        var loaded = MailMessage.Load(path);
        Assert.Contains("Updated Subject", loaded.Subject);
    }

    [Fact]
    public void SetSubject_WithOutputPath_ShouldSaveToDifferentFile()
    {
        var path = CreateEmlFile("test_set_subject_source.eml");
        var outputPath = CreateTestFilePath("test_set_subject_output.eml");
        var result = _tool.Execute("set_subject", path, outputPath, subject: "New Subject");

        Assert.NotNull(result);
        Assert.True(File.Exists(outputPath));

        var loaded = MailMessage.Load(outputPath);
        Assert.Contains("New Subject", loaded.Subject);
    }

    #endregion

    #region Set Recipients

    [Fact]
    public void SetRecipients_ShouldSetFrom()
    {
        var path = CreateEmlFile("test_set_from.eml");
        var result = _tool.Execute("set_recipients", path, from: "newfrom@example.com");

        Assert.NotNull(result);
        var data = GetResultData<SuccessResult>(result);
        Assert.Contains("recipients updated", data.Message);

        var loaded = MailMessage.Load(path);
        Assert.Equal("newfrom@example.com", loaded.From.Address);
    }

    [Fact]
    public void SetRecipients_ShouldSetTo()
    {
        var path = CreateEmlFile("test_set_to.eml");
        _tool.Execute("set_recipients", path, to: "a@b.com,c@d.com");

        var loaded = MailMessage.Load(path);
        Assert.Equal(2, loaded.To.Count);
    }

    [Fact]
    public void SetRecipients_ShouldSetCc()
    {
        var path = CreateEmlFile("test_set_cc.eml");
        _tool.Execute("set_recipients", path, cc: "cc@example.com");

        var loaded = MailMessage.Load(path);
        Assert.Single(loaded.CC);
        Assert.Equal("cc@example.com", loaded.CC[0].Address);
    }

    [Fact]
    public void SetRecipients_ShouldSetBcc()
    {
        var path = CreateEmlFile("test_set_bcc.eml");
        _tool.Execute("set_recipients", path, bcc: "bcc@example.com");

        var loaded = MailMessage.Load(path);
        Assert.Single(loaded.Bcc);
        Assert.Equal("bcc@example.com", loaded.Bcc[0].Address);
    }

    [Fact]
    public void SetRecipients_WithOutputPath_ShouldSaveToDifferentFile()
    {
        var path = CreateEmlFile("test_set_recipients_source.eml");
        var outputPath = CreateTestFilePath("test_set_recipients_output.eml");
        _tool.Execute("set_recipients", path, outputPath, to: "out@example.com");

        Assert.True(File.Exists(outputPath));
        var loaded = MailMessage.Load(outputPath);
        Assert.Contains(loaded.To, a => a.Address == "out@example.com");
    }

    #endregion

    #region Operation Routing

    [Theory]
    [InlineData("GET_BODY")]
    [InlineData("Get_Body")]
    [InlineData("get_body")]
    public void Operation_ShouldBeCaseInsensitive(string operation)
    {
        var path = CreateEmlFile($"test_case_{operation.Replace("_", "")}.eml", "Test", "Body");
        var result = _tool.Execute(operation, path);

        Assert.NotNull(result);
        var data = GetResultData<EmailBodyResult>(result);
        Assert.NotNull(data);
    }

    [Fact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        var path = CreateEmlFile("test_unknown_op.eml");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("unknown_operation", path));
        Assert.Contains("Unknown operation", ex.Message);
    }

    #endregion
}
