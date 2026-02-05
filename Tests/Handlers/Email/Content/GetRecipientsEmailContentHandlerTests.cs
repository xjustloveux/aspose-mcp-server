using Aspose.Email;
using AsposeMcpServer.Handlers.Email.Content;
using AsposeMcpServer.Results.Email.Content;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Email.Content;

/// <summary>
///     Tests for <see cref="GetRecipientsEmailContentHandler" />.
///     Verifies retrieval of From, To, CC, and BCC recipients from email files.
/// </summary>
public class GetRecipientsEmailContentHandlerTests : HandlerTestBase<object>
{
    private readonly GetRecipientsEmailContentHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_GetRecipients()
    {
        Assert.Equal("get_recipients", _handler.Operation);
    }

    #endregion

    #region Basic Operations

    [Fact]
    public void Execute_ReturnsFromAddress()
    {
        var path = CreateTestEmlFile("test_get_from.eml");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", path }
        });

        var result = _handler.Execute(context, parameters);

        var recipientsResult = Assert.IsType<EmailRecipientsResult>(result);
        Assert.Equal("sender@example.com", recipientsResult.From);
    }

    [Fact]
    public void Execute_ReturnsToRecipients()
    {
        var path = CreateTestEmlFile("test_get_to.eml");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", path }
        });

        var result = _handler.Execute(context, parameters);

        var recipientsResult = Assert.IsType<EmailRecipientsResult>(result);
        Assert.NotEmpty(recipientsResult.To);
        Assert.Contains("recipient@example.com", recipientsResult.To);
    }

    [Fact]
    public void Execute_WithMultipleToRecipients_ReturnsAll()
    {
        var path = CreateTestEmlFileWithMultipleRecipients("test_multiple_to.eml");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", path }
        });

        var result = _handler.Execute(context, parameters);

        var recipientsResult = Assert.IsType<EmailRecipientsResult>(result);
        Assert.True(recipientsResult.To.Count >= 2);
        Assert.Contains("to1@example.com", recipientsResult.To);
        Assert.Contains("to2@example.com", recipientsResult.To);
    }

    [Fact]
    public void Execute_WithCcRecipients_ReturnsCc()
    {
        var path = CreateTestEmlFileWithCcBcc("test_cc.eml");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", path }
        });

        var result = _handler.Execute(context, parameters);

        var recipientsResult = Assert.IsType<EmailRecipientsResult>(result);
        Assert.NotEmpty(recipientsResult.Cc);
        Assert.Contains("cc@example.com", recipientsResult.Cc);
    }

    [Fact]
    public void Execute_WithBccRecipients_ReturnsBcc()
    {
        var path = CreateTestEmlFileWithCcBcc("test_bcc.eml");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", path }
        });

        var result = _handler.Execute(context, parameters);

        var recipientsResult = Assert.IsType<EmailRecipientsResult>(result);
        Assert.NotEmpty(recipientsResult.Bcc);
        Assert.Contains("bcc@example.com", recipientsResult.Bcc);
    }

    [Fact]
    public void Execute_MessageContainsRecipientCount()
    {
        var path = CreateTestEmlFileWithCcBcc("test_count_message.eml");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", path }
        });

        var result = _handler.Execute(context, parameters);

        var recipientsResult = Assert.IsType<EmailRecipientsResult>(result);
        Assert.Contains("recipient(s)", recipientsResult.Message);
    }

    [Fact]
    public void Execute_WithNoOptionalRecipients_ReturnsEmptyLists()
    {
        var path = CreateTestEmlFileMinimal("test_minimal.eml");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", path }
        });

        var result = _handler.Execute(context, parameters);

        var recipientsResult = Assert.IsType<EmailRecipientsResult>(result);
        Assert.Empty(recipientsResult.Cc);
        Assert.Empty(recipientsResult.Bcc);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithMissingPath_ThrowsArgumentException()
    {
        var context = CreateContext(new object());
        var parameters = CreateEmptyParameters();

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithNonExistentFile_ThrowsFileNotFoundException()
    {
        var path = CreateTestFilePath("nonexistent.eml");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", path }
        });

        Assert.Throws<FileNotFoundException>(() => _handler.Execute(context, parameters));
    }

    #endregion

    #region Helper Methods

    /// <summary>
    ///     Creates a test EML file with a single To recipient.
    /// </summary>
    private string CreateTestEmlFile(string fileName)
    {
        var filePath = CreateTestFilePath(fileName);
        var message = new MailMessage
        {
            From = "sender@example.com",
            Subject = "Test Subject",
            Body = "Test Body"
        };
        message.To.Add("recipient@example.com");
        message.Save(filePath, SaveOptions.DefaultEml);
        return filePath;
    }

    /// <summary>
    ///     Creates a test EML file with multiple To recipients.
    /// </summary>
    private string CreateTestEmlFileWithMultipleRecipients(string fileName)
    {
        var filePath = CreateTestFilePath(fileName);
        var message = new MailMessage
        {
            From = "sender@example.com",
            Subject = "Test Subject",
            Body = "Test Body"
        };
        message.To.Add("to1@example.com");
        message.To.Add("to2@example.com");
        message.Save(filePath, SaveOptions.DefaultEml);
        return filePath;
    }

    /// <summary>
    ///     Creates a test EML file with To, CC, and BCC recipients.
    /// </summary>
    private string CreateTestEmlFileWithCcBcc(string fileName)
    {
        var filePath = CreateTestFilePath(fileName);
        var message = new MailMessage
        {
            From = "sender@example.com",
            Subject = "Test Subject",
            Body = "Test Body"
        };
        message.To.Add("recipient@example.com");
        message.CC.Add("cc@example.com");
        message.Bcc.Add("bcc@example.com");
        message.Save(filePath, SaveOptions.DefaultEml);
        return filePath;
    }

    /// <summary>
    ///     Creates a minimal test EML file with only From and To.
    /// </summary>
    private string CreateTestEmlFileMinimal(string fileName)
    {
        var filePath = CreateTestFilePath(fileName);
        var message = new MailMessage
        {
            From = "sender@example.com",
            Subject = "Test Subject",
            Body = "Test Body"
        };
        message.To.Add("recipient@example.com");
        message.Save(filePath, SaveOptions.DefaultEml);
        return filePath;
    }

    #endregion
}
