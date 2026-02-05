using Aspose.Email;
using AsposeMcpServer.Handlers.Email.Content;
using AsposeMcpServer.Results.Email.Content;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Email.Content;

/// <summary>
///     Tests for <see cref="GetBodyEmailContentHandler" />.
///     Verifies retrieval of email body content from EML files.
/// </summary>
public class GetBodyEmailContentHandlerTests : HandlerTestBase<object>
{
    private readonly GetBodyEmailContentHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_GetBody()
    {
        Assert.Equal("get_body", _handler.Operation);
    }

    #endregion

    #region Basic Operations

    [Fact]
    public void Execute_WithPlainTextBody_ReturnsBody()
    {
        var path = CreateTestEmlFile("test_plain.eml", "Test Subject", "Plain text body");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", path }
        });

        var result = _handler.Execute(context, parameters);

        var bodyResult = Assert.IsType<EmailBodyResult>(result);
        Assert.Contains("Plain text body", bodyResult.Body);
        Assert.False(bodyResult.IsHtml);
        Assert.Contains("plain text", bodyResult.Message);
    }

    [Fact]
    public void Execute_WithHtmlBody_ReturnsHtmlBody()
    {
        var path = CreateTestEmlFileWithHtmlBody("test_html.eml", "HTML Subject", "<h1>Hello</h1><p>World</p>");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", path }
        });

        var result = _handler.Execute(context, parameters);

        var bodyResult = Assert.IsType<EmailBodyResult>(result);
        Assert.NotNull(bodyResult.HtmlBody);
        Assert.Contains("<h1>Hello</h1>", bodyResult.HtmlBody);
        Assert.True(bodyResult.IsHtml);
        Assert.Contains("HTML", bodyResult.Message);
    }

    [Fact]
    public void Execute_WithEmptyBody_ReturnsEmptyContent()
    {
        var path = CreateTestEmlFile("test_empty.eml", "Empty Subject", "");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", path }
        });

        var result = _handler.Execute(context, parameters);

        var bodyResult = Assert.IsType<EmailBodyResult>(result);
        Assert.NotNull(bodyResult);
        Assert.False(bodyResult.IsHtml);
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
    ///     Creates a test EML file with a plain text body.
    /// </summary>
    private string CreateTestEmlFile(string fileName, string subject, string body)
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

    /// <summary>
    ///     Creates a test EML file with an HTML body.
    /// </summary>
    private string CreateTestEmlFileWithHtmlBody(string fileName, string subject, string htmlBody)
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
}
