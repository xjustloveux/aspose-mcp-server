using Aspose.Email;
using AsposeMcpServer.Handlers.Email.Content;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Email.Content;

/// <summary>
///     Tests for <see cref="SetBodyEmailContentHandler" />.
///     Verifies setting plain text and HTML body content on email files.
/// </summary>
public class SetBodyEmailContentHandlerTests : HandlerTestBase<object>
{
    private readonly SetBodyEmailContentHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_SetBody()
    {
        Assert.Equal("set_body", _handler.Operation);
    }

    #endregion

    #region Helper Methods

    /// <summary>
    ///     Creates a test EML file with default content.
    /// </summary>
    private string CreateTestEmlFile(string fileName)
    {
        var filePath = CreateTestFilePath(fileName);
        var message = new MailMessage
        {
            From = "sender@example.com",
            Subject = "Test Subject",
            Body = "Original body"
        };
        message.To.Add("recipient@example.com");
        message.Save(filePath, SaveOptions.DefaultEml);
        return filePath;
    }

    #endregion

    #region Basic Operations

    [Fact]
    public void Execute_SetsPlainTextBody()
    {
        var path = CreateTestEmlFile("test_set_plain.eml");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", path },
            { "body", "New plain text body" }
        });

        var result = _handler.Execute(context, parameters);

        var success = Assert.IsType<SuccessResult>(result);
        Assert.Contains("body updated", success.Message);

        var loaded = MailMessage.Load(path);
        Assert.Contains("New plain text body", loaded.Body);
    }

    [Fact]
    public void Execute_SetsHtmlBody()
    {
        var path = CreateTestEmlFile("test_set_html.eml");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", path },
            { "body", "<h1>New HTML</h1>" },
            { "isHtml", true }
        });

        var result = _handler.Execute(context, parameters);

        var success = Assert.IsType<SuccessResult>(result);
        Assert.Contains("HTML body updated", success.Message);

        var loaded = MailMessage.Load(path);
        Assert.Contains("<h1>New HTML</h1>", loaded.HtmlBody);
    }

    [Fact]
    public void Execute_WithOutputPath_SavesToDifferentFile()
    {
        var path = CreateTestEmlFile("test_set_body_source.eml");
        var outputPath = CreateTestFilePath("test_set_body_output.eml");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", path },
            { "body", "Updated body" },
            { "outputPath", outputPath }
        });

        var result = _handler.Execute(context, parameters);

        var success = Assert.IsType<SuccessResult>(result);
        Assert.Contains(outputPath, success.Message);
        Assert.True(File.Exists(outputPath));

        var loaded = MailMessage.Load(outputPath);
        Assert.Contains("Updated body", loaded.Body);
    }

    [Fact]
    public void Execute_WithIsHtmlFalse_SetsPlainTextBody()
    {
        var path = CreateTestEmlFile("test_set_not_html.eml");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", path },
            { "body", "Plain text content" },
            { "isHtml", false }
        });

        var result = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(result);
        var loaded = MailMessage.Load(path);
        Assert.Contains("Plain text content", loaded.Body);
    }

    [Fact]
    public void Execute_DefaultsToOverwriteSourceFile()
    {
        var path = CreateTestEmlFile("test_set_body_default_output.eml");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", path },
            { "body", "Overwritten body" }
        });

        _handler.Execute(context, parameters);

        var loaded = MailMessage.Load(path);
        Assert.Contains("Overwritten body", loaded.Body);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithMissingPath_ThrowsArgumentException()
    {
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "body", "Some body" }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithMissingBody_ThrowsArgumentException()
    {
        var path = CreateTestEmlFile("test_set_no_body.eml");
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
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", path },
            { "body", "Some body" }
        });

        Assert.Throws<FileNotFoundException>(() => _handler.Execute(context, parameters));
    }

    #endregion
}
