using Aspose.Email;
using AsposeMcpServer.Handlers.Email.Content;
using AsposeMcpServer.Results.Email.Content;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Email.Content;

/// <summary>
///     Tests for <see cref="GetHeadersEmailContentHandler" />.
///     Verifies retrieval of email headers from EML files.
/// </summary>
public class GetHeadersEmailContentHandlerTests : HandlerTestBase<object>
{
    private readonly GetHeadersEmailContentHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_GetHeaders()
    {
        Assert.Equal("get_headers", _handler.Operation);
    }

    #endregion

    #region Basic Operations

    [Fact]
    public void Execute_ReturnsHeaders()
    {
        var path = CreateTestEmlFile("test_get_headers.eml");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", path }
        });

        var result = _handler.Execute(context, parameters);

        var headersResult = Assert.IsType<EmailHeadersResult>(result);
        Assert.NotNull(headersResult.Headers);
        Assert.True(headersResult.Count > 0);
        Assert.Equal(headersResult.Headers.Count, headersResult.Count);
        Assert.Contains("header(s)", headersResult.Message);
    }

    [Fact]
    public void Execute_HeadersContainStandardFields()
    {
        var path = CreateTestEmlFile("test_standard_headers.eml");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", path }
        });

        var result = _handler.Execute(context, parameters);

        var headersResult = Assert.IsType<EmailHeadersResult>(result);
        // ReSharper disable once ParameterOnlyUsedForPreconditionCheck.Local - Assert.All parameter is intended for validation
        Assert.All(headersResult.Headers, h =>
        {
            Assert.NotNull(h.Name);
            Assert.NotNull(h.Value);
        });
    }

    [Fact]
    public void Execute_WithCustomHeader_ReturnsCustomHeader()
    {
        var path = CreateTestEmlFileWithCustomHeader("test_custom_header.eml", "X-Custom-Header", "CustomValue");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", path }
        });

        var result = _handler.Execute(context, parameters);

        var headersResult = Assert.IsType<EmailHeadersResult>(result);
        Assert.Contains(headersResult.Headers, h => h is { Name: "X-Custom-Header", Value: "CustomValue" });
    }

    [Fact]
    public void Execute_CountMatchesHeadersListLength()
    {
        var path = CreateTestEmlFile("test_count_match.eml");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", path }
        });

        var result = _handler.Execute(context, parameters);

        var headersResult = Assert.IsType<EmailHeadersResult>(result);
        Assert.Equal(headersResult.Headers.Count, headersResult.Count);
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
    ///     Creates a test EML file with default content.
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
    ///     Creates a test EML file with a custom header.
    /// </summary>
    private string CreateTestEmlFileWithCustomHeader(string fileName, string headerName, string headerValue)
    {
        var filePath = CreateTestFilePath(fileName);
        var message = new MailMessage
        {
            From = "sender@example.com",
            Subject = "Test Subject",
            Body = "Test Body"
        };
        message.To.Add("recipient@example.com");
        message.Headers.Add(headerName, headerValue);
        message.Save(filePath, SaveOptions.DefaultEml);
        return filePath;
    }

    #endregion
}
