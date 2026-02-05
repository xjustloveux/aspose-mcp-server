using Aspose.Email;
using AsposeMcpServer.Handlers.Email.Content;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Email.Content;

/// <summary>
///     Tests for <see cref="SetHeadersEmailContentHandler" />.
///     Verifies setting and updating email headers on EML files.
/// </summary>
public class SetHeadersEmailContentHandlerTests : HandlerTestBase<object>
{
    private readonly SetHeadersEmailContentHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_SetHeaders()
    {
        Assert.Equal("set_headers", _handler.Operation);
    }

    #endregion

    #region Basic Operations

    [Fact]
    public void Execute_SetsNewHeader()
    {
        var path = CreateTestEmlFile("test_set_new_header.eml");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", path },
            { "name", "X-Custom-Header" },
            { "value", "TestValue" }
        });

        var result = _handler.Execute(context, parameters);

        var success = Assert.IsType<SuccessResult>(result);
        Assert.Contains("X-Custom-Header", success.Message);
        Assert.Contains("TestValue", success.Message);

        var loaded = MailMessage.Load(path);
        Assert.Equal("TestValue", loaded.Headers["X-Custom-Header"]);
    }

    [Fact]
    public void Execute_UpdatesExistingHeader()
    {
        var path = CreateTestEmlFileWithHeader("test_update_header.eml", "X-Existing", "OldValue");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", path },
            { "name", "X-Existing" },
            { "value", "NewValue" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(result);
        var loaded = MailMessage.Load(path);
        Assert.Equal("NewValue", loaded.Headers["X-Existing"]);
    }

    [Fact]
    public void Execute_WithOutputPath_SavesToDifferentFile()
    {
        var path = CreateTestEmlFile("test_set_header_source.eml");
        var outputPath = CreateTestFilePath("test_set_header_output.eml");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", path },
            { "name", "X-New-Header" },
            { "value", "OutputValue" },
            { "outputPath", outputPath }
        });

        var result = _handler.Execute(context, parameters);

        var success = Assert.IsType<SuccessResult>(result);
        Assert.Contains(outputPath, success.Message);
        Assert.True(File.Exists(outputPath));

        var loaded = MailMessage.Load(outputPath);
        Assert.Equal("OutputValue", loaded.Headers["X-New-Header"]);
    }

    [Fact]
    public void Execute_DefaultsToOverwriteSourceFile()
    {
        var path = CreateTestEmlFile("test_set_header_default.eml");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", path },
            { "name", "X-Overwrite" },
            { "value", "OverwriteValue" }
        });

        _handler.Execute(context, parameters);

        var loaded = MailMessage.Load(path);
        Assert.Equal("OverwriteValue", loaded.Headers["X-Overwrite"]);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithMissingPath_ThrowsArgumentException()
    {
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "name", "X-Test" },
            { "value", "TestValue" }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithMissingName_ThrowsArgumentException()
    {
        var path = CreateTestEmlFile("test_set_no_name.eml");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", path },
            { "value", "TestValue" }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithMissingValue_ThrowsArgumentException()
    {
        var path = CreateTestEmlFile("test_set_no_value.eml");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", path },
            { "name", "X-Test" }
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
            { "name", "X-Test" },
            { "value", "TestValue" }
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
    ///     Creates a test EML file with a pre-set custom header.
    /// </summary>
    private string CreateTestEmlFileWithHeader(string fileName, string headerName, string headerValue)
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
