using Aspose.Email;
using AsposeMcpServer.Handlers.Email.Content;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Email.Content;

/// <summary>
///     Tests for <see cref="SetRecipientsEmailContentHandler" />.
///     Verifies setting From, To, CC, and BCC recipients on email files.
/// </summary>
public class SetRecipientsEmailContentHandlerTests : HandlerTestBase<object>
{
    private readonly SetRecipientsEmailContentHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_SetRecipients()
    {
        Assert.Equal("set_recipients", _handler.Operation);
    }

    #endregion

    #region Basic Operations

    [Fact]
    public void Execute_SetsFromAddress()
    {
        var path = CreateTestEmlFile("test_set_from.eml");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", path },
            { "from", "newsender@example.com" }
        });

        var result = _handler.Execute(context, parameters);

        var success = Assert.IsType<SuccessResult>(result);
        Assert.Contains("recipients updated", success.Message);

        var loaded = MailMessage.Load(path);
        Assert.Equal("newsender@example.com", loaded.From.Address);
    }

    [Fact]
    public void Execute_SetsToRecipients()
    {
        var path = CreateTestEmlFile("test_set_to.eml");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", path },
            { "to", "newto@example.com" }
        });

        _handler.Execute(context, parameters);

        var loaded = MailMessage.Load(path);
        Assert.Single(loaded.To);
        Assert.Equal("newto@example.com", loaded.To[0].Address);
    }

    [Fact]
    public void Execute_SetsMultipleToRecipients()
    {
        var path = CreateTestEmlFile("test_set_multi_to.eml");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", path },
            { "to", "to1@example.com,to2@example.com,to3@example.com" }
        });

        _handler.Execute(context, parameters);

        var loaded = MailMessage.Load(path);
        Assert.Equal(3, loaded.To.Count);
        Assert.Contains(loaded.To, a => a.Address == "to1@example.com");
        Assert.Contains(loaded.To, a => a.Address == "to2@example.com");
        Assert.Contains(loaded.To, a => a.Address == "to3@example.com");
    }

    [Fact]
    public void Execute_SetsCcRecipients()
    {
        var path = CreateTestEmlFile("test_set_cc.eml");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", path },
            { "cc", "cc1@example.com,cc2@example.com" }
        });

        _handler.Execute(context, parameters);

        var loaded = MailMessage.Load(path);
        Assert.Equal(2, loaded.CC.Count);
        Assert.Contains(loaded.CC, a => a.Address == "cc1@example.com");
        Assert.Contains(loaded.CC, a => a.Address == "cc2@example.com");
    }

    [Fact]
    public void Execute_SetsBccRecipients()
    {
        var path = CreateTestEmlFile("test_set_bcc.eml");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", path },
            { "bcc", "bcc1@example.com" }
        });

        _handler.Execute(context, parameters);

        var loaded = MailMessage.Load(path);
        Assert.Single(loaded.Bcc);
        Assert.Equal("bcc1@example.com", loaded.Bcc[0].Address);
    }

    [Fact]
    public void Execute_SetsAllRecipientsAtOnce()
    {
        var path = CreateTestEmlFile("test_set_all.eml");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", path },
            { "from", "allfrom@example.com" },
            { "to", "allto@example.com" },
            { "cc", "allcc@example.com" },
            { "bcc", "allbcc@example.com" }
        });

        _handler.Execute(context, parameters);

        var loaded = MailMessage.Load(path);
        Assert.Equal("allfrom@example.com", loaded.From.Address);
        Assert.Contains(loaded.To, a => a.Address == "allto@example.com");
        Assert.Contains(loaded.CC, a => a.Address == "allcc@example.com");
        Assert.Contains(loaded.Bcc, a => a.Address == "allbcc@example.com");
    }

    [Fact]
    public void Execute_ClearsExistingToBeforeSettingNew()
    {
        var path = CreateTestEmlFileWithMultipleTo("test_clear_to.eml");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", path },
            { "to", "newonly@example.com" }
        });

        _handler.Execute(context, parameters);

        var loaded = MailMessage.Load(path);
        Assert.Single(loaded.To);
        Assert.Equal("newonly@example.com", loaded.To[0].Address);
    }

    [Fact]
    public void Execute_WithOutputPath_SavesToDifferentFile()
    {
        var path = CreateTestEmlFile("test_set_recipients_source.eml");
        var outputPath = CreateTestFilePath("test_set_recipients_output.eml");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", path },
            { "to", "output@example.com" },
            { "outputPath", outputPath }
        });

        var result = _handler.Execute(context, parameters);

        var success = Assert.IsType<SuccessResult>(result);
        Assert.Contains(outputPath, success.Message);
        Assert.True(File.Exists(outputPath));

        var loaded = MailMessage.Load(outputPath);
        Assert.Contains(loaded.To, a => a.Address == "output@example.com");
    }

    [Fact]
    public void Execute_WithNoOptionalParams_PreservesExistingRecipients()
    {
        var path = CreateTestEmlFile("test_no_change.eml");
        var originalLoaded = MailMessage.Load(path);
        var originalToCount = originalLoaded.To.Count;

        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", path }
        });

        _handler.Execute(context, parameters);

        var loaded = MailMessage.Load(path);
        Assert.Equal(originalToCount, loaded.To.Count);
    }

    [Fact]
    public void Execute_WithCommaSeparatedAndSpaces_TrimsAddresses()
    {
        var path = CreateTestEmlFile("test_trim_addresses.eml");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", path },
            { "to", "  a@b.com , c@d.com , e@f.com  " }
        });

        _handler.Execute(context, parameters);

        var loaded = MailMessage.Load(path);
        Assert.Equal(3, loaded.To.Count);
        Assert.Contains(loaded.To, a => a.Address == "a@b.com");
        Assert.Contains(loaded.To, a => a.Address == "c@d.com");
        Assert.Contains(loaded.To, a => a.Address == "e@f.com");
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithMissingPath_ThrowsArgumentException()
    {
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "to", "test@example.com" }
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
            { "to", "test@example.com" }
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
    ///     Creates a test EML file with multiple To recipients.
    /// </summary>
    private string CreateTestEmlFileWithMultipleTo(string fileName)
    {
        var filePath = CreateTestFilePath(fileName);
        var message = new MailMessage
        {
            From = "sender@example.com",
            Subject = "Test Subject",
            Body = "Test Body"
        };
        message.To.Add("old1@example.com");
        message.To.Add("old2@example.com");
        message.Save(filePath, SaveOptions.DefaultEml);
        return filePath;
    }

    #endregion
}
