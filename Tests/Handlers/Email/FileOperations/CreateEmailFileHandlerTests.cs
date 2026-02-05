using Aspose.Email;
using AsposeMcpServer.Handlers.Email.FileOperations;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Email.FileOperations;

/// <summary>
///     Tests for <see cref="CreateEmailFileHandler" />.
///     Verifies email file creation with various parameters and format detection.
/// </summary>
public class CreateEmailFileHandlerTests : HandlerTestBase<object>
{
    private readonly CreateEmailFileHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Create()
    {
        Assert.Equal("create", _handler.Operation);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutOutputPath_ThrowsArgumentException()
    {
        var context = CreateContext(new object());
        var parameters = CreateEmptyParameters();

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion

    #region Parameter Validation

    [Fact]
    public void Execute_WithEmptyTo_DoesNotAddRecipient()
    {
        var outputPath = CreateTestFilePath("test_empty_to.eml");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "outputPath", outputPath },
            { "to", "" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(result);
        var loaded = MailMessage.Load(outputPath);
        Assert.Empty(loaded.To);
    }

    #endregion

    #region Basic Operations

    [Fact]
    public void Execute_CreatesEmlFileWithDefaults()
    {
        var outputPath = CreateTestFilePath("test_create.eml");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "outputPath", outputPath }
        });

        var result = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(result);
        var success = (SuccessResult)result;
        Assert.Contains(outputPath, success.Message);
        Assert.True(File.Exists(outputPath));

        var loaded = MailMessage.Load(outputPath);
        Assert.Equal("noreply@example.com", loaded.From.Address);
        Assert.Equal("", loaded.Subject);
    }

    [SkippableFact]
    public void Execute_CreatesEmlFileWithAllParameters()
    {
        SkipInEvaluationMode(AsposeLibraryType.Email, "Evaluation mode appends watermark to subject");
        var outputPath = CreateTestFilePath("test_full.eml");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "outputPath", outputPath },
            { "subject", "Hello World" },
            { "body", "This is a test email." },
            { "from", "sender@example.com" },
            { "to", "recipient@example.com" },
            { "isHtml", false }
        });

        var result = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(result);
        Assert.True(File.Exists(outputPath));

        var loaded = MailMessage.Load(outputPath);
        Assert.Equal("Hello World", loaded.Subject);
        Assert.Equal("sender@example.com", loaded.From.Address);
        Assert.Contains(loaded.To, addr => addr.Address == "recipient@example.com");
        Assert.Contains("This is a test email.", loaded.Body);
    }

    [Fact]
    public void Execute_CreatesHtmlEmail()
    {
        var outputPath = CreateTestFilePath("test_html.eml");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "outputPath", outputPath },
            { "subject", "HTML Email" },
            { "body", "<h1>Hello</h1><p>World</p>" },
            { "isHtml", true }
        });

        var result = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(result);
        Assert.True(File.Exists(outputPath));

        var loaded = MailMessage.Load(outputPath);
        Assert.Contains("<h1>Hello</h1>", loaded.HtmlBody);
    }

    [SkippableFact]
    public void Execute_CreatesMsgFile()
    {
        SkipInEvaluationMode(AsposeLibraryType.Email);
        var outputPath = CreateTestFilePath("test_create.msg");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "outputPath", outputPath },
            { "subject", "MSG Test" },
            { "from", "sender@example.com" },
            { "to", "recipient@example.com" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(result);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void Execute_CreatesMhtmlFile()
    {
        var outputPath = CreateTestFilePath("test_create.mhtml");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "outputPath", outputPath },
            { "subject", "MHTML Test" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(result);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void Execute_CreatesHtmlFormatFile()
    {
        var outputPath = CreateTestFilePath("test_create.html");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "outputPath", outputPath },
            { "subject", "HTML Format Test" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(result);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void Execute_WithoutTo_CreatesEmailWithoutRecipient()
    {
        var outputPath = CreateTestFilePath("test_no_to.eml");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "outputPath", outputPath },
            { "subject", "No Recipient" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(result);
        Assert.True(File.Exists(outputPath));

        var loaded = MailMessage.Load(outputPath);
        Assert.Empty(loaded.To);
    }

    [Fact]
    public void Execute_WithNullSubjectAndBody_UsesDefaults()
    {
        var outputPath = CreateTestFilePath("test_null_params.eml");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "outputPath", outputPath }
        });

        var result = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(result);
        var loaded = MailMessage.Load(outputPath);
        Assert.Equal("", loaded.Subject);
    }

    #endregion

    #region DetectSaveOptions

    [Theory]
    [InlineData("test.msg")]
    [InlineData("test.MSG")]
    public void DetectSaveOptions_ForMsg_ReturnsMsgUnicode(string path)
    {
        var options = CreateEmailFileHandler.DetectSaveOptions(path);
        Assert.IsType<MsgSaveOptions>(options);
    }

    [Theory]
    [InlineData("test.mhtml")]
    [InlineData("test.mht")]
    public void DetectSaveOptions_ForMhtml_ReturnsMhtml(string path)
    {
        var options = CreateEmailFileHandler.DetectSaveOptions(path);
        Assert.IsType<MhtSaveOptions>(options);
    }

    [Theory]
    [InlineData("test.html")]
    [InlineData("test.htm")]
    public void DetectSaveOptions_ForHtml_ReturnsHtml(string path)
    {
        var options = CreateEmailFileHandler.DetectSaveOptions(path);
        Assert.IsType<HtmlSaveOptions>(options);
    }

    [Theory]
    [InlineData("test.eml")]
    [InlineData("test.xyz")]
    [InlineData("test.txt")]
    public void DetectSaveOptions_ForEmlOrUnknown_ReturnsDefaultEml(string path)
    {
        var options = CreateEmailFileHandler.DetectSaveOptions(path);
        Assert.IsType<EmlSaveOptions>(options);
    }

    #endregion
}
