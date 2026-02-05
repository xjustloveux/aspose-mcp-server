using Aspose.Email;
using AsposeMcpServer.Handlers.Email.Content;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Email.Content;

/// <summary>
///     Tests for <see cref="GetSubjectEmailContentHandler" />.
///     Verifies retrieval of the subject line from email files.
/// </summary>
public class GetSubjectEmailContentHandlerTests : HandlerTestBase<object>
{
    private readonly GetSubjectEmailContentHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_GetSubject()
    {
        Assert.Equal("get_subject", _handler.Operation);
    }

    #endregion

    #region Helper Methods

    /// <summary>
    ///     Creates a test EML file with the specified subject.
    /// </summary>
    private string CreateTestEmlFile(string fileName, string subject)
    {
        var filePath = CreateTestFilePath(fileName);
        var message = new MailMessage
        {
            From = "sender@example.com",
            Subject = subject,
            Body = "Test Body"
        };
        message.To.Add("recipient@example.com");
        message.Save(filePath, SaveOptions.DefaultEml);
        return filePath;
    }

    #endregion

    #region Basic Operations

    [SkippableFact]
    public void Execute_ReturnsSubject()
    {
        SkipInEvaluationMode(AsposeLibraryType.Email, "Evaluation mode appends watermark to subject");
        var path = CreateTestEmlFile("test_get_subject.eml", "My Test Subject");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", path }
        });

        var result = _handler.Execute(context, parameters);

        var success = Assert.IsType<SuccessResult>(result);
        Assert.Equal("My Test Subject", success.Message);
    }

    [Fact]
    public void Execute_WithEmptySubject_ReturnsEmptyString()
    {
        var path = CreateTestEmlFile("test_empty_subject.eml", "");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", path }
        });

        var result = _handler.Execute(context, parameters);

        var success = Assert.IsType<SuccessResult>(result);
        Assert.Equal("", success.Message);
    }

    [Fact]
    public void Execute_WithSpecialCharactersInSubject_ReturnsSubjectCorrectly()
    {
        var path = CreateTestEmlFile("test_special_subject.eml", "Re: [URGENT] Hello & Goodbye <test>");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", path }
        });

        var result = _handler.Execute(context, parameters);

        var success = Assert.IsType<SuccessResult>(result);
        Assert.Contains("Re:", success.Message);
        Assert.Contains("[URGENT]", success.Message);
    }

    [SkippableFact]
    public void Execute_WithUnicodeSubject_ReturnsSubjectCorrectly()
    {
        SkipInEvaluationMode(AsposeLibraryType.Email, "Evaluation mode appends watermark to subject");
        var path = CreateTestEmlFile("test_unicode_subject.eml", "Test Unicode Subject");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", path }
        });

        var result = _handler.Execute(context, parameters);

        var success = Assert.IsType<SuccessResult>(result);
        Assert.Equal("Test Unicode Subject", success.Message);
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
}
