using Aspose.Email;
using AsposeMcpServer.Handlers.Email.Content;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Email.Content;

/// <summary>
///     Tests for <see cref="SetSubjectEmailContentHandler" />.
///     Verifies setting the subject line on email files.
/// </summary>
public class SetSubjectEmailContentHandlerTests : HandlerTestBase<object>
{
    private readonly SetSubjectEmailContentHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_SetSubject()
    {
        Assert.Equal("set_subject", _handler.Operation);
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
            Subject = "Original Subject",
            Body = "Test Body"
        };
        message.To.Add("recipient@example.com");
        message.Save(filePath, SaveOptions.DefaultEml);
        return filePath;
    }

    #endregion

    #region Basic Operations

    [SkippableFact]
    public void Execute_SetsSubject()
    {
        SkipInEvaluationMode(AsposeLibraryType.Email, "Evaluation mode appends watermark to subject");
        var path = CreateTestEmlFile("test_set_subject.eml");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", path },
            { "subject", "New Subject" }
        });

        var result = _handler.Execute(context, parameters);

        var success = Assert.IsType<SuccessResult>(result);
        Assert.Contains("New Subject", success.Message);

        var loaded = MailMessage.Load(path);
        Assert.Equal("New Subject", loaded.Subject);
    }

    [SkippableFact]
    public void Execute_OverwritesExistingSubject()
    {
        SkipInEvaluationMode(AsposeLibraryType.Email, "Evaluation mode appends watermark to subject");
        var path = CreateTestEmlFile("test_overwrite_subject.eml");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", path },
            { "subject", "Replaced Subject" }
        });

        _handler.Execute(context, parameters);

        var loaded = MailMessage.Load(path);
        Assert.Equal("Replaced Subject", loaded.Subject);
    }

    [SkippableFact]
    public void Execute_WithOutputPath_SavesToDifferentFile()
    {
        SkipInEvaluationMode(AsposeLibraryType.Email, "Evaluation mode appends watermark to subject");
        var path = CreateTestEmlFile("test_set_subject_source.eml");
        var outputPath = CreateTestFilePath("test_set_subject_output.eml");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", path },
            { "subject", "Output Subject" },
            { "outputPath", outputPath }
        });

        var result = _handler.Execute(context, parameters);

        var success = Assert.IsType<SuccessResult>(result);
        Assert.Contains(outputPath, success.Message);
        Assert.True(File.Exists(outputPath));

        var loaded = MailMessage.Load(outputPath);
        Assert.Equal("Output Subject", loaded.Subject);
    }

    [Fact]
    public void Execute_WithEmptySubject_SetsEmptySubject()
    {
        var path = CreateTestEmlFile("test_set_empty_subject.eml");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", path },
            { "subject", "" }
        });

        _handler.Execute(context, parameters);

        var loaded = MailMessage.Load(path);
        Assert.Equal("", loaded.Subject);
    }

    [SkippableFact]
    public void Execute_DefaultsToOverwriteSourceFile()
    {
        SkipInEvaluationMode(AsposeLibraryType.Email, "Evaluation mode appends watermark to subject");
        var path = CreateTestEmlFile("test_set_subject_default.eml");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", path },
            { "subject", "Default Overwrite Subject" }
        });

        _handler.Execute(context, parameters);

        var loaded = MailMessage.Load(path);
        Assert.Equal("Default Overwrite Subject", loaded.Subject);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithMissingPath_ThrowsArgumentException()
    {
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "subject", "Some Subject" }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithMissingSubject_ThrowsArgumentException()
    {
        var path = CreateTestEmlFile("test_set_no_subject.eml");
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
            { "subject", "Some Subject" }
        });

        Assert.Throws<FileNotFoundException>(() => _handler.Execute(context, parameters));
    }

    #endregion
}
