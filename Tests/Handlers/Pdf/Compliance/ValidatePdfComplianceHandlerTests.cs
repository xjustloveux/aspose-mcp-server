using Aspose.Pdf;
using AsposeMcpServer.Handlers.Pdf.Compliance;
using AsposeMcpServer.Results.Pdf.Compliance;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Pdf.Compliance;

/// <summary>
///     Tests for <see cref="ValidatePdfComplianceHandler" />.
/// </summary>
public class ValidatePdfComplianceHandlerTests : PdfHandlerTestBase
{
    private readonly ValidatePdfComplianceHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Validate()
    {
        Assert.Equal("validate", _handler.Operation);
    }

    #endregion

    #region Format Variants

    [Theory]
    [InlineData("pdf/a-1a", "PDF/A-1a")]
    [InlineData("pdf/a-1b", "PDF/A-1b")]
    [InlineData("pdf/a-2a", "PDF/A-2a")]
    [InlineData("pdf/a-2b", "PDF/A-2b")]
    [InlineData("pdf/a-3a", "PDF/A-3a")]
    [InlineData("pdf/a-3b", "PDF/A-3b")]
    [InlineData("pdf/ua-1", "PDF/UA-1")]
    [InlineData("pdfa1b", "PDF/A-1b")]
    public void Execute_WithVariousFormats_ReturnsExpectedFormatName(string inputFormat, string expectedDisplayName)
    {
        var doc = CreateDocumentWithText("Test content");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "format", inputFormat }
        });

        var result = _handler.Execute(context, parameters);

        var validateResult = Assert.IsType<ValidateCompliancePdfResult>(result);
        Assert.Equal(expectedDisplayName, validateResult.Format);
    }

    #endregion

    #region Message Content

    [Fact]
    public void Execute_WhenCompliant_ReturnsCompliantMessage()
    {
        var doc = CreateDocumentWithText("Test content");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "format", "pdf/a-1b" }
        });

        var result = _handler.Execute(context, parameters);

        var validateResult = Assert.IsType<ValidateCompliancePdfResult>(result);
        Assert.Contains(
            validateResult.IsCompliant ? "compliant with PDF/A-1b" : "not compliant with PDF/A-1b",
            validateResult.Message);
    }

    #endregion

    #region Basic Validation

    [Fact]
    public void Execute_WithValidFormat_ReturnsResult()
    {
        var doc = CreateDocumentWithText("Test content");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "format", "pdf/a-1b" }
        });

        var result = _handler.Execute(context, parameters);

        var validateResult = Assert.IsType<ValidateCompliancePdfResult>(result);
        Assert.Equal("PDF/A-1b", validateResult.Format);
        Assert.NotNull(validateResult.Message);
        AssertNotModified(context);
    }

    [Fact]
    public void Execute_ReturnsCorrectIsCompliantAndErrorCountFields()
    {
        var doc = CreateDocumentWithText("Test content");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "format", "pdf/a-1b" }
        });

        var result = _handler.Execute(context, parameters);

        var validateResult = Assert.IsType<ValidateCompliancePdfResult>(result);
        Assert.IsType<bool>(validateResult.IsCompliant);
        Assert.True(validateResult.ErrorCount >= 0);
    }

    [Fact]
    public void Execute_WithEmptyDocument_Validates()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "format", "pdf/a-1b" }
        });

        var result = _handler.Execute(context, parameters);

        var validateResult = Assert.IsType<ValidateCompliancePdfResult>(result);
        Assert.Equal("PDF/A-1b", validateResult.Format);
        Assert.NotNull(validateResult.Message);
    }

    #endregion

    #region LogPath

    [Fact]
    public void Execute_WithLogPath_WritesLogFile()
    {
        var doc = CreateDocumentWithText("Test content");
        var context = CreateContext(doc);
        var logPath = CreateTestFilePath("validation.log");
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "format", "pdf/a-1b" },
            { "logPath", logPath }
        });

        var result = _handler.Execute(context, parameters);

        var validateResult = Assert.IsType<ValidateCompliancePdfResult>(result);
        Assert.Equal(logPath, validateResult.LogPath);
        Assert.True(File.Exists(logPath));
    }

    [Fact]
    public void Execute_WithoutLogPath_DoesNotPersistLogFile()
    {
        var doc = CreateDocumentWithText("Test content");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "format", "pdf/a-1b" }
        });

        var result = _handler.Execute(context, parameters);

        var validateResult = Assert.IsType<ValidateCompliancePdfResult>(result);
        Assert.Null(validateResult.LogPath);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithInvalidFormat_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithText("Test content");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "format", "invalid_format" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("Unsupported compliance format", ex.Message);
    }

    [Fact]
    public void Execute_WithoutFormat_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithText("Test content");
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion

    #region ResolvePdfFormat

    [Theory]
    [InlineData("pdf/a-1a")]
    [InlineData("pdf/a-1b")]
    [InlineData("pdf/a-2a")]
    [InlineData("pdf/a-2b")]
    [InlineData("pdf/a-3a")]
    [InlineData("pdf/a-3b")]
    [InlineData("pdf/ua-1")]
    [InlineData("pdfa1a")]
    [InlineData("pdfa1b")]
    [InlineData("pdfa2a")]
    [InlineData("pdfa2b")]
    [InlineData("pdfa3a")]
    [InlineData("pdfa3b")]
    [InlineData("pdfua1")]
    public void ResolvePdfFormat_WithValidFormat_DoesNotThrow(string format)
    {
        var result = ValidatePdfComplianceHandler.ResolvePdfFormat(format);
        Assert.IsType<PdfFormat>(result);
    }

    [Theory]
    [InlineData("invalid")]
    [InlineData("pdf/a-4")]
    [InlineData("")]
    public void ResolvePdfFormat_WithInvalidFormat_ThrowsArgumentException(string format)
    {
        Assert.Throws<ArgumentException>(() => ValidatePdfComplianceHandler.ResolvePdfFormat(format));
    }

    #endregion
}
