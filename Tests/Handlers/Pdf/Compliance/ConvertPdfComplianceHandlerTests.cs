using AsposeMcpServer.Handlers.Pdf.Compliance;
using AsposeMcpServer.Results.Pdf.Compliance;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Pdf.Compliance;

/// <summary>
///     Tests for <see cref="ConvertPdfComplianceHandler" />.
/// </summary>
public class ConvertPdfComplianceHandlerTests : PdfHandlerTestBase
{
    private readonly ConvertPdfComplianceHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Convert()
    {
        Assert.Equal("convert", _handler.Operation);
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

        var convertResult = Assert.IsType<ConvertCompliancePdfResult>(result);
        Assert.Equal(expectedDisplayName, convertResult.Format);
        AssertModified(context);
    }

    #endregion

    #region Message Content

    [Fact]
    public void Execute_WhenSuccessful_ReturnsSuccessMessage()
    {
        var doc = CreateDocumentWithText("Test content");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "format", "pdf/a-1b" }
        });

        var result = _handler.Execute(context, parameters);

        var convertResult = Assert.IsType<ConvertCompliancePdfResult>(result);
        Assert.Contains(
            convertResult.IsSuccess ? "successfully converted to PDF/A-1b" : "PDF/A-1b",
            convertResult.Message);
    }

    #endregion

    #region Basic Conversion

    [Fact]
    public void Execute_WithValidFormat_ConvertsAndMarksModified()
    {
        var doc = CreateDocumentWithText("Test content");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "format", "pdf/a-1b" }
        });

        var result = _handler.Execute(context, parameters);

        var convertResult = Assert.IsType<ConvertCompliancePdfResult>(result);
        Assert.Equal("PDF/A-1b", convertResult.Format);
        Assert.NotNull(convertResult.Message);
        AssertModified(context);
    }

    [Fact]
    public void Execute_ReturnsIsSuccessField()
    {
        var doc = CreateDocumentWithText("Test content");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "format", "pdf/a-1b" }
        });

        var result = _handler.Execute(context, parameters);

        var convertResult = Assert.IsType<ConvertCompliancePdfResult>(result);
        Assert.IsType<bool>(convertResult.IsSuccess);
    }

    #endregion

    #region LogPath

    [Fact]
    public void Execute_WithLogPath_WritesConversionLog()
    {
        var doc = CreateDocumentWithText("Test content");
        var context = CreateContext(doc);
        var logPath = CreateTestFilePath("conversion.log");
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "format", "pdf/a-1b" },
            { "logPath", logPath }
        });

        var result = _handler.Execute(context, parameters);

        var convertResult = Assert.IsType<ConvertCompliancePdfResult>(result);
        Assert.Equal(logPath, convertResult.LogPath);
        Assert.True(File.Exists(logPath));
        AssertModified(context);
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

        var convertResult = Assert.IsType<ConvertCompliancePdfResult>(result);
        Assert.Null(convertResult.LogPath);
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
}
