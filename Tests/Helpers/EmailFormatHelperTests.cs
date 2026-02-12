using Aspose.Email;
using AsposeMcpServer.Helpers;

namespace AsposeMcpServer.Tests.Helpers;

/// <summary>
///     Unit tests for EmailFormatHelper class.
/// </summary>
public class EmailFormatHelperTests
{
    #region DetermineEmailSaveFormat Tests

    [Theory]
    [InlineData(".eml", typeof(EmlSaveOptions))]
    [InlineData(".EML", typeof(EmlSaveOptions))]
    [InlineData(".Eml", typeof(EmlSaveOptions))]
    public void DetermineEmailSaveFormat_WithEmlExtension_ReturnsEmlSaveOptions(string extension, Type expectedType)
    {
        var result = EmailFormatHelper.DetermineEmailSaveFormat($"test{extension}");

        Assert.IsType(expectedType, result);
    }

    [Theory]
    [InlineData(".msg")]
    [InlineData(".MSG")]
    [InlineData(".Msg")]
    public void DetermineEmailSaveFormat_WithMsgExtension_ReturnsMsgSaveOptions(string extension)
    {
        var result = EmailFormatHelper.DetermineEmailSaveFormat($"test{extension}");

        Assert.IsType<MsgSaveOptions>(result);
    }

    [Theory]
    [InlineData(".mht")]
    [InlineData(".mhtml")]
    [InlineData(".MHT")]
    [InlineData(".MHTML")]
    public void DetermineEmailSaveFormat_WithMhtExtension_ReturnsMhtmlSaveOptions(string extension)
    {
        var result = EmailFormatHelper.DetermineEmailSaveFormat($"test{extension}");

        Assert.IsType<MhtSaveOptions>(result);
    }

    [Theory]
    [InlineData(".html")]
    [InlineData(".htm")]
    [InlineData(".HTML")]
    [InlineData(".HTM")]
    public void DetermineEmailSaveFormat_WithHtmlExtension_ReturnsHtmlSaveOptions(string extension)
    {
        var result = EmailFormatHelper.DetermineEmailSaveFormat($"test{extension}");

        Assert.IsType<HtmlSaveOptions>(result);
    }

    [Theory]
    [InlineData(".txt")]
    [InlineData(".pdf")]
    [InlineData(".docx")]
    [InlineData(".unknown")]
    [InlineData("")]
    public void DetermineEmailSaveFormat_WithUnknownExtension_ReturnsDefaultEml(string extension)
    {
        var result = EmailFormatHelper.DetermineEmailSaveFormat($"test{extension}");

        Assert.IsType<EmlSaveOptions>(result);
    }

    [Fact]
    public void DetermineEmailSaveFormat_WithFullPath_ExtractsExtensionCorrectly()
    {
        var result = EmailFormatHelper.DetermineEmailSaveFormat("/path/to/file.msg");

        Assert.IsType<MsgSaveOptions>(result);
    }

    [Fact]
    public void DetermineEmailSaveFormat_WithWindowsPath_ExtractsExtensionCorrectly()
    {
        var result = EmailFormatHelper.DetermineEmailSaveFormat(@"C:\Users\test\Documents\email.html");

        Assert.IsType<HtmlSaveOptions>(result);
    }

    #endregion
}
