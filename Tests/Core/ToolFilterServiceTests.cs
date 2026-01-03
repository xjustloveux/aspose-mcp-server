using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Session;

namespace AsposeMcpServer.Tests.Core;

/// <summary>
///     Unit tests for ToolFilterService class
/// </summary>
public class ToolFilterServiceTests
{
    #region Helper Methods

    private static ServerConfig CreateServerConfig(
        bool enableWord = false,
        bool enableExcel = false,
        bool enablePowerPoint = false,
        bool enablePdf = false)
    {
        var args = new List<string>();

        if (enableWord) args.Add("--word");
        if (enableExcel) args.Add("--excel");
        if (enablePowerPoint) args.Add("--ppt");
        if (enablePdf) args.Add("--pdf");

        // If nothing enabled, we need to enable something to avoid validation error
        // But for testing, we want to test with nothing enabled
        if (args.Count == 0)
        {
            // Use environment variable to set tools to empty
            Environment.SetEnvironmentVariable("ASPOSE_TOOLS", "invalid");
            var config = ServerConfig.LoadFromArgs([]);
            Environment.SetEnvironmentVariable("ASPOSE_TOOLS", null);
            return config;
        }

        return ServerConfig.LoadFromArgs(args.ToArray());
    }

    #endregion

    #region Word Tool Tests

    [Fact]
    public void IsToolEnabled_WordTool_WhenWordEnabled_ShouldReturnTrue()
    {
        var serverConfig = CreateServerConfig(true);
        var sessionConfig = new SessionConfig();
        var service = new ToolFilterService(serverConfig, sessionConfig);

        Assert.True(service.IsToolEnabled("word_text"));
        Assert.True(service.IsToolEnabled("word_table"));
        Assert.True(service.IsToolEnabled("word_file"));
    }

    [Fact]
    public void IsToolEnabled_WordTool_WhenWordDisabled_ShouldReturnFalse()
    {
        var serverConfig = CreateServerConfig();
        var sessionConfig = new SessionConfig();
        var service = new ToolFilterService(serverConfig, sessionConfig);

        Assert.False(service.IsToolEnabled("word_text"));
        Assert.False(service.IsToolEnabled("word_table"));
    }

    #endregion

    #region Excel Tool Tests

    [Fact]
    public void IsToolEnabled_ExcelTool_WhenExcelEnabled_ShouldReturnTrue()
    {
        var serverConfig = CreateServerConfig(enableExcel: true);
        var sessionConfig = new SessionConfig();
        var service = new ToolFilterService(serverConfig, sessionConfig);

        Assert.True(service.IsToolEnabled("excel_cell"));
        Assert.True(service.IsToolEnabled("excel_chart"));
    }

    [Fact]
    public void IsToolEnabled_ExcelTool_WhenExcelDisabled_ShouldReturnFalse()
    {
        var serverConfig = CreateServerConfig(enableExcel: false);
        var sessionConfig = new SessionConfig();
        var service = new ToolFilterService(serverConfig, sessionConfig);

        Assert.False(service.IsToolEnabled("excel_cell"));
        Assert.False(service.IsToolEnabled("excel_chart"));
    }

    #endregion

    #region PowerPoint Tool Tests

    [Fact]
    public void IsToolEnabled_PptTool_WhenPowerPointEnabled_ShouldReturnTrue()
    {
        var serverConfig = CreateServerConfig(enablePowerPoint: true);
        var sessionConfig = new SessionConfig();
        var service = new ToolFilterService(serverConfig, sessionConfig);

        Assert.True(service.IsToolEnabled("ppt_slide"));
        Assert.True(service.IsToolEnabled("ppt_text"));
    }

    [Fact]
    public void IsToolEnabled_PptTool_WhenPowerPointDisabled_ShouldReturnFalse()
    {
        var serverConfig = CreateServerConfig(enablePowerPoint: false);
        var sessionConfig = new SessionConfig();
        var service = new ToolFilterService(serverConfig, sessionConfig);

        Assert.False(service.IsToolEnabled("ppt_slide"));
        Assert.False(service.IsToolEnabled("ppt_text"));
    }

    #endregion

    #region PDF Tool Tests

    [Fact]
    public void IsToolEnabled_PdfTool_WhenPdfEnabled_ShouldReturnTrue()
    {
        var serverConfig = CreateServerConfig(enablePdf: true);
        var sessionConfig = new SessionConfig();
        var service = new ToolFilterService(serverConfig, sessionConfig);

        Assert.True(service.IsToolEnabled("pdf_text"));
        Assert.True(service.IsToolEnabled("pdf_page"));
    }

    [Fact]
    public void IsToolEnabled_PdfTool_WhenPdfDisabled_ShouldReturnFalse()
    {
        var serverConfig = CreateServerConfig(enablePdf: false);
        var sessionConfig = new SessionConfig();
        var service = new ToolFilterService(serverConfig, sessionConfig);

        Assert.False(service.IsToolEnabled("pdf_text"));
        Assert.False(service.IsToolEnabled("pdf_page"));
    }

    #endregion

    #region Session Tool Tests

    [Fact]
    public void IsToolEnabled_SessionTool_WhenSessionEnabled_ShouldReturnTrue()
    {
        var serverConfig = CreateServerConfig();
        var sessionConfig = new SessionConfig { Enabled = true };
        var service = new ToolFilterService(serverConfig, sessionConfig);

        Assert.True(service.IsToolEnabled("document_session"));
    }

    [Fact]
    public void IsToolEnabled_SessionTool_WhenSessionDisabled_ShouldReturnFalse()
    {
        var serverConfig = CreateServerConfig();
        var sessionConfig = new SessionConfig { Enabled = false };
        var service = new ToolFilterService(serverConfig, sessionConfig);

        Assert.False(service.IsToolEnabled("document_session"));
    }

    #endregion

    #region Conversion Tool Tests

    [Fact]
    public void IsToolEnabled_ConvertToPdf_WhenAnyDocumentEnabled_ShouldReturnTrue()
    {
        var serverConfig = CreateServerConfig(true);
        var sessionConfig = new SessionConfig();
        var service = new ToolFilterService(serverConfig, sessionConfig);

        Assert.True(service.IsToolEnabled("convert_to_pdf"));
    }

    [Fact]
    public void IsToolEnabled_ConvertToPdf_WhenOnlyPdfEnabled_ShouldReturnFalse()
    {
        var serverConfig = CreateServerConfig(enablePdf: true);
        var sessionConfig = new SessionConfig();
        var service = new ToolFilterService(serverConfig, sessionConfig);

        Assert.False(service.IsToolEnabled("convert_to_pdf"));
    }

    [Fact]
    public void IsToolEnabled_ConvertDocument_WhenTwoOrMoreEnabled_ShouldReturnTrue()
    {
        var serverConfig = CreateServerConfig(true, true);
        var sessionConfig = new SessionConfig();
        var service = new ToolFilterService(serverConfig, sessionConfig);

        Assert.True(service.IsToolEnabled("convert_document"));
    }

    [Fact]
    public void IsToolEnabled_ConvertDocument_WhenOnlyOneEnabled_ShouldReturnFalse()
    {
        var serverConfig = CreateServerConfig(true);
        var sessionConfig = new SessionConfig();
        var service = new ToolFilterService(serverConfig, sessionConfig);

        Assert.False(service.IsToolEnabled("convert_document"));
    }

    #endregion

    #region GetEnabledCategories Tests

    [Fact]
    public void GetEnabledCategories_AllEnabled_ShouldReturnAll()
    {
        var serverConfig = CreateServerConfig(true, true, true, true);
        var sessionConfig = new SessionConfig { Enabled = true };
        var service = new ToolFilterService(serverConfig, sessionConfig);

        var result = service.GetEnabledCategories();

        Assert.Contains("Word", result);
        Assert.Contains("Excel", result);
        Assert.Contains("PowerPoint", result);
        Assert.Contains("PDF", result);
        Assert.Contains("Session", result);
    }

    [Fact]
    public void GetEnabledCategories_NoneEnabled_ShouldReturnNone()
    {
        var serverConfig = CreateServerConfig();
        var sessionConfig = new SessionConfig { Enabled = false };
        var service = new ToolFilterService(serverConfig, sessionConfig);

        var result = service.GetEnabledCategories();

        Assert.Equal("None", result);
    }

    [Fact]
    public void GetEnabledCategories_PartialEnabled_ShouldReturnEnabled()
    {
        var serverConfig = CreateServerConfig(true, enablePdf: true);
        var sessionConfig = new SessionConfig { Enabled = false };
        var service = new ToolFilterService(serverConfig, sessionConfig);

        var result = service.GetEnabledCategories();

        Assert.Contains("Word", result);
        Assert.Contains("PDF", result);
        Assert.DoesNotContain("Excel", result);
        Assert.DoesNotContain("PowerPoint", result);
        Assert.DoesNotContain("Session", result);
    }

    #endregion

    #region Edge Cases

    [Fact]
    public void IsToolEnabled_NullOrEmpty_ShouldReturnFalse()
    {
        var serverConfig = CreateServerConfig(true, true, true, true);
        var sessionConfig = new SessionConfig { Enabled = true };
        var service = new ToolFilterService(serverConfig, sessionConfig);

        Assert.False(service.IsToolEnabled(null!));
        Assert.False(service.IsToolEnabled(""));
    }

    [Fact]
    public void IsToolEnabled_UnknownTool_ShouldReturnTrue()
    {
        var serverConfig = CreateServerConfig();
        var sessionConfig = new SessionConfig();
        var service = new ToolFilterService(serverConfig, sessionConfig);

        Assert.True(service.IsToolEnabled("unknown_tool"));
    }

    [Fact]
    public void IsToolEnabled_CaseInsensitive_ShouldWork()
    {
        var serverConfig = CreateServerConfig(true);
        var sessionConfig = new SessionConfig();
        var service = new ToolFilterService(serverConfig, sessionConfig);

        Assert.True(service.IsToolEnabled("WORD_text"));
        Assert.True(service.IsToolEnabled("Word_Text"));
    }

    #endregion
}