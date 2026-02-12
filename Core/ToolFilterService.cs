using AsposeMcpServer.Core.Extension;
using AsposeMcpServer.Core.Session;

namespace AsposeMcpServer.Core;

/// <summary>
///     Service for filtering tools based on configuration.
/// </summary>
public class ToolFilterService
{
    /// <summary>
    ///     The extension configuration for extension tool filtering.
    /// </summary>
    private readonly ExtensionConfig _extensionConfig;

    /// <summary>
    ///     The server configuration containing tool category enablement settings.
    /// </summary>
    private readonly ServerConfig _serverConfig;

    /// <summary>
    ///     The session configuration for session tool filtering.
    /// </summary>
    private readonly SessionConfig _sessionConfig;

    /// <summary>
    ///     Initializes a new instance of the <see cref="ToolFilterService" /> class.
    /// </summary>
    /// <param name="serverConfig">The server configuration.</param>
    /// <param name="sessionConfig">The session configuration.</param>
    public ToolFilterService(ServerConfig serverConfig, SessionConfig sessionConfig)
        : this(serverConfig, sessionConfig, new ExtensionConfig())
    {
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="ToolFilterService" /> class.
    /// </summary>
    /// <param name="serverConfig">The server configuration.</param>
    /// <param name="sessionConfig">The session configuration.</param>
    /// <param name="extensionConfig">The extension configuration.</param>
    public ToolFilterService(ServerConfig serverConfig, SessionConfig sessionConfig, ExtensionConfig extensionConfig)
    {
        _serverConfig = serverConfig;
        _sessionConfig = sessionConfig;
        _extensionConfig = extensionConfig;
    }

    /// <summary>
    ///     Determines if a tool should be enabled based on its name.
    /// </summary>
    /// <param name="toolName">The tool name (e.g., "word_text", "excel_cell").</param>
    /// <returns><c>true</c> if the tool should be enabled; otherwise, <c>false</c>.</returns>
    public bool IsToolEnabled(string toolName)
    {
        if (string.IsNullOrEmpty(toolName))
            return false;

        if (toolName == "document_session")
            return _sessionConfig.Enabled;

        if (toolName == "extension")
            return _extensionConfig.Enabled;

        if (toolName.StartsWith("word_", StringComparison.OrdinalIgnoreCase))
            return _serverConfig.EnableWord;

        if (toolName.StartsWith("excel_", StringComparison.OrdinalIgnoreCase))
            return _serverConfig.EnableExcel;

        if (toolName.StartsWith("ppt_", StringComparison.OrdinalIgnoreCase))
            return _serverConfig.EnablePowerPoint;

        if (toolName.StartsWith("pdf_", StringComparison.OrdinalIgnoreCase))
            return _serverConfig.EnablePdf;

        if (toolName.StartsWith("ocr_", StringComparison.OrdinalIgnoreCase))
            return _serverConfig.EnableOcr;

        if (toolName.StartsWith("email_", StringComparison.OrdinalIgnoreCase))
            return _serverConfig.EnableEmail;

        if (toolName.StartsWith("barcode_", StringComparison.OrdinalIgnoreCase))
            return _serverConfig.EnableBarCode;

        if (toolName == "convert_document")
            return _serverConfig.EnableWord || _serverConfig.EnableExcel || _serverConfig.EnablePowerPoint ||
                   _serverConfig.EnablePdf;

        return true;
    }

    /// <summary>
    ///     Gets a list of enabled tool categories for logging.
    /// </summary>
    /// <returns>A comma-separated list of enabled categories.</returns>
    public string GetEnabledCategories()
    {
        var categories = new List<string>();

        if (_serverConfig.EnableWord) categories.Add("Word");
        if (_serverConfig.EnableExcel) categories.Add("Excel");
        if (_serverConfig.EnablePowerPoint) categories.Add("PowerPoint");
        if (_serverConfig.EnablePdf) categories.Add("PDF");
        if (_serverConfig.EnableOcr) categories.Add("OCR");
        if (_serverConfig.EnableEmail) categories.Add("Email");
        if (_serverConfig.EnableBarCode) categories.Add("BarCode");
        if (_sessionConfig.Enabled) categories.Add("Session");
        if (_extensionConfig.Enabled) categories.Add("Extension");

        return categories.Count > 0 ? string.Join(", ", categories) : "None";
    }
}
