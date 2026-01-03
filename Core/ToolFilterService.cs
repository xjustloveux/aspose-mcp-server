using AsposeMcpServer.Core.Session;

namespace AsposeMcpServer.Core;

/// <summary>
///     Service for filtering tools based on configuration
/// </summary>
public class ToolFilterService
{
    private readonly ServerConfig _serverConfig;
    private readonly SessionConfig _sessionConfig;

    /// <summary>
    ///     Creates a new tool filter service
    /// </summary>
    /// <param name="serverConfig">Server configuration</param>
    /// <param name="sessionConfig">Session configuration</param>
    public ToolFilterService(ServerConfig serverConfig, SessionConfig sessionConfig)
    {
        _serverConfig = serverConfig;
        _sessionConfig = sessionConfig;
    }

    /// <summary>
    ///     Determines if a tool should be enabled based on its name
    /// </summary>
    /// <param name="toolName">The tool name (e.g., "word_text", "excel_cell")</param>
    /// <returns>True if the tool should be enabled</returns>
    public bool IsToolEnabled(string toolName)
    {
        if (string.IsNullOrEmpty(toolName))
            return false;

        // Session tool
        if (toolName == "document_session")
            return _sessionConfig.Enabled;

        // Word tools
        if (toolName.StartsWith("word_", StringComparison.OrdinalIgnoreCase))
            return _serverConfig.EnableWord;

        // Excel tools
        if (toolName.StartsWith("excel_", StringComparison.OrdinalIgnoreCase))
            return _serverConfig.EnableExcel;

        // PowerPoint tools
        if (toolName.StartsWith("ppt_", StringComparison.OrdinalIgnoreCase))
            return _serverConfig.EnablePowerPoint;

        // PDF tools
        if (toolName.StartsWith("pdf_", StringComparison.OrdinalIgnoreCase))
            return _serverConfig.EnablePdf;

        // Conversion tools
        if (toolName == "convert_to_pdf")
            return _serverConfig.EnableWord || _serverConfig.EnableExcel || _serverConfig.EnablePowerPoint;

        if (toolName == "convert_document")
        {
            // Need at least 2 document types enabled for cross-format conversion
            var enabledCount = 0;
            if (_serverConfig.EnableWord) enabledCount++;
            if (_serverConfig.EnableExcel) enabledCount++;
            if (_serverConfig.EnablePowerPoint) enabledCount++;
            return enabledCount >= 2;
        }

        // Unknown tools are enabled by default
        return true;
    }

    /// <summary>
    ///     Gets a list of enabled tool categories for logging
    /// </summary>
    /// <returns>Comma-separated list of enabled categories</returns>
    public string GetEnabledCategories()
    {
        var categories = new List<string>();

        if (_serverConfig.EnableWord) categories.Add("Word");
        if (_serverConfig.EnableExcel) categories.Add("Excel");
        if (_serverConfig.EnablePowerPoint) categories.Add("PowerPoint");
        if (_serverConfig.EnablePdf) categories.Add("PDF");
        if (_sessionConfig.Enabled) categories.Add("Session");

        return categories.Count > 0 ? string.Join(", ", categories) : "None";
    }
}