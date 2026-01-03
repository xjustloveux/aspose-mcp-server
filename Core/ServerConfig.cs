namespace AsposeMcpServer.Core;

/// <summary>
///     Server configuration for enabling/disabling tool categories and license management
/// </summary>
public class ServerConfig
{
    /// <summary>
    ///     Enable Word document tools
    /// </summary>
    public bool EnableWord { get; private set; } = true;

    /// <summary>
    ///     Enable Excel spreadsheet tools
    /// </summary>
    public bool EnableExcel { get; private set; } = true;

    /// <summary>
    ///     Enable PowerPoint presentation tools
    /// </summary>
    public bool EnablePowerPoint { get; private set; } = true;

    /// <summary>
    ///     Enable PDF document tools
    /// </summary>
    public bool EnablePdf { get; private set; } = true;

    /// <summary>
    ///     License file path or filename. Can be absolute path, relative path, or just filename.
    ///     If not specified, will search for common license file names.
    /// </summary>
    public string? LicensePath { get; private set; }

    /// <summary>
    ///     Loads configuration from environment variables and command line arguments.
    ///     Command line arguments take precedence over environment variables.
    /// </summary>
    /// <param name="args">Command line arguments</param>
    /// <returns>ServerConfig instance</returns>
    public static ServerConfig LoadFromArgs(string[] args)
    {
        var config = new ServerConfig();

        // Load from environment variables first (as defaults)
        config.LoadFromEnvironment();

        // Command line arguments override environment variables
        config.LoadFromCommandLine(args);

        return config;
    }

    /// <summary>
    ///     Loads configuration from environment variables
    /// </summary>
    private void LoadFromEnvironment()
    {
        // License path
        var licensePath = Environment.GetEnvironmentVariable("ASPOSE_LICENSE_PATH");
        if (!string.IsNullOrEmpty(licensePath))
            LicensePath = licensePath;

        // Tools (format: "all" or "word,excel,pdf,ppt")
        var tools = Environment.GetEnvironmentVariable("ASPOSE_TOOLS");
        if (!string.IsNullOrEmpty(tools)) ParseTools(tools);
    }

    /// <summary>
    ///     Parses tool specification string and enables corresponding tools
    /// </summary>
    /// <param name="tools">Comma-separated tool names or "all"</param>
    private void ParseTools(string tools)
    {
        // Reset all tools first
        EnableWord = false;
        EnableExcel = false;
        EnablePowerPoint = false;
        EnablePdf = false;

        var toolList = tools.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        foreach (var tool in toolList)
            switch (tool.ToLower())
            {
                case "all":
                    EnableWord = true;
                    EnableExcel = true;
                    EnablePowerPoint = true;
                    EnablePdf = true;
                    return; // "all" includes everything, no need to continue
                case "word":
                    EnableWord = true;
                    break;
                case "excel":
                    EnableExcel = true;
                    break;
                case "powerpoint":
                case "ppt":
                    EnablePowerPoint = true;
                    break;
                case "pdf":
                    EnablePdf = true;
                    break;
            }
    }

    /// <summary>
    ///     Loads configuration from command line arguments (overrides environment variables)
    /// </summary>
    /// <param name="args">Command line arguments</param>
    private void LoadFromCommandLine(string[] args)
    {
        if (args.Length == 0) return;

        // Check if any tool argument is specified
        var hasToolArg = args.Any(a =>
            a.Equals("--word", StringComparison.OrdinalIgnoreCase) ||
            a.Equals("--excel", StringComparison.OrdinalIgnoreCase) ||
            a.Equals("--powerpoint", StringComparison.OrdinalIgnoreCase) ||
            a.Equals("--ppt", StringComparison.OrdinalIgnoreCase) ||
            a.Equals("--pdf", StringComparison.OrdinalIgnoreCase) ||
            a.Equals("--all", StringComparison.OrdinalIgnoreCase));

        // Only reset tools if command line specifies tools (override env var)
        if (hasToolArg)
        {
            EnableWord = false;
            EnableExcel = false;
            EnablePowerPoint = false;
            EnablePdf = false;
        }

        foreach (var originalArg in args)
        {
            var arg = originalArg.ToLower();
            switch (arg)
            {
                case "--word":
                    EnableWord = true;
                    break;
                case "--excel":
                    EnableExcel = true;
                    break;
                case "--powerpoint":
                case "--ppt":
                    EnablePowerPoint = true;
                    break;
                case "--pdf":
                    EnablePdf = true;
                    break;
                case "--all":
                    EnableWord = true;
                    EnableExcel = true;
                    EnablePowerPoint = true;
                    EnablePdf = true;
                    break;
                default:
                    if (originalArg.StartsWith("--license:", StringComparison.OrdinalIgnoreCase))
                        LicensePath = originalArg["--license:".Length..];
                    else if (originalArg.StartsWith("--license=", StringComparison.OrdinalIgnoreCase))
                        LicensePath = originalArg["--license=".Length..];
                    break;
            }
        }
    }

    /// <summary>
    ///     Gets a comma-separated string of enabled tool categories
    /// </summary>
    /// <returns>String listing enabled tools, or "None" if none are enabled</returns>
    public string GetEnabledToolsInfo()
    {
        List<string> enabled = [];
        if (EnableWord) enabled.Add("Word");
        if (EnableExcel) enabled.Add("Excel");
        if (EnablePowerPoint) enabled.Add("PowerPoint");
        if (EnablePdf) enabled.Add("PDF");

        return enabled.Count > 0 ? string.Join(", ", enabled) : "None";
    }

    /// <summary>
    ///     Validates the configuration and throws an exception if invalid
    /// </summary>
    /// <exception cref="InvalidOperationException">Thrown when configuration is invalid</exception>
    public void Validate()
    {
        if (!EnableWord && !EnableExcel && !EnablePowerPoint && !EnablePdf)
            throw new InvalidOperationException(
                "At least one tool category must be enabled. Use --word, --excel, --powerpoint, --pdf, or --all");
    }
}