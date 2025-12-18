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
    ///     Loads configuration from command line arguments
    /// </summary>
    /// <param name="args">Command line arguments</param>
    /// <returns>ServerConfig instance</returns>
    public static ServerConfig LoadFromArgs(string[] args)
    {
        var config = new ServerConfig
        {
            LicensePath = Environment.GetEnvironmentVariable("ASPOSE_LICENSE_PATH")
        };

        // If no arguments provided, enable all tools by default
        if (args.Length == 0) return config;

        // If arguments provided, disable all by default and only enable specified ones
        config.EnableWord = false;
        config.EnableExcel = false;
        config.EnablePowerPoint = false;
        config.EnablePdf = false;

        foreach (var arg in args)
            switch (arg.ToLower())
            {
                case "--word":
                    config.EnableWord = true;
                    break;
                case "--excel":
                    config.EnableExcel = true;
                    break;
                case "--powerpoint":
                case "--ppt":
                    config.EnablePowerPoint = true;
                    break;
                case "--pdf":
                    config.EnablePdf = true;
                    break;
                case "--all":
                    config.EnableWord = true;
                    config.EnableExcel = true;
                    config.EnablePowerPoint = true;
                    config.EnablePdf = true;
                    break;
                default:
                    if (arg.StartsWith("--license:", StringComparison.OrdinalIgnoreCase))
                        config.LicensePath = arg.Substring("--license:".Length);
                    else if (arg.StartsWith("--license=", StringComparison.OrdinalIgnoreCase))
                        config.LicensePath = arg.Substring("--license=".Length);
                    break;
            }

        return config;
    }

    /// <summary>
    ///     Gets a comma-separated string of enabled tool categories
    /// </summary>
    /// <returns>String listing enabled tools, or "None" if none are enabled</returns>
    public string GetEnabledToolsInfo()
    {
        var enabled = new List<string>();
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