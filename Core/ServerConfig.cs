namespace AsposeMcpServer.Core;

/// <summary>
///     Server configuration for enabling/disabling tool categories and license management.
/// </summary>
public class ServerConfig
{
    /// <summary>
    ///     Gets a value indicating whether Word document tools are enabled.
    ///     Enabled by default. Use specific tool arguments (e.g. <c>--word</c>) to enable only selected tools.
    /// </summary>
    public bool EnableWord { get; private set; } = true;

    /// <summary>
    ///     Gets a value indicating whether Excel spreadsheet tools are enabled.
    ///     Enabled by default. Use specific tool arguments (e.g. <c>--excel</c>) to enable only selected tools.
    /// </summary>
    public bool EnableExcel { get; private set; } = true;

    /// <summary>
    ///     Gets a value indicating whether PowerPoint presentation tools are enabled.
    ///     Enabled by default. Use specific tool arguments (e.g. <c>--powerpoint</c>) to enable only selected tools.
    /// </summary>
    public bool EnablePowerPoint { get; private set; } = true;

    /// <summary>
    ///     Gets a value indicating whether PDF document tools are enabled.
    ///     Enabled by default. Use specific tool arguments (e.g. <c>--pdf</c>) to enable only selected tools.
    /// </summary>
    public bool EnablePdf { get; private set; } = true;

    /// <summary>
    ///     Gets a value indicating whether OCR text recognition tools are enabled.
    ///     Enabled by default. Use specific tool arguments (e.g. <c>--ocr</c>) to enable only selected tools.
    /// </summary>
    public bool EnableOcr { get; private set; } = true;

    /// <summary>
    ///     Gets a value indicating whether Email tools are enabled.
    ///     Enabled by default. Use specific tool arguments (e.g. <c>--email</c>) to enable only selected tools.
    /// </summary>
    public bool EnableEmail { get; private set; } = true;

    /// <summary>
    ///     Gets a value indicating whether BarCode tools are enabled.
    ///     Enabled by default. Use specific tool arguments (e.g. <c>--barcode</c>) to enable only selected tools.
    /// </summary>
    public bool EnableBarCode { get; private set; } = true;

    /// <summary>
    ///     Gets the license file path or filename. Can be absolute path, relative path, or just filename.
    ///     If not specified, will search for common license file names.
    /// </summary>
    public string? LicensePath { get; private set; }

    /// <summary>
    ///     Loads configuration from environment variables and command line arguments.
    ///     Command line arguments take precedence over environment variables.
    /// </summary>
    /// <param name="args">Command line arguments.</param>
    /// <returns>A new <see cref="ServerConfig" /> instance.</returns>
    public static ServerConfig LoadFromArgs(string[] args)
    {
        var config = new ServerConfig();
        config.LoadFromEnvironment();
        config.LoadFromCommandLine(args);
        return config;
    }

    /// <summary>
    ///     Loads configuration from environment variables.
    /// </summary>
    private void LoadFromEnvironment()
    {
        var licensePath = Environment.GetEnvironmentVariable("ASPOSE_LICENSE_PATH");
        if (!string.IsNullOrEmpty(licensePath))
            LicensePath = licensePath;

        var tools = Environment.GetEnvironmentVariable("ASPOSE_TOOLS");
        if (!string.IsNullOrEmpty(tools)) ParseTools(tools);
    }

    /// <summary>
    ///     Parses tool specification string and enables corresponding tools.
    /// </summary>
    /// <param name="tools">Comma-separated tool names or "all".</param>
    private void ParseTools(string tools)
    {
        EnableWord = false;
        EnableExcel = false;
        EnablePowerPoint = false;
        EnablePdf = false;
        EnableOcr = false;
        EnableEmail = false;
        EnableBarCode = false;

        var toolList = tools.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        foreach (var tool in toolList)
            switch (tool.ToLower())
            {
                case "all":
                    EnableWord = true;
                    EnableExcel = true;
                    EnablePowerPoint = true;
                    EnablePdf = true;
                    EnableOcr = true;
                    EnableEmail = true;
                    EnableBarCode = true;
                    break;
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
                case "ocr":
                    EnableOcr = true;
                    break;
                case "email":
                    EnableEmail = true;
                    break;
                case "barcode":
                    EnableBarCode = true;
                    break;
            }
    }

    /// <summary>
    ///     Loads configuration from command line arguments (overrides environment variables).
    /// </summary>
    /// <param name="args">Command line arguments.</param>
    private void LoadFromCommandLine(string[] args)
    {
        if (args.Length == 0) return;

        var hasToolArg = args.Any(a =>
            a.Equals("--word", StringComparison.OrdinalIgnoreCase) ||
            a.Equals("--excel", StringComparison.OrdinalIgnoreCase) ||
            a.Equals("--powerpoint", StringComparison.OrdinalIgnoreCase) ||
            a.Equals("--ppt", StringComparison.OrdinalIgnoreCase) ||
            a.Equals("--pdf", StringComparison.OrdinalIgnoreCase) ||
            a.Equals("--ocr", StringComparison.OrdinalIgnoreCase) ||
            a.Equals("--email", StringComparison.OrdinalIgnoreCase) ||
            a.Equals("--barcode", StringComparison.OrdinalIgnoreCase) ||
            a.Equals("--all", StringComparison.OrdinalIgnoreCase));

        if (hasToolArg)
        {
            EnableWord = false;
            EnableExcel = false;
            EnablePowerPoint = false;
            EnablePdf = false;
            EnableOcr = false;
            EnableEmail = false;
            EnableBarCode = false;
        }

        for (var i = 0; i < args.Length; i++)
        {
            var originalArg = args[i];
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
                case "--ocr":
                    EnableOcr = true;
                    break;
                case "--email":
                    EnableEmail = true;
                    break;
                case "--barcode":
                    EnableBarCode = true;
                    break;
                case "--all":
                    EnableWord = true;
                    EnableExcel = true;
                    EnablePowerPoint = true;
                    EnablePdf = true;
                    EnableOcr = true;
                    EnableEmail = true;
                    EnableBarCode = true;
                    break;
                default:
                    if (arg == "--license" && i + 1 < args.Length)
                    {
                        LicensePath = args[i + 1];
                        i++;
                    }
                    else if (originalArg.StartsWith("--license:", StringComparison.OrdinalIgnoreCase))
                    {
                        LicensePath = originalArg["--license:".Length..];
                    }
                    else if (originalArg.StartsWith("--license=", StringComparison.OrdinalIgnoreCase))
                    {
                        LicensePath = originalArg["--license=".Length..];
                    }

                    break;
            }
        }
    }

    /// <summary>
    ///     Gets a comma-separated string of enabled tool categories.
    /// </summary>
    /// <returns>A string listing enabled tools, or "None" if none are enabled.</returns>
    public string GetEnabledToolsInfo()
    {
        List<string> enabled = [];
        if (EnableWord) enabled.Add("Word");
        if (EnableExcel) enabled.Add("Excel");
        if (EnablePowerPoint) enabled.Add("PowerPoint");
        if (EnablePdf) enabled.Add("PDF");
        if (EnableOcr) enabled.Add("OCR");
        if (EnableEmail) enabled.Add("Email");
        if (EnableBarCode) enabled.Add("BarCode");

        return enabled.Count > 0 ? string.Join(", ", enabled) : "None";
    }

    /// <summary>
    ///     Validates the configuration and throws an exception if invalid.
    /// </summary>
    /// <exception cref="InvalidOperationException">Thrown when no tool category is enabled.</exception>
    public void Validate()
    {
        if (!EnableWord && !EnableExcel && !EnablePowerPoint && !EnablePdf && !EnableOcr && !EnableEmail &&
            !EnableBarCode)
            throw new InvalidOperationException(
                "At least one tool category must be enabled. Use --word, --excel, --powerpoint, --pdf, --ocr, --email, --barcode, or --all");
    }
}
