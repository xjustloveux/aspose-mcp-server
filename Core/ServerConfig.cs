using System.Text.Json;

namespace AsposeMcpServer.Core;

public class ServerConfig
{
    public bool EnableWord { get; set; } = true;
    public bool EnableExcel { get; set; } = true;
    public bool EnablePowerPoint { get; set; } = true;
    public bool EnablePdf { get; set; } = true;
    
    /// <summary>
    /// License file path or filename. Can be absolute path, relative path, or just filename.
    /// If not specified, will search for common license file names.
    /// </summary>
    public string? LicensePath { get; set; }

    public static ServerConfig LoadFromArgs(string[] args)
    {
        var config = new ServerConfig();
        
        // Check for license path from environment variable
        config.LicensePath = Environment.GetEnvironmentVariable("ASPOSE_LICENSE_PATH");
        
        // 如果没有指定任何参数，默认启用所有工具
        if (args.Length == 0)
        {
            return config;
        }

        // 如果指定了参数，则默认禁用所有，只启用指定的
        config.EnableWord = false;
        config.EnableExcel = false;
        config.EnablePowerPoint = false;
        config.EnablePdf = false;

        foreach (var arg in args)
        {
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
                    // Check if it's a license path argument (--license:path or --license=path)
                    if (arg.StartsWith("--license:", StringComparison.OrdinalIgnoreCase))
                    {
                        config.LicensePath = arg.Substring("--license:".Length);
                    }
                    else if (arg.StartsWith("--license=", StringComparison.OrdinalIgnoreCase))
                    {
                        config.LicensePath = arg.Substring("--license=".Length);
                    }
                    break;
            }
        }

        return config;
    }

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
    /// Validates the configuration and throws an exception if invalid
    /// </summary>
    /// <exception cref="InvalidOperationException">Thrown when configuration is invalid</exception>
    public void Validate()
    {
        if (!EnableWord && !EnableExcel && !EnablePowerPoint && !EnablePdf)
        {
            throw new InvalidOperationException(
                "At least one tool category must be enabled. Use --word, --excel, --powerpoint, --pdf, or --all");
        }
    }
}

