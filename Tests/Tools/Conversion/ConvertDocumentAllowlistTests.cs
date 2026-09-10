using System.Reflection;
using Aspose.Words;
using AsposeMcpServer.Core;
using AsposeMcpServer.Tests.Infrastructure;
using AsposeMcpServer.Tools.Conversion;

namespace AsposeMcpServer.Tests.Tools.Conversion;

/// <summary>
///     Covers A-02 for <c>convert_document</c>. The tool put <c>AllowedBasePaths</c> into the
///     conversion options, which govern where output may be written, but its own input path only
///     went through <c>ValidateFilePath</c> — a lexical check that says nothing about the
///     allowlist. Every reader downstream then opened the raw path, so a caller could convert a
///     file from outside the permitted roots and receive its content in the output.
/// </summary>
public class ConvertDocumentAllowlistTests : TestBase
{
    /// <summary>Builds a config whose allowlist holds the given roots; the setter is private.</summary>
    /// <param name="allowedPaths">Roots to permit.</param>
    /// <returns>The configured instance.</returns>
    private static ServerConfig BuildServerConfig(params string[] allowedPaths)
    {
        var config = new ServerConfig();
        typeof(ServerConfig)
            .GetProperty(nameof(ServerConfig.AllowedBasePaths), BindingFlags.Instance | BindingFlags.Public)!
            .SetValue(config, allowedPaths.Select(Path.GetFullPath).ToList());
        return config;
    }

    /// <summary>Writes a small Word document at the given path.</summary>
    /// <param name="path">Destination path; parent directories are created.</param>
    private static void WriteWordDocument(string path)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        var document = new Document();
        new DocumentBuilder(document).Writeln("secret contents");
        document.Save(path, SaveFormat.Docx);
    }

    [Fact]
    public void Convert_WithInputOutsideTheAllowlist_ShouldBeRefused()
    {
        var allowed = Path.Combine(TestDir, "allowed");
        var outside = Path.Combine(TestDir, "outside");
        Directory.CreateDirectory(allowed);
        var inputPath = Path.Combine(outside, "secret.docx");
        WriteWordDocument(inputPath);
        var outputPath = Path.Combine(allowed, "leaked.pdf");

        var tool = new ConvertDocumentTool(serverConfig: BuildServerConfig(allowed));

        Assert.ThrowsAny<ArgumentException>(() => tool.Execute(inputPath, outputPath: outputPath));
        Assert.False(File.Exists(outputPath));
    }

    [Fact]
    public void Convert_WithInputInsideTheAllowlist_ShouldStillWork()
    {
        var allowed = Path.Combine(TestDir, "allowed_ok");
        Directory.CreateDirectory(allowed);
        var inputPath = Path.Combine(allowed, "source.docx");
        WriteWordDocument(inputPath);
        var outputPath = Path.Combine(allowed, "converted.pdf");

        var tool = new ConvertDocumentTool(serverConfig: BuildServerConfig(allowed));

        tool.Execute(inputPath, outputPath: outputPath);

        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void Convert_WithNoAllowlistConfigured_ShouldStillWork()
    {
        var inputPath = Path.Combine(TestDir, "unrestricted", "source.docx");
        WriteWordDocument(inputPath);
        var outputPath = Path.Combine(TestDir, "unrestricted", "converted.pdf");

        var tool = new ConvertDocumentTool(serverConfig: new ServerConfig());

        tool.Execute(inputPath, outputPath: outputPath);

        Assert.True(File.Exists(outputPath));
    }
}
