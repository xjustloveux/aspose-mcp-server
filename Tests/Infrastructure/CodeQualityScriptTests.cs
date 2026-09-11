using System.Diagnostics;

namespace AsposeMcpServer.Tests.Infrastructure;

/// <summary>Behavioral checks for the local code-quality runner.</summary>
public class CodeQualityScriptTests
{
    [Fact]
    public async Task InspectCode_ShouldReceiveConfiguredExclusions()
    {
        var root = RepositoryRoot();
        var script = Path.Combine(root.FullName, "code-quality.ps1").Replace("'", "''");
        var command = $$"""
                        function global:jb {
                            foreach ($argument in $args) {
                                [Console]::Out.WriteLine("JB_ARG=" + $argument)
                            }
                            $global:LASTEXITCODE = 0
                        }
                        . '{{script}}' -InspectCode
                        """;

        var startInfo = new ProcessStartInfo("pwsh")
        {
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            UseShellExecute = false
        };
        startInfo.ArgumentList.Add("-NoProfile");
        startInfo.ArgumentList.Add("-Command");
        startInfo.ArgumentList.Add(command);

        using var process = Process.Start(startInfo);
        Assert.NotNull(process);
        var outputTask = process.StandardOutput.ReadToEndAsync();
        var errorTask = process.StandardError.ReadToEndAsync();

        try
        {
            await process.WaitForExitAsync().WaitAsync(TimeSpan.FromSeconds(60));
        }
        catch (TimeoutException)
        {
            try
            {
                if (!process.HasExited) process.Kill(true);
            }
            catch (InvalidOperationException)
            {
                // The process exited between the state check and the termination request.
            }

            await process.WaitForExitAsync();
            throw;
        }

        var output = await outputTask;
        var error = await errorTask;

        Assert.True(process.ExitCode == 0, error);
        Assert.Contains("JB_ARG=--exclude=*.txt;report.xml;Tests/TestResults/**", output,
            StringComparison.Ordinal);
    }

    /// <summary>Finds the repository root containing the code-quality script.</summary>
    /// <returns>The repository root.</returns>
    private static DirectoryInfo RepositoryRoot()
    {
        for (var directory = new DirectoryInfo(AppContext.BaseDirectory);
             directory != null;
             directory = directory.Parent)
            if (File.Exists(Path.Combine(directory.FullName, "code-quality.ps1")))
                return directory;

        throw new DirectoryNotFoundException("Could not locate code-quality.ps1 from the test output directory.");
    }
}
