using System.Diagnostics;
using AsposeMcpServer.Helpers;
using AsposeMcpServer.Tests.Infrastructure;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;

namespace AsposeMcpServer.Tests.Helpers;

/// <summary>
///     R18-ARCH02: two real processes on one queue file.
///     <para>
///         The queue synchronised its read-modify-write with a <c>lock</c> on a process-local
///         object. Two servers pointed at one temp directory share one queue — which is the point of
///         the queue, since a debt has to outlive the process that recorded it — and both read the
///         same list, each appended its own debt, and whichever wrote second erased the other's.
///     </para>
///     <para>
///         Two <see cref="CleanupDebtQueue" /> instances in this process would not show that: they
///         share the static gate registry and so are already ordered. The child here is a program
///         compiled at test time against the production assembly and run by <c>dotnet</c>, so the
///         only thing standing between the two writers is the operating system.
///     </para>
/// </summary>
public class CleanupDebtQueueCrossProcessTests : TestBase
{
    /// <summary>How many debts each child records.</summary>
    private const int PerChild = 30;

    /// <summary>The child program: the real queue, driven from another process.</summary>
    private const string ChildSource = """
                                       using System;
                                       using System.IO;
                                       using System.Threading;
                                       using AsposeMcpServer.Helpers;

                                       var queueFile = args[0];
                                       var recoveryRoot = args[1];
                                       var root = args[2];
                                       var prefix = args[3];
                                       var count = int.Parse(args[4]);
                                       var startFile = args[5];

                                       var queue = new CleanupDebtQueue(queueFile, new[] { root },
                                           RecoveryContext.For(recoveryRoot));

                                       var problems = 0;
                                       queue.OnQueueProblem = _ => problems++;

                                       // Both children wait on the same file so their writes overlap rather than queue up behind
                                       // each other's start-up.
                                       while (!File.Exists(startFile)) Thread.Sleep(2);

                                       for (var i = 0; i < count; i++)
                                       {
                                           var path = Path.Combine(root, prefix + "_" + i + ".txt");
                                           File.WriteAllText(path, "content");
                                           queue.Record(path, "locked");
                                       }

                                       Console.WriteLine("problems=" + problems);
                                       return 0;
                                       """;

    [SkippableFact]
    public void TwoProcessesRecordingIntoOneQueue_ShouldNotLoseEachOthersDebts()
    {
        var child = CompileChild();
        Skip.If(child == null, "The child program could not be compiled for this run.");

        try
        {
            RaceTwoChildren(child);
        }
        finally
        {
            Discard(child);
        }
    }

    /// <summary>Runs the race and checks that nothing was lost.</summary>
    /// <param name="child">The compiled child assembly.</param>
    private void RaceTwoChildren(string child)
    {
        var recoveryRoot = Path.Combine(TestDir, "shared-temp");
        Directory.CreateDirectory(recoveryRoot);

        // Established here, so the two children race on the queue rather than on the key file.
        var recovery = RecoveryContext.For(recoveryRoot);
        Skip.If(recovery.Capability == null, "No signing key could be established here.");

        var queueFile = Path.Combine(recovery.Directory, "cleanup-debts.json");
        var start = Path.Combine(TestDir, "go");

        var first = Run(child, queueFile, recoveryRoot, TestDir, "first", start);
        var second = Run(child, queueFile, recoveryRoot, TestDir, "second", start);

        File.WriteAllText(start, "go");

        var firstSaid = first.StandardOutput.ReadToEnd() + first.StandardError.ReadToEnd();
        var secondSaid = second.StandardOutput.ReadToEnd() + second.StandardError.ReadToEnd();

        Assert.True(first.WaitForExit(120_000), "the first child did not finish");
        Assert.True(second.WaitForExit(120_000), "the second child did not finish");
        Assert.True(first.ExitCode == 0, $"the first child exited {first.ExitCode}: {firstSaid}");
        Assert.True(second.ExitCode == 0, $"the second child exited {second.ExitCode}: {secondSaid}");

        var recorded = new CleanupDebtQueue(queueFile, [TestDir], recovery).Pending()
            .Select(debt => Path.GetFileName(debt.Path))
            .ToHashSet(StringComparer.OrdinalIgnoreCase);

        var expected = Enumerable.Range(0, PerChild)
            .SelectMany(i => new[] { $"first_{i}.txt", $"second_{i}.txt" })
            .ToList();

        // Said out loud by the children, so a host that could not establish its recovery context
        // fails here by name rather than as thirty silently missing debts.
        Assert.Contains("problems=0", firstSaid, StringComparison.Ordinal);
        Assert.Contains("problems=0", secondSaid, StringComparison.Ordinal);

        var lost = expected.Where(name => !recorded.Contains(name)).ToList();

        Assert.True(lost.Count == 0,
            $"{lost.Count} of {expected.Count} debts were lost to the other process: "
            + string.Join(", ", lost.Take(10)));
    }

    /// <summary>Removes a file the test left in the binary directory.</summary>
    /// <param name="path">The file.</param>
    private static void Discard(string path)
    {
        try
        {
            if (File.Exists(path)) File.Delete(path);
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            // A child assembly left behind is inert: its name carries a fresh guid, so nothing
            // will load it again.
        }
    }

    [SkippableFact]
    public void AQueueWhoseGateIsHeld_ShouldRefuseRatherThanWriteWithoutIt()
    {
        // The other half of R18-ARCH02, and the one the child processes cannot show: what happens
        // to the caller when the gate does not come. Writing anyway is the behaviour being fixed,
        // so "refuses and says so" has to be pinned, not assumed.
        Skip.IfNot(OperatingSystem.IsWindows(),
            "A second open of a held file is refused per-process only on Windows.");

        var queueFile = CreateTestFilePath("held_queue.json");
        var target = CreateTestFilePath("held_target.txt");
        File.WriteAllText(target, "content");

        var problems = new List<string>();
        var queue = new CleanupDebtQueue(queueFile, [TestDir], Recovery)
        {
            LockTimeout = TimeSpan.FromMilliseconds(200),
            OnQueueProblem = problems.Add
        };

        using (new FileStream(queueFile + ".lock", FileMode.OpenOrCreate, FileAccess.ReadWrite,
                   FileShare.None))
        {
            queue.Record(target, "locked");

            Assert.Contains(problems, problem =>
                problem.Contains("held by another process", StringComparison.Ordinal));
            Assert.False(File.Exists(queueFile), "the queue was written without the gate");

            var swept = queue.Sweep();
            Assert.Empty(swept.Deleted);
            Assert.True(File.Exists(target), "a sweep without the gate deleted the target");
        }

        // And once the gate is free the same queue works normally.
        queue.Record(target, "locked");
        Assert.Contains(target, queue.Pending().Select(debt => debt.Path));
    }

    /// <summary>Starts one child.</summary>
    /// <param name="child">The compiled child assembly.</param>
    /// <param name="queueFile">The queue both children write.</param>
    /// <param name="recoveryRoot">The temp root whose recovery context they share.</param>
    /// <param name="root">The allowlisted root the recorded paths live under.</param>
    /// <param name="prefix">Which child this is, so its debts are distinguishable.</param>
    /// <param name="start">The file whose appearance releases both children at once.</param>
    /// <returns>The running process.</returns>
    private static Process Run(string child, string queueFile, string recoveryRoot, string root,
        string prefix, string start)
    {
        var info = new ProcessStartInfo("dotnet")
        {
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true
        };

        // The test host's own manifests. Without them the child resolves the framework's
        // System.Text.Json rather than the app-local one this server is built against, and the
        // queue's static initialiser fails before it can race anything.
        var directory = Path.GetDirectoryName(child)!;

        foreach (var argument in new[]
                 {
                     "exec",
                     "--depsfile", Path.Combine(directory, "AsposeMcpServer.Tests.deps.json"),
                     "--runtimeconfig",
                     Path.Combine(directory, "AsposeMcpServer.Tests.runtimeconfig.json"),
                     child, queueFile, recoveryRoot, root, prefix,
                     PerChild.ToString(), start
                 })
            info.ArgumentList.Add(argument);

        var process = Process.Start(info);
        Assert.NotNull(process);
        return process;
    }

    /// <summary>Compiles the child program against the production assembly.</summary>
    /// <returns>The path to the child assembly, or null when it could not be built.</returns>
    /// <remarks>
    ///     Emitted beside the test binaries, and run with the test host's own dependency manifest
    ///     and runtime config, so the child loads the same assemblies this run does. Given neither,
    ///     it resolves the framework's <c>System.Text.Json</c> instead of the app-local one the
    ///     server is built against, and the queue fails to initialise.
    /// </remarks>
    private static string? CompileChild()
    {
        var appDirectory = AppContext.BaseDirectory;
        var name = "queue-race-child-" + Guid.NewGuid().ToString("N")[..8];
        var assemblyPath = Path.Combine(appDirectory, name + ".dll");

        if (!File.Exists(Path.Combine(appDirectory, "AsposeMcpServer.Tests.deps.json"))) return null;

        var platform = AppContext.GetData("TRUSTED_PLATFORM_ASSEMBLIES") as string;
        if (string.IsNullOrEmpty(platform)) return null;

        var references = platform.Split(Path.PathSeparator)
            .Where(path => path.EndsWith(".dll", StringComparison.OrdinalIgnoreCase))
            .Append(typeof(CleanupDebtQueue).Assembly.Location)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .Where(File.Exists)
            .Select(path => MetadataReference.CreateFromFile(path))
            .ToList();

        var compilation = CSharpCompilation.Create(name,
            [
                CSharpSyntaxTree.ParseText(ChildSource,
                    CSharpParseOptions.Default.WithLanguageVersion(LanguageVersion.Latest))
            ],
            references,
            new CSharpCompilationOptions(OutputKind.ConsoleApplication,
                optimizationLevel: OptimizationLevel.Release));

        return compilation.Emit(assemblyPath).Success ? assemblyPath : null;
    }
}
