using System.Diagnostics;
using System.Reflection;
using System.Text;

namespace AsposeMcpServer.Core.Transport;

/// <summary>
///     Builds the command line and executable used to launch the per-connection stdio child in
///     WebSocket mode. The child must run with the same tool selection, licence, path allowlist,
///     session and extension settings as the parent; dropping any of them silently changes the
///     child's security posture, so every argument is copied through except the ones that select
///     the parent's own transport.
/// </summary>
public static class ChildProcessArguments
{
    /// <summary>
    ///     Copies the parent's arguments for the child, dropping transport selection and appending
    ///     <c>--stdio</c>.
    /// </summary>
    /// <param name="args">The parent process arguments, verbatim.</param>
    /// <returns>
    ///     Argument tokens for <see cref="System.Diagnostics.ProcessStartInfo.ArgumentList" />.
    ///     Tokens are passed individually, so a value containing spaces needs no quoting.
    /// </returns>
    /// <exception cref="ArgumentException">
    ///     Thrown when a value-carrying transport option has no value. Skipping the following token
    ///     regardless of what it was removed whatever came next — including an option the child
    ///     needs to run with the parent's security posture (R3-C01). Refusing to start says so,
    ///     rather than launching children that quietly differ from the parent.
    /// </exception>
    public static IReadOnlyList<string> BuildChildArguments(IReadOnlyList<string> args)
    {
        List<string> forwarded = [];

        for (var i = 0; i < args.Count; i++)
        {
            var arg = args[i];

            if (TransportOptionGrammar.Flags.Any(flag =>
                    arg.Equals(flag, StringComparison.OrdinalIgnoreCase)))
                continue;

            if (TransportOptionGrammar.IsSeparatedValueOption(arg))
            {
                var next = i + 1 < args.Count ? args[i + 1] : null;
                if (!TransportOptionGrammar.IsValueFor(arg, next))
                    throw new ArgumentException(
                        $"Transport option '{arg}' has no value" +
                        (next is null ? "." : $"; '{next}' is not one.") +
                        " The child process arguments cannot be derived from an ambiguous command line.",
                        nameof(args));

                i++;
                continue;
            }

            if (TransportOptionGrammar.IsInlineValueOption(arg))
                continue;

            forwarded.Add(arg);
        }

        forwarded.Add("--stdio");
        return forwarded;
    }

    /// <summary>
    ///     Resolves the executable and any leading arguments needed to start another copy of this
    ///     server.
    /// </summary>
    /// <returns>
    ///     The executable to run and the arguments that must precede the forwarded ones. For a
    ///     self-contained apphost the executable is the host itself and the prefix is empty. When
    ///     the process was started through the <c>dotnet</c> muxer the executable is that muxer and
    ///     the prefix is the entry assembly, without which the child would receive the server's
    ///     options as if they were the muxer's own and fail immediately.
    /// </returns>
    /// <exception cref="InvalidOperationException">
    ///     Thrown when the current process is not a launchable server host, which is the case under
    ///     IIS in-process hosting where the process is the IIS worker.
    /// </exception>
    public static (string Executable, IReadOnlyList<string> Prefix) ResolveHostCommand()
    {
        var processPath = Environment.ProcessPath;
        if (string.IsNullOrEmpty(processPath))
            throw new InvalidOperationException(
                "WebSocket mode cannot determine the current executable and cannot start child processes.");

        var processName = Path.GetFileNameWithoutExtension(processPath);

        if (processName.Equals("dotnet", StringComparison.OrdinalIgnoreCase))
        {
            var entryAssembly = Assembly.GetEntryAssembly()?.Location;
            if (string.IsNullOrEmpty(entryAssembly))
                throw new InvalidOperationException(
                    "WebSocket mode is running through the dotnet muxer but the entry assembly path is unavailable.");

            return (processPath, [entryAssembly]);
        }

        if (processName.Equals("w3wp", StringComparison.OrdinalIgnoreCase) ||
            processName.Equals("iisexpress", StringComparison.OrdinalIgnoreCase))
            throw new InvalidOperationException(
                "WebSocket mode is not supported under IIS in-process hosting, because each connection " +
                "starts a child server process and the current process is the IIS worker. Use HTTP transport, " +
                "or host the server out-of-process.");

        return (processPath, []);
    }

    /// <summary>
    ///     Builds the start info the per-connection stdio child is launched with.
    ///     <para>
    ///         The tokens go into <see cref="ProcessStartInfo.ArgumentList" /> and never into
    ///         <see cref="ProcessStartInfo.Arguments" />: the list hands each token to the child as
    ///         one argument, while the string is re-split by the platform, so an allowlist entry
    ///         such as <c>D:\srv\my docs</c> would arrive as two arguments and the child would run
    ///         with a different allowlist than the parent. Having a single place build this is what
    ///         lets a test check it (R7-T01).
    ///     </para>
    /// </summary>
    /// <param name="executablePath">The executable resolved by <see cref="ResolveHostCommand" />.</param>
    /// <param name="arguments">The tokens from <see cref="BuildChildArguments" />.</param>
    /// <param name="groupId">Session group to pass through the environment, if any.</param>
    /// <param name="userId">Session user to pass through the environment, if any.</param>
    /// <returns>The start info for the child process.</returns>
    public static ProcessStartInfo CreateChildStartInfo(
        string executablePath, IReadOnlyList<string> arguments, string? groupId, string? userId)
    {
        var startInfo = new ProcessStartInfo
        {
            FileName = executablePath,
            UseShellExecute = false,
            RedirectStandardInput = true,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            CreateNoWindow = true,
            StandardInputEncoding = Encoding.UTF8,
            StandardOutputEncoding = Encoding.UTF8,
            StandardErrorEncoding = Encoding.UTF8
        };

        foreach (var argument in arguments)
            startInfo.ArgumentList.Add(argument);

        if (!string.IsNullOrEmpty(groupId))
            startInfo.Environment["ASPOSE_SESSION_GROUP_ID"] = groupId;
        if (!string.IsNullOrEmpty(userId))
            startInfo.Environment["ASPOSE_SESSION_USER_ID"] = userId;

        return startInfo;
    }
}
