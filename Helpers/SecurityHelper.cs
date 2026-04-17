namespace AsposeMcpServer.Helpers;

/// <summary>
///     Security helper class for file path and name validation
/// </summary>
public static class SecurityHelper
{
    /// <summary>
    ///     Windows MAX_PATH limit for file paths
    /// </summary>
    private const int MaxPathLength = 260;

    /// <summary>
    ///     Maximum allowed file name length
    /// </summary>
    private const int MaxFileNameLength = 255;

    /// <summary>
    ///     Maximum allowed array size to prevent resource exhaustion
    /// </summary>
    private const int MaxArraySize = 1000;

    /// <summary>
    ///     Maximum allowed string length to prevent resource exhaustion
    /// </summary>
    private const int MaxStringLength = 10000;

    /// <summary>
    ///     Maximum directory recursion depth for <see cref="SafeRecursiveDelete" />.
    ///     Bounds pathological nesting that could otherwise cause stack overflow.
    /// </summary>
    private const int MaxRecursionDepth = 256;

    /// <summary>
    ///     Windows reserved device names that cause OS-level side effects when used as file names.
    /// </summary>
    private static readonly HashSet<string> WindowsReservedNames = new(StringComparer.OrdinalIgnoreCase)
    {
        "CON", "PRN", "AUX", "NUL",
        "COM1", "COM2", "COM3", "COM4", "COM5", "COM6", "COM7", "COM8", "COM9",
        "LPT1", "LPT2", "LPT3", "LPT4", "LPT5", "LPT6", "LPT7", "LPT8", "LPT9"
    };

    /// <summary>
    ///     Sanitizes a file name to prevent path traversal attacks
    /// </summary>
    /// <param name="fileName">Original file name</param>
    /// <returns>Sanitized file name safe for use in file operations</returns>
    public static string SanitizeFileName(string fileName)
    {
        if (string.IsNullOrWhiteSpace(fileName)) return "file";

        if (fileName.Length > MaxFileNameLength) fileName = fileName.Substring(0, MaxFileNameLength);

        var invalidChars = Path.GetInvalidFileNameChars();
        var sanitized = string.Join("_", fileName.Split(invalidChars, StringSplitOptions.RemoveEmptyEntries));

        sanitized = sanitized.Replace("..", "");
        sanitized = sanitized.Replace("\\", "_");
        sanitized = sanitized.Replace("/", "_");
        sanitized = sanitized.Replace(":", "_");
        sanitized = sanitized.Trim();
        sanitized = sanitized.Trim('.', ' ');

        if (string.IsNullOrWhiteSpace(sanitized)) sanitized = "file";

        if (sanitized.Length > MaxFileNameLength) sanitized = sanitized.Substring(0, MaxFileNameLength);

        return sanitized;
    }

    /// <summary>
    ///     Validates that a file path is safe and doesn't contain path traversal attempts,
    ///     control characters, Windows reserved device names, trailing dots/spaces, or
    ///     NTFS Alternate Data Stream markers.
    /// </summary>
    /// <param name="filePath">File path to validate.</param>
    /// <param name="allowAbsolutePaths">Whether to allow absolute paths (default: false for security).</param>
    /// <returns>True if path is safe, false otherwise.</returns>
    /// <remarks>
    ///     <para>
    ///         Control characters (<c>\x01</c>–<c>\x1F</c>) are explicitly rejected on all platforms.
    ///         <see cref="Path.GetInvalidPathChars" /> on Linux only returns <c>\0</c>, so without this
    ///         explicit check, control characters would pass through on Linux.
    ///     </para>
    ///     <para>
    ///         Windows reserved device names (CON, NUL, AUX, PRN, COM1–9, LPT1–9) are rejected
    ///         regardless of platform to prevent OS-level side effects when the server runs on Windows.
    ///     </para>
    ///     <para>
    ///         Trailing dots and spaces are rejected because Windows silently strips them,
    ///         which can bypass extension-based validation (e.g. <c>file.exe.</c> → <c>file.exe</c>).
    ///     </para>
    ///     <para>
    ///         NTFS Alternate Data Streams (colon in path segments, e.g. <c>file.txt:hidden</c>) are
    ///         rejected to prevent data hiding and path-validation bypass on NTFS volumes.
    ///     </para>
    /// </remarks>
    public static bool IsSafeFilePath(string filePath, bool allowAbsolutePaths = false)
    {
        if (string.IsNullOrWhiteSpace(filePath)) return false;

        if (filePath.Length > MaxPathLength) return false;

        // Reject control characters \x01–\x1F explicitly (Linux GetInvalidPathChars only has \0).
        if (filePath.Any(c => c is >= '\x01' and <= '\x1F'))
            return false;

        if (filePath.Contains("..", StringComparison.Ordinal) || filePath.Contains('~')) return false;

        if (filePath.Contains("//", StringComparison.Ordinal) ||
            filePath.Contains("\\\\", StringComparison.Ordinal)) return false;

        if (filePath.Length >= 3 && char.IsLetter(filePath[0]) && filePath[1] == ':' &&
            (filePath[2] == '\\' || filePath[2] == '/') && !allowAbsolutePaths)
            return false;

        if (Path.IsPathRooted(filePath))
        {
            if (!allowAbsolutePaths) return false;

            try
            {
                var fullPath = Path.GetFullPath(filePath);
                if (fullPath.Contains("..", StringComparison.Ordinal)) return false;
            }
            catch
            {
                return false;
            }
        }

        if (filePath.IndexOfAny(Path.GetInvalidPathChars()) >= 0) return false;

        // Reject NTFS Alternate Data Streams (colon in path segments beyond drive letter).
        // Allow colon at index 1 for Windows drive letters (e.g. C:\...).
        var adsCheckStart = filePath is [_, ':', ..] ? 2 : 0;
        if (filePath.IndexOf(':', adsCheckStart) >= 0) return false;

        // Reject trailing dots or spaces in any path segment (Windows silently strips them).
        foreach (var segment in filePath.Split(
                     [Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar],
                     StringSplitOptions.None))
        {
            if (segment.Length == 0) continue;
            var last = segment[^1];
            if (last is '.' or ' ') return false;
        }

        // Reject Windows reserved device names in any segment (case-insensitive).
        if (ContainsWindowsReservedName(filePath)) return false;

        return true;
    }

    /// <summary>
    ///     Checks whether any segment of <paramref name="filePath" /> is a Windows reserved device name,
    ///     with or without an extension (e.g. <c>CON</c>, <c>NUL.txt</c>).
    /// </summary>
    /// <param name="filePath">
    ///     The path to inspect; all segments split on both <see cref="Path.DirectorySeparatorChar" />
    ///     and <see cref="Path.AltDirectorySeparatorChar" /> are checked.
    /// </param>
    /// <returns>
    ///     <see langword="true" /> if any segment matches a reserved name (with or without extension);
    ///     <see langword="false" /> otherwise.
    /// </returns>
    private static bool ContainsWindowsReservedName(string filePath)
    {
        foreach (var segment in filePath.Split(
                     [Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar],
                     StringSplitOptions.None))
        {
            if (segment.Length == 0) continue;
            // Reserved names match with or without extension: "CON", "CON.txt", "NUL.anything".
            var nameWithoutExtension = segment;
            var dotIndex = segment.IndexOf('.');
            if (dotIndex >= 0) nameWithoutExtension = segment[..dotIndex];
            if (WindowsReservedNames.Contains(nameWithoutExtension)) return true;
        }

        return false;
    }

    /// <summary>
    ///     Validates and sanitizes a file path, throwing exception if invalid
    /// </summary>
    /// <param name="filePath">File path to validate</param>
    /// <param name="paramName">Parameter name for error message</param>
    /// <param name="allowAbsolutePaths">Whether to allow absolute paths</param>
    /// <returns>Validated file path</returns>
    /// <exception cref="ArgumentException">Thrown if path is invalid</exception>
    public static string ValidateFilePath(string filePath, string paramName = "path", bool allowAbsolutePaths = false)
    {
        if (string.IsNullOrWhiteSpace(filePath)) throw new ArgumentException($"{paramName} cannot be null or empty");

        if (!IsSafeFilePath(filePath, allowAbsolutePaths))
            throw new ArgumentException($"{paramName} contains invalid characters or path traversal attempts");

        return filePath;
    }

    /// <summary>
    ///     Validates that a file path is within the allowed base paths when a whitelist is configured.
    ///     If allowedBasePaths is empty, all paths are permitted (backward compatible).
    /// </summary>
    /// <param name="filePath">File path to validate</param>
    /// <param name="allowedBasePaths">List of allowed base directories (empty = allow all)</param>
    /// <param name="paramName">Parameter name for error message</param>
    /// <exception cref="ArgumentException">Thrown if path is outside all allowed directories</exception>
    public static void ValidatePathWithinAllowedBases(string filePath, IReadOnlyList<string> allowedBasePaths,
        string paramName = "path")
    {
        if (allowedBasePaths.Count == 0)
            return;

        var fullPath = Path.GetFullPath(filePath);
        foreach (var basePath in allowedBasePaths)
        {
            var normalizedBase = basePath.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
                                 + Path.DirectorySeparatorChar;
            if (fullPath.StartsWith(normalizedBase, StringComparison.OrdinalIgnoreCase))
                return;
        }

        throw new ArgumentException(
            $"{paramName} is outside the allowed directories. Configure --allowed-path to permit access.");
    }

    /// <summary>
    ///     Validates a user-supplied file path through the full three-step pipeline:
    ///     <see cref="ValidateFilePath" /> (traversal/character/length shape), followed by
    ///     <see cref="ResolveAndEnsureWithinAllowlist" /> (symbolic-link resolution + allowlist
    ///     re-check). This is the single source of truth for path validation at trust boundaries
    ///     (charter §5); callers in <c>DocumentSessionManager</c>, <c>TempFileManager</c> and
    ///     similar services should prefer this helper over the individual calls to avoid the
    ///     "validated shape but forgot symlink resolution" half-fix pattern.
    ///     Absolute paths are always permitted (matching the project-wide convention used by
    ///     <c>DocumentContext.cs:147-149</c>); the allowlist provides the bounding constraint
    ///     when configured.
    /// </summary>
    /// <param name="filePath">The user-supplied file path to validate.</param>
    /// <param name="allowedBasePaths">
    ///     The admin-configured allowlist of base directories. Pass an empty list (or
    ///     <c>ServerConfig.AllowedBasePaths</c> when no allowlist is configured) to skip the
    ///     allowlist step while still enforcing shape validation.
    /// </param>
    /// <param name="paramName">Parameter name used in thrown exception messages.</param>
    /// <exception cref="ArgumentException">
    ///     Thrown when the path is null/empty, contains traversal sequences or invalid characters,
    ///     contains a circular symbolic link, or resolves outside the configured allowlist.
    /// </exception>
    public static void ValidateUserPath(string filePath, IReadOnlyList<string> allowedBasePaths,
        string paramName = "path")
    {
        ValidateFilePath(filePath, paramName, true);
        ResolveAndEnsureWithinAllowlist(filePath, allowedBasePaths, paramName);
    }

    /// <summary>
    ///     Resolves <paramref name="path" /> through the filesystem (following symbolic links and
    ///     reparse points) and asserts that the resolved, canonical path still falls within at least
    ///     one of the <paramref name="allowedBases" /> directories. This closes the TOCTOU gap where
    ///     <see cref="ValidatePathWithinAllowedBases" /> passes the lexical check but the OS later
    ///     follows a planted symlink to a path outside the allowlist.
    /// </summary>
    /// <param name="path">
    ///     The path to resolve and check. Must be non-null/non-empty. May be absolute or relative;
    ///     relative paths are normalised via <see cref="Path.GetFullPath(string)" /> first.
    ///     If the path does not exist (common for write-sinks), every existing ancestor is resolved
    ///     in turn to catch mid-chain symlinks; the non-existent leaf is then re-appended.
    /// </param>
    /// <param name="allowedBases">
    ///     Admin-configured allowlist of base directories. Pass an empty list to skip the allowlist
    ///     step (same semantics as <see cref="ValidatePathWithinAllowedBases" />).
    /// </param>
    /// <param name="paramName">
    ///     Caller-supplied parameter name used in thrown exception messages. Must NOT be an internal
    ///     field name such as <c>"metadata.TempPath"</c> — use a name appropriate for user-facing
    ///     error messages.
    /// </param>
    /// <returns>
    ///     The resolved absolute path string (canonical form after following all symlinks). Callers
    ///     may optionally pass this to the sink instead of the original string to minimise the
    ///     residual TOCTOU window.
    /// </returns>
    /// <exception cref="ArgumentException">
    ///     Thrown when: (a) the path is null/empty; (b) a symlink in the path chain cannot be
    ///     resolved due to circular references or insufficient permissions; (c) the resolved target
    ///     lies outside all configured allowlist directories.
    /// </exception>
    /// <remarks>
    ///     <para>
    ///         <b>Residual TOCTOU window:</b> a sub-millisecond window remains between the
    ///         <see cref="FileSystemInfo.ResolveLinkTarget" /> call and the actual filesystem sink
    ///         (File.Copy, Aspose .Save, etc.). An attacker who already holds write access to the
    ///         directory can swap a symlink during this window. .NET 8 exposes no
    ///         <c>O_NOFOLLOW</c>-equivalent for the affected API surface. Calling this method
    ///         immediately before the sink (not at parameter-entry time) keeps the window as narrow
    ///         as possible. This residual is acceptable under the triage threat model (severity MEDIUM,
    ///         precondition: attacker holds write access to an allowed directory).
    ///     </para>
    ///     <para>
    ///         <b>Platform notes:</b> uses <see cref="FileSystemInfo.ResolveLinkTarget" /> (NET 8)
    ///         which handles POSIX symlinks and Windows reparse points (symlinks, junctions) uniformly.
    ///         Windows volume mount points may return <c>null</c> from <c>LinkTarget</c> and are
    ///         therefore not resolved — they pass through to the standard allowlist prefix check.
    ///         NTFS dedup/OneDrive placeholder reparse points are similarly not resolved (safe: no
    ///         bypass vector). POSIX bind mounts are not exposed via <c>LinkTarget</c> and are out
    ///         of scope.
    ///     </para>
    ///     <para>
    ///         <b>Dangling symlinks:</b> if the resolved target does not exist (<c>.Exists</c> is
    ///         false) the allowlist check still runs on the target path — so the path must reside
    ///         inside the allowlist even if it does not yet exist. A symlink in a mid-chain position
    ///         that points to a non-existent path outside the allowlist is rejected.
    ///     </para>
    ///     <para>
    ///         <b>Circular symlinks:</b> the OS/runtime caps the follow depth (~40 on Linux, ~63 on
    ///         Windows) and throws <see cref="IOException" />. The helper catches this and converts
    ///         it to <see cref="ArgumentException" /> with a sanitised message.
    ///     </para>
    /// </remarks>
    public static string ResolveAndEnsureWithinAllowlist(
        string path,
        IReadOnlyList<string> allowedBases,
        string paramName)
    {
        if (string.IsNullOrWhiteSpace(path))
            throw new ArgumentException($"{paramName} cannot be null or empty");

        // Step 1: lexical normalisation + initial allowlist check (fast path, no I/O).
        var fullPath = Path.GetFullPath(path);
        ValidatePathWithinAllowedBases(fullPath, allowedBases, paramName);

        // Step 2: resolve symlinks. For a path that does not exist yet (write sinks),
        // walk up ancestors until we find an existing node, resolve that, then re-append
        // the remaining segments so a symlinked intermediate directory is caught.
        try
        {
            var resolved = ResolveSymlinkChain(fullPath);

            // Step 3: re-check the resolved path against the allowlist.
            ValidatePathWithinAllowedBases(resolved, allowedBases, paramName);

            return resolved;
        }
        catch (IOException)
        {
            // Circular symlink — OS/runtime throws IOException("Too many levels of symbolic links").
            // Do not include path detail in the user-visible message (charter §5 / NV-1).
            throw new ArgumentException($"{paramName} cannot be resolved due to a circular symbolic link chain");
        }
        catch (UnauthorizedAccessException)
        {
            // ResolveLinkTarget may throw on paths the process cannot stat.
            throw new ArgumentException($"{paramName} cannot be resolved due to insufficient permissions");
        }
    }

    /// <summary>
    ///     Resolves all symlink components in <paramref name="fullPath" />, walking up to the
    ///     nearest existing ancestor when the leaf does not exist (write-sink pattern).
    ///     Re-appends the non-existent tail so the returned path preserves the intended filename.
    /// </summary>
    /// <param name="fullPath">
    ///     An absolute, lexically-normalised path (output of <see cref="Path.GetFullPath(string)" />).
    ///     Must not be null or empty.
    /// </param>
    /// <returns>
    ///     The resolved absolute path after following all symbolic links in the existing prefix.
    ///     Non-existent leaf segments are re-appended verbatim after the resolved ancestor.
    /// </returns>
    /// <exception cref="IOException">
    ///     Propagated from <see cref="FileSystemInfo.ResolveLinkTarget" /> when the OS detects a
    ///     circular symbolic link chain.
    /// </exception>
    /// <exception cref="UnauthorizedAccessException">
    ///     Propagated from <see cref="FileSystemInfo.ResolveLinkTarget" /> when the process cannot
    ///     stat an ancestor directory.
    /// </exception>
    private static string ResolveSymlinkChain(string fullPath)
    {
        // Try direct resolution first (common case: path exists and is not a symlink).
        // File.Exists is checked before Directory.Exists to avoid a redundant I/O call.
        FileSystemInfo? fsi;
        if (File.Exists(fullPath))
            fsi = new FileInfo(fullPath);
        else if (Directory.Exists(fullPath))
            fsi = new DirectoryInfo(fullPath);
        else
            fsi = null;

        if (fsi != null)
        {
            var target = fsi.ResolveLinkTarget(true);
            // null means the path is not a symlink/reparse point — resolved == fullPath.
            return target != null ? Path.GetFullPath(target.FullName) : fullPath;
        }

        // Path does not exist (write sink). Walk up to find the nearest existing ancestor
        // to catch a symlinked intermediate directory (NV-3).
        var segments = new Stack<string>();
        var current = fullPath;

        while (true)
        {
            var parent = Path.GetDirectoryName(current);
            if (parent == null || parent == current)
                // Reached the root; nothing more to resolve.
                return fullPath;

            segments.Push(Path.GetFileName(current));
            current = parent;

            if (Directory.Exists(current))
            {
                var di = new DirectoryInfo(current);
                var resolvedTarget = di.ResolveLinkTarget(true);
                var resolvedBase = resolvedTarget != null
                    ? Path.GetFullPath(resolvedTarget.FullName)
                    : current;

                // Re-append the non-existent tail segments after the resolved ancestor.
                while (segments.Count > 0)
                    resolvedBase = Path.Combine(resolvedBase, segments.Pop());

                return resolvedBase;
            }
        }
    }

    /// <summary>
    ///     Safely deletes a directory and all of its contents without following symbolic links
    ///     into directories outside <paramref name="allowedBases" />. This replaces
    ///     <c>Directory.Delete(path, recursive:true)</c> which follows child symlinks on .NET.
    /// </summary>
    /// <param name="path">
    ///     The directory to delete. Validated via <see cref="ResolveAndEnsureWithinAllowlist" />
    ///     before any deletion begins. Must be non-null/non-empty.
    /// </param>
    /// <param name="allowedBases">
    ///     Admin-configured allowlist of base directories used to validate both the root path and
    ///     each recursive descent. Pass a single-element list with the server-private temp directory
    ///     for transport cleanup calls.
    /// </param>
    /// <param name="paramName">
    ///     Caller-contextual name of the <paramref name="path" /> argument, used in error messages
    ///     to identify which call site triggered the violation. Defaults to <c>"path"</c> for
    ///     backward compatibility. Callers should pass <c>nameof(theirVariable)</c> for
    ///     log-driven debugging. Must not be an internal field name (RV-6 / NV-1 compliance).
    /// </param>
    /// <exception cref="ArgumentException">
    ///     Thrown when <paramref name="path" /> or any resolved ancestor lies outside
    ///     <paramref name="allowedBases" />, or when a circular symlink is encountered.
    /// </exception>
    /// <remarks>
    ///     <para>
    ///         <b>Symlink handling (NV-2):</b> at each recursion entry the directory's
    ///         <see cref="DirectoryInfo.LinkTarget" /> is re-checked (not only during the initial
    ///         enumeration) to close the mid-walk plant window where an attacker replaces a real
    ///         directory with a symlink between enumeration and descent.
    ///     </para>
    ///     <para>
    ///         <b>Windows junctions:</b> <c>DirectoryInfo.LinkTarget</c> returns non-null for NTFS
    ///         junctions on .NET 8. <c>Directory.Delete(junctionPath, recursive:false)</c> removes
    ///         only the junction point, not the target contents — this is the intended behaviour.
    ///     </para>
    ///     <para>
    ///         <b>Windows volume mount points:</b> <c>LinkTarget</c> returns null for volume mount
    ///         points in some .NET versions; the walker treats them as real directories and recurses.
    ///         This is a known gap documented here; volume mount points inside server temp directories
    ///         are an unusual configuration.
    ///     </para>
    ///     <para>
    ///         <b>Depth limit:</b> recursion is bounded to <c>256</c> levels to defeat
    ///         pathological nesting that would cause a stack overflow.
    ///     </para>
    /// </remarks>
    public static void SafeRecursiveDelete(string path, IReadOnlyList<string> allowedBases,
        string paramName = "path")
    {
        // Validate and resolve the root before touching anything.
        ResolveAndEnsureWithinAllowlist(path, allowedBases, paramName);

        if (!Directory.Exists(path))
            return;

        SafeRecursiveDeleteCore(path, 0);
    }

    /// <summary>
    ///     Core recursive implementation for <see cref="SafeRecursiveDelete" />.
    ///     Re-checks the link status of <paramref name="dirPath" /> at every recursion entry
    ///     (NV-2: closes the mid-walk symlink plant window).
    /// </summary>
    /// <param name="dirPath">The current directory being processed.</param>
    /// <param name="depth">Current recursion depth; capped at <see cref="MaxRecursionDepth" />.</param>
    /// <remarks>
    ///     When <paramref name="depth" /> exceeds <see cref="MaxRecursionDepth" />, the method
    ///     stops recursing and returns rather than throwing. This avoids leaving a partially deleted
    ///     tree (the caller would catch an exception after some children were already removed).
    ///     The excessively deep subtree is left intact — an operational concern, not a security one,
    ///     since the allowlist check at the root already passed.
    /// </remarks>
    private static void SafeRecursiveDeleteCore(string dirPath, int depth)
    {
        if (depth > MaxRecursionDepth)
            // Stop recursing rather than throw — throwing would leave a partially deleted tree.
            // The remaining subtree is left intact; this is an operational concern only.
            return;

        // NV-2: re-check LinkTarget at recursion entry, not just at enumeration time.
        // An attacker could swap a real directory for a symlink after we enumerated its parent.
        var currentDi = new DirectoryInfo(dirPath);
        if (currentDi.LinkTarget != null)
        {
            // This entry is a symlink — remove only the link node, not the target.
            Directory.Delete(dirPath, false);
            return;
        }

        try
        {
            foreach (var entry in currentDi.EnumerateFileSystemInfos())
                if (entry is FileInfo fi)
                {
                    // Symlinked file: File.Delete unlinks the symlink itself on POSIX (correct).
                    // On Windows it also removes only the reparse point, not the target.
                    fi.Delete();
                }
                else if (entry is DirectoryInfo di)
                {
                    if (di.LinkTarget != null)
                        // Remove symlinked directory entry without recursing into the target.
                        Directory.Delete(di.FullName, false);
                    else
                        SafeRecursiveDeleteCore(di.FullName, depth + 1);
                }

            // All children removed; delete the now-empty directory itself.
            Directory.Delete(dirPath, false);
        }
        catch (DirectoryNotFoundException)
        {
            // Another process may have removed the directory concurrently — treat as success.
        }
    }

    /// <summary>
    ///     Returns <c>true</c> when <paramref name="candidate" />, after lexical normalisation,
    ///     is prefixed by <paramref name="baseDirectory" />.  Internal helper used by
    ///     <see cref="ValidatePathWithinAllowedBases" /> and callable by tests via
    ///     <c>InternalsVisibleTo</c>.
    /// </summary>
    /// <param name="candidate">
    ///     Absolute path to test. Must not be null or empty; <see cref="Path.GetFullPath(string)" />
    ///     normalisation is applied.
    /// </param>
    /// <param name="baseDirectory">
    ///     The directory the candidate must reside under. Trailing separators are normalised before
    ///     comparison. Must not be null or empty.
    /// </param>
    /// <returns>
    ///     <c>true</c> if the normalised candidate starts with the normalised base directory prefix;
    ///     <c>false</c> if either argument is null/empty or if path normalisation throws.
    /// </returns>
    internal static bool IsPathUnder(string candidate, string baseDirectory)
    {
        if (string.IsNullOrEmpty(candidate) || string.IsNullOrEmpty(baseDirectory))
            return false;
        try
        {
            var fullCandidate = Path.GetFullPath(candidate);
            var fullBase = Path.GetFullPath(baseDirectory)
                               .TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
                           + Path.DirectorySeparatorChar;
            return fullCandidate.StartsWith(fullBase, StringComparison.OrdinalIgnoreCase);
        }
        catch
        {
            return false;
        }
    }

    /// <summary>
    ///     Validates and sanitizes a file name pattern (for use in split/export tools)
    /// </summary>
    /// <param name="pattern">File name pattern (may contain placeholders like {index}, {name})</param>
    /// <returns>Sanitized pattern safe for use</returns>
    public static string SanitizeFileNamePattern(string pattern)
    {
        if (string.IsNullOrWhiteSpace(pattern)) return "file_{index}";

        if (pattern.Length > MaxFileNameLength) pattern = pattern.Substring(0, MaxFileNameLength);

        var sanitized = pattern.Replace("\\", "_").Replace("/", "_");
        sanitized = sanitized.Replace("..", "");
        sanitized = sanitized.Replace(":", "_");
        sanitized = sanitized.Trim('.', ' ');

        if (string.IsNullOrWhiteSpace(sanitized)) sanitized = "file_{index}";

        return sanitized;
    }

    /// <summary>
    ///     Validates array size to prevent resource exhaustion
    /// </summary>
    /// <param name="array">Array to validate</param>
    /// <param name="paramName">Parameter name for error message</param>
    /// <param name="maxSize">Maximum allowed size (default: MaxArraySize)</param>
    /// <exception cref="ArgumentException">Thrown if array is too large</exception>
    public static void ValidateArraySize<T>(IEnumerable<T> array, string paramName = "array", int? maxSize = null)
    {
        var limit = maxSize ?? MaxArraySize;

        var count = array is ICollection<T> collection
            ? collection.Count
            : array.Count();

        if (count > limit) throw new ArgumentException($"{paramName} exceeds maximum allowed size of {limit}");
    }

    /// <summary>
    ///     Validates string length to prevent resource exhaustion
    /// </summary>
    /// <param name="value">String to validate</param>
    /// <param name="paramName">Parameter name for error message</param>
    /// <param name="maxLength">Maximum allowed length (default: MaxStringLength)</param>
    /// <exception cref="ArgumentException">Thrown if string is too long</exception>
    public static void ValidateStringLength(string value, string paramName = "value", int? maxLength = null)
    {
        var limit = maxLength ?? MaxStringLength;

        if (value.Length > limit) throw new ArgumentException($"{paramName} exceeds maximum allowed length of {limit}");
    }
}
