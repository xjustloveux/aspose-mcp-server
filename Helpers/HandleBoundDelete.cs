using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;
using System.Runtime.Versioning;
using System.Security.Cryptography;
using Microsoft.Win32.SafeHandles;

namespace AsposeMcpServer.Helpers;

/// <summary>
///     Deletes a file through a handle opened on it, so that nothing which happens after the open
///     can redirect the deletion to a different file.
///     <para>
///         <c>File.Delete(path)</c> resolves the name a second time, so anything that replaced a
///         component of that path between the caller's decision and the call is what gets deleted.
///         Opening with <see cref="FileOptions.DeleteOnClose" /> removes that second resolution:
///         the name is resolved once, and the disposition is attached to the object it resolved
///         to. A rename afterwards moves the name, not the deletion.
///     </para>
///     <para>
///         What this does not do is close the whole check/use window, and an earlier version of
///         this text said it did (R13-F02). The cleanup queue validates a <em>path</em> — it
///         canonicalises it and walks the ancestors for reparse points — and this method then
///         resolves that same path itself, so a swap between the queue's check and this open is
///         still a swap this cannot see. Closing that half would mean checking the object in hand
///         rather than the name, which needs an identity for an open handle that .NET does not
///         expose portably. Narrowed, then, not closed: from "the whole time between the check and
///         the delete" to "nothing at all after the open" (§23.13.1).
///     </para>
///     <para>
///         The file goes when the last handle closes, which is also what lets this succeed on a
///         file another reader still holds — the case the debt exists for.
///     </para>
/// </summary>
public static class HandleBoundDelete
{
    private const uint FileReadData = 0x0001;
    private const uint FileReadAttributes = 0x0080;
    private const uint Synchronize = 0x00100000;
    private const uint DeleteAccess = 0x00010000;
    private const uint FileShareReadWriteDelete = 0x00000007;
    private const uint OpenExisting = 3;
    private const uint FileAttributeNormal = 0x00000080;
    private const uint FileFlagOpenReparsePoint = 0x00200000;

    /// <summary>The <c>FileDispositionInfo</c> class for <c>SetFileInformationByHandle</c>.</summary>
    private const int FileDispositionInfoClass = 4;

    /// <summary>
    ///     Deletes the file at <paramref name="path" />, through the handle opened here.
    /// </summary>
    /// <param name="path">The canonical path the caller has already validated.</param>
    /// <exception cref="IOException">
    ///     Thrown when the file cannot be opened for deletion. The caller treats this like any
    ///     other failed delete: the debt stays queued and is retried.
    /// </exception>
    /// <exception cref="UnauthorizedAccessException">
    ///     Thrown when the process may not delete the file.
    /// </exception>
    public static void Delete(string path)
    {
        Delete(path, null);
    }

    /// <summary>
    ///     Deletes the file at <paramref name="path" />, running <paramref name="whileOpen" /> in
    ///     the window between the open and the close.
    /// </summary>
    /// <param name="path">The canonical path the caller has already validated.</param>
    /// <param name="whileOpen">
    ///     Runs while the handle is held. A seam: the property this type exists for is about what
    ///     happens <em>during</em> that window, and a fixture with no way into it can only assert
    ///     that the named file went — which a path-based delete satisfies too (R13-F02).
    /// </param>
    /// <exception cref="IOException">
    ///     Thrown when the file cannot be opened for deletion.
    /// </exception>
    /// <exception cref="UnauthorizedAccessException">
    ///     Thrown when the process may not delete the file.
    /// </exception>
    internal static void Delete(string path, Action? whileOpen)
    {
        // FileShare.Delete lets this open succeed while another handle exists, and lets that
        // holder keep reading until it closes. Opening with DeleteOnClose is what asks for the
        // DELETE right and marks the disposition in one step.
        using var handle = new FileStream(path, FileMode.Open, FileAccess.Read,
            FileShare.ReadWrite | FileShare.Delete, 1, FileOptions.DeleteOnClose);

        whileOpen?.Invoke();
    }

    /// <summary>
    ///     Deletes the file at <paramref name="path" /> only if, judged through the very handle
    ///     that will delete it, it is still the file the caller means.
    /// </summary>
    /// <param name="path">The canonical path the caller has already validated.</param>
    /// <param name="stillTheFile">
    ///     Asked of the open stream. Everything the caller wants to know before deleting — length,
    ///     digest, anything readable — is answered from this handle and nothing else, so there is
    ///     no second resolution between the answer and the delete (R20-REC02).
    /// </param>
    /// <returns><c>true</c> when the file was deleted; <c>false</c> when the predicate refused.</returns>
    /// <exception cref="IOException">Thrown when the file cannot be opened or removed.</exception>
    /// <exception cref="UnauthorizedAccessException">Thrown when the process may not delete it.</exception>
    /// <remarks>
    ///     On Windows the handle is opened with the DELETE right and no delete-on-close, and the
    ///     disposition is set on that handle only after the predicate passes; a refusal closes a
    ///     handle that asked for nothing, so nothing happens to the file. On
    ///     Unix a delete-on-close open unlinks immediately, which cannot be taken back, so the
    ///     file is opened plainly, judged, and then unlinked <em>by path</em> while the handle is
    ///     still held. That last step is path-based: a rename of the directory entry between the
    ///     judgement and the unlink is not stopped by the open handle. There is no handle-bound
    ///     unlink in .NET; this is the weaker guarantee, stated rather than implied.
    /// </remarks>
    public static bool DeleteIf(string path, Func<FileStream, bool> stillTheFile)
    {
        if (OperatingSystem.IsWindows()) return DeleteIfOnWindows(path, stillTheFile);

        using var handle = new FileStream(path, FileMode.Open, FileAccess.Read,
            FileShare.ReadWrite | FileShare.Delete, 1, FileOptions.None);

        if (!stillTheFile(handle)) return false;

        File.Delete(path);
        return true;
    }

    /// <summary>The Windows shape of <see cref="DeleteIf" />: judged and disposed on one handle.</summary>
    /// <param name="path">The canonical path.</param>
    /// <param name="stillTheFile">The predicate, asked of the open stream.</param>
    /// <returns><c>true</c> when the file was deleted.</returns>
    /// <remarks>
    ///     Opened through <c>CreateFileW</c> with the DELETE right and <em>without</em>
    ///     <c>FILE_FLAG_DELETE_ON_CLOSE</c>. That flag is a property of the open and cannot be
    ///     withdrawn afterwards — measured: a refused predicate followed by
    ///     <c>FILE_DISPOSITION_INFO{FALSE}</c> still lost the file. So the disposition is set only
    ///     after the predicate passes, on this handle, and a refusal closes a handle that asked for
    ///     nothing. <c>FILE_FLAG_OPEN_REPARSE_POINT</c> is set so a link swapped in at the name is
    ///     opened as itself, fails the length or digest, and is refused rather than followed.
    /// </remarks>
    [SupportedOSPlatform("windows")]
    private static bool DeleteIfOnWindows(string path, Func<FileStream, bool> stillTheFile)
    {
        var raw = CreateFileW(path,
            FileReadData | FileReadAttributes | Synchronize | DeleteAccess,
            FileShareReadWriteDelete, IntPtr.Zero, OpenExisting,
            FileAttributeNormal | FileFlagOpenReparsePoint, IntPtr.Zero);

        var handle = new SafeFileHandle(raw, true);
        if (handle.IsInvalid)
        {
            var error = Marshal.GetLastWin32Error();
            handle.Dispose();
            throw error is 2 or 3
                ? new FileNotFoundException($"{path}: not found (Win32 error {error})", path)
                : new IOException($"{path}: could not be opened for a judged delete (Win32 error {error})");
        }

        using (handle)
        using (var stream = new FileStream(handle, FileAccess.Read, 1))
        {
            if (!stillTheFile(stream)) return false;

            var dispose = new FileDispositionInfo { DeleteFile = 1 };
            if (!SetFileInformationByHandle(handle, FileDispositionInfoClass, ref dispose,
                    Marshal.SizeOf<FileDispositionInfo>()))
                throw new IOException(
                    $"{path}: the file is the one recorded, but it could not be marked for deletion "
                    + $"(Win32 error {Marshal.GetLastWin32Error()})");

            return true;
        }
    }

    [SuppressMessage("Interoperability", "SYSLIB1054",
        Justification =
            "DllImport preferred over LibraryImport to avoid AllowUnsafeBlocks; these APIs are called infrequently")]
    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern IntPtr CreateFileW(string lpFileName, uint dwDesiredAccess, uint dwShareMode,
        IntPtr lpSecurityAttributes, uint dwCreationDisposition, uint dwFlagsAndAttributes,
        IntPtr hTemplateFile);

    /// <summary>The SHA-256 of a stream's content, read from its current position to the end.</summary>
    /// <param name="stream">The open file.</param>
    /// <returns>The digest as lower-case hex.</returns>
    /// <remarks>
    ///     Here so both the queue and the journal ask their "is this still the file" question of
    ///     the handle that is about to act, in one agreed way.
    /// </remarks>
    public static string DigestOf(Stream stream)
    {
        stream.Position = 0;
        return Convert.ToHexString(SHA256.HashData(stream)).ToLowerInvariant();
    }

    [SuppressMessage("Interoperability", "SYSLIB1054",
        Justification =
            "DllImport preferred over LibraryImport to avoid AllowUnsafeBlocks; these APIs are called infrequently")]
    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern bool SetFileInformationByHandle(SafeFileHandle hFile, int infoClass,
        ref FileDispositionInfo info, int size);

    /// <summary><c>FILE_DISPOSITION_INFO</c>: a single BOOLEAN.</summary>
    [StructLayout(LayoutKind.Sequential)]
    private struct FileDispositionInfo
    {
        /// <summary>Non-zero to delete on close; zero to withdraw a pending delete.</summary>
        public byte DeleteFile;
    }
}
