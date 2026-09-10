using System.Security.Cryptography;
using System.Text;

namespace AsposeMcpServer.Helpers;

/// <summary>
///     Reading and writing the files recovery depends on, in the two ways that turned out to matter.
///     <para>
///         Both the publish journal and the cleanup queue staged through a predictable
///         <c>&lt;name&gt;.writing</c> and wrote it with a path-based call, which opens whatever is
///         at that name — so a link planted there was followed and the server wrote through it
///         (R18-SEC02). And both checked a size with <c>FileInfo.Length</c> and then read the
///         content through a second open, so the file the cap was measured on did not have to be
///         the file that was read (R18-SEC03).
///     </para>
///     <para>
///         One place for both, because they were two copies of the same two mistakes — which is the
///         shape this codebase keeps paying for.
///     </para>
/// </summary>
public static class SecureFile
{
    /// <summary>
    ///     Writes text to a path, through a staging name nobody can predict, and replaces
    ///     atomically.
    /// </summary>
    /// <param name="path">The final path.</param>
    /// <param name="content">What to write.</param>
    /// <param name="write">
    ///     How to put the bytes on the open staging stream. A seam, so a fixture can make a write
    ///     fail or observe where it went. It receives the <em>stream</em>, never a path: the
    ///     previous seam took a path, so the staging file was created with <c>CreateNew</c>, its
    ///     handle closed, and the name reopened by the callback — a second open that a link
    ///     planted in that instant would redirect (R20-REC05, the original R19-REC02).
    /// </param>
    /// <exception cref="IOException">Thrown when the file could not be written.</exception>
    /// <remarks>
    ///     The staging name carries a nonce, so it is not a name anything could have planted a link
    ///     at in advance, and it is created with <c>CreateNew</c>, so an existing one is refused
    ///     rather than opened. The replace is atomic, which is what makes a crash mid-write leave
    ///     either the old file or the new one and never half of either.
    /// </remarks>
    public static void ReplaceAtomically(string path, string content,
        Action<Stream, string>? write = null)
    {
        var directory = Path.GetDirectoryName(path);
        if (!string.IsNullOrEmpty(directory)) Directory.CreateDirectory(directory);

        var staging = path + ".writing-" + Convert.ToHexString(RandomNumberGenerator.GetBytes(8))
            .ToLowerInvariant();

        try
        {
            // One open, held from creation to the end of the write. Whatever writes — the seam or
            // the default below — writes through this handle and nothing else, so the file that
            // was created is the file that gets the bytes.
            using (var stream = new FileStream(staging, FileMode.CreateNew, FileAccess.Write,
                       FileShare.None))
            {
                if (write != null)
                {
                    write(stream, content);
                }
                else
                {
                    using var writer = new StreamWriter(stream, leaveOpen: true);
                    writer.Write(content);
                }
            }

            File.Move(staging, path, true);
        }
        catch
        {
            try
            {
                if (File.Exists(staging)) File.Delete(staging);
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                // The staging file is named with a nonce, so one left behind is inert.
            }

            throw;
        }
    }

    /// <summary>Reads a file whose size is bounded, on the handle the bound was measured on.</summary>
    /// <param name="path">The file to read.</param>
    /// <param name="maximumBytes">The most this may read.</param>
    /// <param name="content">The content, when it was within the bound.</param>
    /// <param name="length">
    ///     How many bytes were read. When the bound was exceeded this is one past it, which is
    ///     enough to report "at least this large" without reading the rest.
    /// </param>
    /// <returns><c>true</c> when the file was within the bound.</returns>
    /// <remarks>
    ///     One open. Checking <c>FileInfo.Length</c> and then reading through a second open bounds
    ///     a file that need not be the one read: between the two calls it can be replaced or grown
    ///     (R18-SEC03). Reads one byte past the bound so "exactly at the limit" and "over it" are
    ///     distinguishable without reading the whole of an oversized file.
    /// </remarks>
    public static bool TryReadBounded(string path, long maximumBytes, out string content,
        out long length)
    {
        content = string.Empty;
        length = 0;

        using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);

        var buffer = new byte[maximumBytes + 1];
        var read = 0;
        int step;
        while (read < buffer.Length && (step = stream.Read(buffer, read, buffer.Length - read)) > 0)
            read += step;

        length = read;
        if (read > maximumBytes) return false;

        content = Encoding.UTF8.GetString(buffer, 0, read);
        return true;
    }
}
