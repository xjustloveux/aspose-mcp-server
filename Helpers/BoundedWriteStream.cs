namespace AsposeMcpServer.Helpers;

/// <summary>
///     Passes writes through to another stream and refuses the one that would take the total past
///     a limit.
///     <para>
///         Checking a size after the bytes exist is not a limit: the export paths saved a whole
///         worksheet into memory, or straight onto the caller's destination, and only then asked
///         whether it was too large (R3-R06). The allocation, or the file, already existed by
///         then. This stops at the write that crosses the line, so nothing beyond the limit is
///         ever held or written.
///     </para>
/// </summary>
/// <param name="inner">The stream being written to.</param>
/// <param name="maxBytes">Largest total this stream will accept.</param>
/// <param name="what">What is being written, for the error message.</param>
/// <remarks>The inner stream remains owned by the caller and is not disposed by this wrapper.</remarks>
public sealed class BoundedWriteStream(Stream inner, long maxBytes, string what) : Stream
{
    /// <summary>Bytes written so far.</summary>
    public long Written { get; private set; }

    /// <summary>Whether a write was refused for taking the total past the limit.</summary>
    public bool Refused { get; private set; }

    /// <inheritdoc />
    public override bool CanRead => inner.CanRead;

    /// <inheritdoc />
    public override bool CanSeek => inner.CanSeek;

    /// <inheritdoc />
    public override bool CanWrite => inner.CanWrite;

    /// <inheritdoc />
    public override long Length => inner.Length;

    /// <inheritdoc />
    /// <exception cref="ArgumentException">
    ///     Thrown when the position would place later writes past the limit. Forwarding it
    ///     unchecked let a caller seek past the cap and write there, which grew the underlying
    ///     stream well beyond it while this one still reported almost nothing written (R4-R09).
    /// </exception>
    public override long Position
    {
        get => inner.Position;
        set
        {
            EnsureWithinLimit(value, "seek to");
            inner.Position = value;
        }
    }

    /// <inheritdoc />
    public override void Flush()
    {
        inner.Flush();
    }

    /// <inheritdoc />
    public override int Read(byte[] buffer, int offset, int count)
    {
        return inner.Read(buffer, offset, count);
    }

    /// <inheritdoc />
    /// <exception cref="ArgumentException">Thrown when the target position is past the limit.</exception>
    public override long Seek(long offset, SeekOrigin origin)
    {
        var target = origin switch
        {
            SeekOrigin.Begin => offset,
            SeekOrigin.Current => inner.Position + offset,
            _ => inner.Length + offset
        };

        EnsureWithinLimit(target, "seek to");
        return inner.Seek(offset, origin);
    }

    /// <inheritdoc />
    /// <exception cref="ArgumentException">Thrown when the new length is past the limit.</exception>
    public override void SetLength(long value)
    {
        // A length is bytes the stream holds whether or not they were written through Write, so
        // this counts: SetLength(cap + 1) grew the underlying stream past the cap while Written
        // still read zero (R4-R09).
        EnsureWithinLimit(value, "set the length to");
        inner.SetLength(value);
        Written = Math.Max(Written, value);
    }

    /// <inheritdoc />
    /// <exception cref="ArgumentException">
    ///     Thrown when this write would take the total past the limit. It is thrown before the
    ///     bytes are passed on, so the underlying stream never holds more than the limit.
    /// </exception>
    public override void Write(byte[] buffer, int offset, int count)
    {
        // Measured from where the write lands, not from a running sum: a caller that seeks and
        // then writes puts bytes at that position, and counting only the bytes handed to this
        // method missed everything the seek skipped over (R4-R09).
        var end = (CanSeek ? inner.Position : Written) + count;
        EnsureWithinLimit(end, "write");

        Written = Math.Max(Written, end);
        inner.Write(buffer, offset, count);
    }

    /// <summary>
    ///     Refuses a position or length that would take the stream past its limit.
    /// </summary>
    /// <param name="value">The position, length or write end being requested.</param>
    /// <param name="action">What the caller is doing, for the message.</param>
    /// <exception cref="ArgumentException">Thrown when the value is above the limit.</exception>
    private void EnsureWithinLimit(long value, string action)
    {
        if (value <= maxBytes) return;

        Refused = true;
        throw new ArgumentException(
            $"The request would {action} {value:N0} bytes of {what}, above the limit of "
            + $"{maxBytes:N0}. Export a subset instead of the whole document.");
    }
}
