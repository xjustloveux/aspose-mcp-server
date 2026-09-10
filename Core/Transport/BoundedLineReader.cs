using System.Text;

namespace AsposeMcpServer.Core.Transport;

/// <summary>
///     Reads newline-delimited text with a bound on how long one line may be.
///     <para>
///         <see cref="TextReader.ReadLineAsync()" /> has no bound at all: a child process that
///         writes without ever writing a newline makes the reader accumulate everything it emits,
///         so a broken or hostile child sets the server's memory use (R3-R02). This reader stops
///         at the limit and says so, leaving the decision to the caller.
///     </para>
/// </summary>
/// <param name="reader">The stream to read.</param>
/// <param name="maxChars">Longest line accepted, in characters.</param>
internal sealed class BoundedLineReader(TextReader reader, int maxChars)
{
    private readonly char[] _buffer = new char[8192];
    private int _length;
    private int _start;

    /// <summary>Whether the last read stopped because the line was longer than the bound.</summary>
    public bool LineTooLong { get; private set; }

    /// <summary>
    ///     Reads the next line.
    /// </summary>
    /// <param name="cancellationToken">Token cancelled when the connection ends.</param>
    /// <returns>
    ///     The line without its terminator, or <c>null</c> at the end of the stream. When
    ///     <see cref="LineTooLong" /> is set afterwards the returned text is the truncated prefix
    ///     and the remainder of that line has been discarded.
    /// </returns>
    public async Task<string?> ReadLineAsync(CancellationToken cancellationToken)
    {
        var builder = new StringBuilder();
        LineTooLong = false;

        while (true)
        {
            if (_start >= _length)
            {
                _length = await reader.ReadAsync(_buffer.AsMemory(), cancellationToken);
                _start = 0;
                if (_length == 0) return builder.Length == 0 ? null : builder.ToString().TrimEnd('\r');
            }

            // Index arithmetic rather than spans: a span local cannot live across an await, and
            // the buffer above is refilled by one.
            var remaining = _length - _start;
            var newline = Array.IndexOf(_buffer, '\n', _start, remaining);
            var take = newline >= 0 ? newline - _start : remaining;

            if (builder.Length + take > maxChars)
            {
                // Enough is kept to report the failure; the rest of the line is discarded, because
                // the caller is expected to stop reading this stream.
                builder.Append(_buffer, _start, Math.Max(0, maxChars - builder.Length));
                _start += take;
                LineTooLong = true;
                return builder.ToString();
            }

            builder.Append(_buffer, _start, take);
            _start += newline >= 0 ? take + 1 : take;

            if (newline >= 0) return builder.ToString().TrimEnd('\r');
        }
    }
}
