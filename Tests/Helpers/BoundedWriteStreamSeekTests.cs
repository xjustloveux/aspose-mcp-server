using AsposeMcpServer.Helpers;

namespace AsposeMcpServer.Tests.Helpers;

/// <summary>
///     A bound that only counts the bytes handed to <c>Write</c> is not a bound on the stream.
///     <para>
///         <c>Position</c>, <c>Seek</c> and <c>SetLength</c> were forwarded unchanged, so with a
///         cap of eight bytes a caller could set the length to nine, or seek to a hundred and
///         write one byte, and the underlying stream grew to that size while this one still
///         reported almost nothing written (R4-R09). Whether the pinned Aspose serialisers use
///         those members is a separate question; the contract is wrong either way, and anything
///         written through this type relies on it.
///     </para>
/// </summary>
public class BoundedWriteStreamSeekTests
{
    [Fact]
    public void SetLength_PastTheLimit_ShouldBeRefused()
    {
        using var inner = new MemoryStream();
        using var bounded = new BoundedWriteStream(inner, 8, "output");

        Assert.Throws<ArgumentException>(() => bounded.SetLength(9));
        Assert.Equal(0, inner.Length);
    }

    [Fact]
    public void SetLength_AtTheLimit_ShouldBeAccepted()
    {
        using var inner = new MemoryStream();
        using var bounded = new BoundedWriteStream(inner, 8, "output");

        bounded.SetLength(8);

        Assert.Equal(8, inner.Length);
        Assert.Equal(8, bounded.Written);
    }

    [Fact]
    public void Position_PastTheLimit_ShouldBeRefused()
    {
        using var inner = new MemoryStream();
        using var bounded = new BoundedWriteStream(inner, 8, "output");

        Assert.Throws<ArgumentException>(() => bounded.Position = 100);
        Assert.Equal(0, inner.Length);
    }

    /// <param name="origin">Where the offset is measured from.</param>
    /// <param name="offset">The offset to seek by.</param>
    [Theory]
    [InlineData(SeekOrigin.Begin, 9)]
    [InlineData(SeekOrigin.Current, 9)]
    [InlineData(SeekOrigin.End, 9)]
    public void Seek_PastTheLimit_ShouldBeRefused(SeekOrigin origin, long offset)
    {
        using var inner = new MemoryStream();
        using var bounded = new BoundedWriteStream(inner, 8, "output");

        Assert.Throws<ArgumentException>(() => bounded.Seek(offset, origin));
        Assert.Equal(0, inner.Length);
    }

    [Fact]
    public void Seek_WithinTheLimit_ShouldBeAllowed()
    {
        using var inner = new MemoryStream();
        using var bounded = new BoundedWriteStream(inner, 8, "output");

        bounded.Write(new byte[4]);
        Assert.Equal(0, bounded.Seek(0, SeekOrigin.Begin));

        // Overwriting from the start is within the limit and must not be refused.
        bounded.Write(new byte[4]);
        Assert.Equal(4, inner.Length);
    }

    /// <summary>
    ///     The case a running sum cannot see: the write itself is small, but it lands past the
    ///     limit because of where the stream was positioned.
    /// </summary>
    [Fact]
    public void Write_AfterSeekingToTheLimit_ShouldBeRefused()
    {
        using var inner = new MemoryStream();
        using var bounded = new BoundedWriteStream(inner, 8, "output");

        bounded.Seek(8, SeekOrigin.Begin);

        Assert.Throws<ArgumentException>(() => bounded.Write(new byte[1]));
        Assert.Equal(0, inner.Length);
    }

    [Fact]
    public void Write_AfterSeekingBack_ShouldCountFromWhereItLands()
    {
        using var inner = new MemoryStream();
        using var bounded = new BoundedWriteStream(inner, 8, "output");

        bounded.Write(new byte[8]);
        bounded.Seek(0, SeekOrigin.Begin);
        bounded.Write(new byte[8]);

        // Rewriting the same eight bytes is still eight bytes.
        Assert.Equal(8, inner.Length);
        Assert.Equal(8, bounded.Written);
    }
}
