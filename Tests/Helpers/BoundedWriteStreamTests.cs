using System.Text;
using AsposeMcpServer.Helpers;

namespace AsposeMcpServer.Tests.Helpers;

/// <summary>
///     The export paths saved a whole worksheet into memory, or straight onto the caller's
///     destination, and only then asked whether it was too large — by which point the allocation
///     or the file already existed (R3-R06).
/// </summary>
public class BoundedWriteStreamTests
{
    [Fact]
    public void Write_BelowTheLimit_ShouldPassThrough()
    {
        using var inner = new MemoryStream();
        using var bounded = new BoundedWriteStream(inner, 100, "output");

        bounded.Write("hello"u8);

        Assert.Equal(5, bounded.Written);
        Assert.Equal("hello", Encoding.UTF8.GetString(inner.ToArray()));
    }

    [Fact]
    public void Write_AtExactlyTheLimit_ShouldBeAccepted()
    {
        using var inner = new MemoryStream();
        using var bounded = new BoundedWriteStream(inner, 8, "output");

        bounded.Write(new byte[8]);

        Assert.Equal(8, bounded.Written);
    }

    [Fact]
    public void Write_OneBytePastTheLimit_ShouldBeRefused()
    {
        using var inner = new MemoryStream();
        using var bounded = new BoundedWriteStream(inner, 8, "output");

        var exception = Assert.Throws<ArgumentException>(() => bounded.Write(new byte[9]));

        Assert.Contains("above", exception.Message);
    }

    /// <summary>
    ///     The refusal has to happen before the bytes are passed on, or the thing being protected
    ///     already holds them.
    /// </summary>
    [Fact]
    public void Write_PastTheLimit_ShouldNotWriteAnythingThrough()
    {
        using var inner = new MemoryStream();
        using var bounded = new BoundedWriteStream(inner, 8, "output");

        bounded.Write(new byte[5]);
        Assert.Throws<ArgumentException>(() => bounded.Write(new byte[5]));

        Assert.Equal(5, inner.Length);
    }

    [Fact]
    public void Write_AcrossSeveralCalls_ShouldCountTheTotal()
    {
        using var inner = new MemoryStream();
        using var bounded = new BoundedWriteStream(inner, 10, "output");

        bounded.Write(new byte[4]);
        bounded.Write(new byte[4]);

        Assert.Throws<ArgumentException>(() => bounded.Write(new byte[4]));
        Assert.Equal(8, inner.Length);
    }

    [Fact]
    public void Dispose_ShouldLeaveTheInnerStreamUsable()
    {
        using var inner = new MemoryStream();

        using (var bounded = new BoundedWriteStream(inner, 100, "output"))
        {
            bounded.Write(new byte[3]);
        }

        // The caller owns the destination and publishes it afterwards.
        inner.Write(new byte[2]);
        Assert.Equal(5, inner.Length);
    }
}
