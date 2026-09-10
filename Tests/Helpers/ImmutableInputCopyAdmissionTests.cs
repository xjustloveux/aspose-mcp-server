using AsposeMcpServer.Helpers;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Helpers;

/// <summary>
///     R23-RES01: an input is admitted on size before it is staged, and a source that grows under
///     the copy is stopped at the cap with nothing left in the staging area.
/// </summary>
[Collection("SerialStaticSeams")]
public class ImmutableInputCopyAdmissionTests : TestBase
{
    private string Staging => Path.Combine(Recovery.Directory, ImmutableInputCopy.DirectoryName);

    /// <summary>What the staging area holds, leases included.</summary>
    private string[] Staged()
    {
        return Directory.Exists(Staging) ? Directory.GetFiles(Staging) : [];
    }

    [Fact]
    public void AnInputOneByteOverTheCap_IsRefusedBeforeAnythingIsStaged()
    {
        var input = CreateTestFilePath("over.mht");
        File.WriteAllBytes(input, new byte[4097]);
        Assert.NotNull(Recovery.Capability);

        var refusal = Assert.Throws<ArgumentException>(() =>
            ImmutableInputCopy.Of(input, Recovery, [TestDir], 4096));

        Assert.Contains("4,097", refusal.Message, StringComparison.Ordinal);
        Assert.Empty(Staged());
    }

    [Fact]
    public void AnInputExactlyAtTheCap_IsStaged()
    {
        var input = CreateTestFilePath("at-cap.mht");
        File.WriteAllBytes(input, new byte[4096]);

        using var copy = ImmutableInputCopy.Of(input, Recovery, [TestDir], 4096);
        Assert.Equal(4096, new FileInfo(copy.Path).Length);
    }

    [Fact]
    public void ASourceThatGrowsUnderTheCopy_IsStoppedAtTheCap_AndLeavesNoPartialCopy()
    {
        // The handle reports the admitted length; the bytes keep coming. Refused at the cap, the
        // partial copy removed under the lease, and the lease released.
        var input = CreateTestFilePath("growing.mht");
        File.WriteAllBytes(input, new byte[1024]);

        ImmutableInputCopy.OpenSource = _ => new GrowingStream(1024, 10_000);
        try
        {
            var refusal = Assert.Throws<ArgumentException>(() =>
                ImmutableInputCopy.Of(input, Recovery, [TestDir], 4096));
            Assert.Contains("grew", refusal.Message, StringComparison.Ordinal);
        }
        finally
        {
            ImmutableInputCopy.OpenSource = null;
        }

        Assert.Empty(Staged());
    }

    /// <summary>A stream whose length promises less than it delivers.</summary>
    private sealed class GrowingStream(long reportedLength, long actualLength) : Stream
    {
        public override bool CanRead => true;
        public override bool CanSeek => true;
        public override bool CanWrite => false;
        public override long Length => reportedLength;

        public override long Position { get; set; }

        public override int Read(byte[] buffer, int offset, int count)
        {
            var remaining = actualLength - Position;
            if (remaining <= 0) return 0;
            var n = (int)Math.Min(count, remaining);
            Array.Fill(buffer, (byte)'x', offset, n);
            Position += n;
            return n;
        }

        public override void Flush()
        {
        }

        public override long Seek(long offset, SeekOrigin origin)
        {
            throw new NotSupportedException();
        }

        public override void SetLength(long value)
        {
            throw new NotSupportedException();
        }

        public override void Write(byte[] buffer, int offset, int count)
        {
            throw new NotSupportedException();
        }
    }
}
