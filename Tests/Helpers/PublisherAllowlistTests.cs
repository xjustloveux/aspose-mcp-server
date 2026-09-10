using AsposeMcpServer.Helpers;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Helpers;

/// <summary>
///     R8-C01: the allowlist the caller resolved its path against has to reach the publisher.
///     <para>
///         <see cref="BoundedFilePublisher.Publish" /> took the allowlist as an optional argument
///         with a null default, and all 21 conversion sinks used the four-argument form. The batch
///         underneath therefore held no allowlist, and the re-canonicalisation it performs
///         immediately before the move — the check that exists because the caller's own resolution
///         happened earlier — had nothing to check against. The parameter is now required, so a
///         caller that omits it does not compile.
///     </para>
/// </summary>
public class PublisherAllowlistTests : TestBase
{
    [Fact]
    public void ADestinationOutsideTheAllowlist_ShouldBeRefusedBeforeAnythingIsWritten()
    {
        var allowed = Path.Combine(TestDir, "allowed");
        var elsewhere = Path.Combine(TestDir, "elsewhere");
        Directory.CreateDirectory(allowed);
        Directory.CreateDirectory(elsewhere);

        var outside = Path.Combine(elsewhere, "published.txt");

        Assert.Throws<ArgumentException>(() =>
            BoundedFilePublisher.Publish(outside, 1_000,
                stream => stream.Write("payload"u8), "output", Recovery, [allowed]));

        Assert.False(File.Exists(outside));
        Assert.Empty(Directory.GetFiles(elsewhere));
    }

    [Fact]
    public void ADestinationInsideTheAllowlist_ShouldStillBePublished()
    {
        var allowed = Path.Combine(TestDir, "allowed_ok");
        Directory.CreateDirectory(allowed);
        var destination = Path.Combine(allowed, "published.txt");

        var written = BoundedFilePublisher.Publish(destination, 1_000,
            stream => stream.Write("payload"u8), "output", Recovery, [allowed]);

        Assert.Equal(7, written.WrittenBytes);
        Assert.Equal("payload", File.ReadAllText(destination));
    }

    [Fact]
    public void WithNoAllowlist_TheSameDestinationIsWritten()
    {
        // The contrast that makes the argument matter, and the reason it is now required: the
        // four-argument call the 21 conversion sinks used left the batch with nothing to check
        // against, so this destination was published rather than refused.
        var allowed = Path.Combine(TestDir, "contrast_allowed");
        var elsewhere = Path.Combine(TestDir, "contrast_elsewhere");
        Directory.CreateDirectory(allowed);
        Directory.CreateDirectory(elsewhere);

        var outside = Path.Combine(elsewhere, "published.txt");

        BoundedFilePublisher.Publish(outside, 1_000,
            stream => stream.Write("payload"u8), "output", Recovery, []);

        Assert.True(File.Exists(outside),
            "With no allowlist there is nothing to refuse, which is exactly what the optional "
            + "parameter allowed every conversion sink to do.");
    }

    // A mid-chain link swap between the caller's resolution and the publisher's own could not be
    // built honestly here: the batch stages beside the destination, so any swap of their shared
    // parent takes the staged file with it and the move fails for an unrelated reason. An earlier
    // attempt passed on exactly that IOException while `swapped` was false, which is the kind of
    // evidence this round exists to stop accepting. What is proven above is the property the
    // required argument restores: with the allowlist the destination is refused, without it the
    // same destination is written.
}
