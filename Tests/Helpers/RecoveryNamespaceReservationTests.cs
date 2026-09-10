using AsposeMcpServer.Core;
using AsposeMcpServer.Helpers;
using AsposeMcpServer.Tests.Infrastructure;
using AsposeMcpServer.Tools.BarCode;

namespace AsposeMcpServer.Tests.Helpers;

/// <summary>
///     R23-REC01: a caller's output path may not name the server's recovery state, whatever the
///     allowlist says. The rule sits in the central resolver, so every tool that resolves its
///     output — which the path meta-test requires of all of them — is covered by it.
/// </summary>
public class RecoveryNamespaceReservationTests : TestBase
{
    /// <summary>Every internal path a caller could aim at, relative to a recovery root.</summary>
    /// <param name="root">The recovery directory.</param>
    /// <returns>The key, ledger, a journal, the queue, a staged copy and its lease.</returns>
    private static IEnumerable<string> InternalPaths(string root)
    {
        return
        [
            Path.Combine(root, RecoveryCapability.KeyFileName),
            Path.Combine(root, PublishJournal.LedgerName),
            Path.Combine(root, "0123456789abcdef0123456789abcdef" + PublishJournal.Extension),
            Path.Combine(root, "aspose_cleanup_debts.json"),
            Path.Combine(root, ImmutableInputCopy.DirectoryName, "input.abc.mht"),
            Path.Combine(root, ImmutableInputCopy.DirectoryName, "input.abc.mht" + ImmutableInputCopy.LeaseSuffix),
            root
        ];
    }

    [Fact]
    public void EveryInternalPath_IsRefusedByTheResolver_UnderAnEmptyAllowlist()
    {
        var root = Path.Combine(TestDir, RecoveryContext.DirectoryName);
        foreach (var path in InternalPaths(root))
        {
            var refusal = Assert.Throws<ArgumentException>(() =>
                SecurityHelper.ResolveAndEnsureWithinAllowlist(path, [], "outputPath"));
            Assert.Contains("recovery state", refusal.Message, StringComparison.Ordinal);
        }
    }

    [Fact]
    public void EveryInternalPath_IsRefusedByTheResolver_WhenTheAllowlistHoldsTheParent()
    {
        // The allowlist names the temp root the recovery directory lives in: allowed by the
        // allowlist, refused by the reservation.
        var root = Path.Combine(TestDir, RecoveryContext.DirectoryName);
        foreach (var path in InternalPaths(root))
            Assert.Throws<ArgumentException>(() =>
                SecurityHelper.ResolveAndEnsureWithinAllowlist(path, [TestDir], "outputPath"));
    }

    [Fact]
    public void ASiblingOfTheRecoveryRoot_IsStillAllowed()
    {
        // The control: only the reserved name is reserved, not the temp root around it.
        var sibling = Path.Combine(TestDir, "ordinary-output", "file.png");
        Directory.CreateDirectory(Path.GetDirectoryName(sibling)!);
        Assert.EndsWith("file.png", SecurityHelper.ResolveAndEnsureWithinAllowlist(sibling, [TestDir], "outputPath"));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void AToolOutputAimedAtTheKeyOrTheLedger_IsRefusedBeforeTheFirstWrite_AndRecoverySurvivesARestart(
        bool allowlistTheParent)
    {
        // Through a real output-capable tool, under both configurations §40 names. The key and
        // the ledger must be byte-for-byte what they were, and the capability the same key.
        var recovery = RecoveryContext.For(TestDir);
        Assert.NotNull(recovery.Capability);
        var keyPath = Path.Combine(recovery.Directory, RecoveryCapability.KeyFileName);
        var keyBefore = File.ReadAllBytes(keyPath);
        PublishJournal.WriteLedger(recovery.Directory, recovery.Capability!, ["a-transaction"], 1,
            new HashSet<string>(StringComparer.Ordinal));
        var ledgerPath = Path.Combine(recovery.Directory, PublishJournal.LedgerName);
        var ledgerBefore = File.ReadAllBytes(ledgerPath);

        var config = allowlistTheParent
            ? ServerConfig.LoadFromArgs(["--allowed-path", TestDir])
            : ServerConfig.LoadFromArgs([]);
        var tool = new BarcodeGenerateTool(config);

        foreach (var target in new[] { keyPath, ledgerPath })
        {
            var refusal = Assert.ThrowsAny<Exception>(() => tool.Execute("generate", "payload", target));
            Assert.Contains("recovery state", refusal.Message, StringComparison.Ordinal);
        }

        Assert.Equal(keyBefore, File.ReadAllBytes(keyPath));
        Assert.Equal(ledgerBefore, File.ReadAllBytes(ledgerPath));
        Assert.Empty(Directory.GetFiles(recovery.Directory, "*.replaced-*"));

        // The restart: the same key is adopted and the ledger still verifies.
        var again = RecoveryContext.For(TestDir).Capability;
        Assert.NotNull(again);
        Assert.Equal(recovery.Capability!.Sign("a record"), again.Sign("a record"));
        var ledger = PublishJournal.ReadLedger(recovery.Directory, again);
        Assert.NotNull(ledger);
        Assert.Contains("a-transaction", ledger.Recovered);
    }
}
