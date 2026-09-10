using AsposeMcpServer.Errors;
using AsposeMcpServer.Helpers;

namespace AsposeMcpServer.Tests.Helpers;

/// <summary>
///     R18-CONTRACT01: the one publish failure that does not mean "nothing changed".
///     <para>
///         Once every output is in place the batch has published. If the record that says so can
///         then be neither marked committed nor removed, the outputs stand and the journal is left
///         describing a publish that finished — which the next start would read as interrupted and
///         undo. The batch threw a plain <see cref="IOException" /> for this and left
///         <c>_published</c> false, so a caller could not tell it apart from the refusals that put
///         everything back, and the same batch would happily publish again.
///     </para>
/// </summary>
public class BoundedFileBatchCommitContractTests : IDisposable
{
    private readonly string _directory =
        Path.Combine(Path.GetTempPath(), "BatchCommit_" + Guid.NewGuid().ToString("N"));

    /// <summary>Creates the directory the batch writes into.</summary>
    public BoundedFileBatchCommitContractTests()
    {
        Directory.CreateDirectory(_directory);
    }

    /// <summary>This test class's recovery context.</summary>
    private RecoveryContext Recovery => RecoveryContext.For(_directory);

    /// <summary>Removes the directory.</summary>
    public void Dispose()
    {
        if (Directory.Exists(_directory)) Directory.Delete(_directory, true);
        GC.SuppressFinalize(this);
    }

    [SkippableFact]
    public void APublishWhoseRecordCannotBeSettled_IsIndeterminateAndTheBatchIsSpent()
    {
        // Only Windows refuses to move onto or delete a file somebody else holds open, which is
        // what puts the record beyond both ways of settling it.
        Skip.IfNot(OperatingSystem.IsWindows(),
            "A held-open file can still be replaced and unlinked on this platform.");

        var destination = Path.Combine(_directory, "output.txt");
        FileStream? held = null;

        using var batch = new BoundedFileBatch(4096, "a fixture", Recovery, [_directory]);
        batch.ReportDebt = _ => { };
        batch.BeforeMove = (_, _) => held ??= HoldTheJournalOpen();

        batch.Stage(destination, stream => stream.Write("the output"u8));

        // Taken while the publish is between its plan and its commit, so the journal exists and is
        // the one this publish will try to settle.
        try
        {
            var failure = Assert.Throws<PublishIndeterminateException>(() => batch.Publish());

            Assert.True(File.Exists(destination), "the output was not published");
            Assert.Equal("the output", File.ReadAllText(destination));
            Assert.Equal(held!.Name, failure.JournalPath);
            Assert.Contains("must not be retried", failure.Message, StringComparison.Ordinal);

            // Spent: the commit point has passed, so this batch cannot be used to publish again.
            Assert.Throws<ArgumentException>(() =>
                batch.Stage(Path.Combine(_directory, "again.txt"), _ => { }));
        }
        finally
        {
            held?.Dispose();
        }
    }

    /// <summary>Opens the journal this publish just wrote, and keeps it open.</summary>
    /// <returns>The handle holding it, or null when there is no journal to hold.</returns>
    /// <remarks>
    ///     Named by its extension rather than passed in: the path carries a fresh guid the batch
    ///     chooses, so a fixture can only find it by looking.
    /// </remarks>
    private FileStream? HoldTheJournalOpen()
    {
        var journal = Directory
            .GetFiles(Recovery.Directory, "*" + PublishJournal.Extension)
            .FirstOrDefault();

        return journal == null
            ? null
            : new FileStream(journal, FileMode.Open, FileAccess.Read, FileShare.None);
    }
}
