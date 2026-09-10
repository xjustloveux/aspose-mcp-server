using System.Security.Cryptography;
using AsposeMcpServer.Helpers;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Helpers;

/// <summary>
///     §23.13.1: the cleanup queue's check/delete window, narrowed to nothing after the open.
///     <para>
///         <c>File.Delete(path)</c> resolves the name a second time, so whatever replaced a
///         component of that path between the check and the call is what gets deleted. Deleting
///         through the handle removes that second resolution: the disposition belongs to the
///         object the open resolved to, and a rename afterwards moves the name, not the deletion.
///     </para>
///     <para>
///         Not the whole window, and this file used to say it was (R13-F02). The queue's check is
///         still made on a path, and <c>HandleBoundDelete</c> resolves that path itself, so a swap
///         between the two is invisible to both — <c>ThePathIsStillResolvedOnce</c> below is the
///         boundary written down. The fixture that was supposed to prove the rest never replaced
///         anything: it wrote two files, deleted one, and asserted the other survived, which
///         <c>File.Delete</c> does too.
///     </para>
/// </summary>
public class HandleBoundDeleteTests : TestBase
{
    [Fact]
    public void AnOrdinaryFile_ShouldBeDeleted()
    {
        var path = CreateTestFilePath("handle_delete_ordinary.txt");
        File.WriteAllText(path, "content");

        HandleBoundDelete.Delete(path);

        Assert.False(File.Exists(path));
    }

    [SkippableFact]
    public void AFileReplacedAfterTheHandleIsOpen_ShouldNotRedirectTheDeletion()
    {
        Skip.IfNot(OperatingSystem.IsWindows(),
            "The handle-bound path is the Windows one; elsewhere unlink resolves the name.");

        // The swap actually happens here, in the window the handle covers: the checked file is
        // moved aside and a different file is put under its name. A path-based delete would take
        // the replacement, because it looks the name up again; this one takes the object it
        // opened, wherever that object has been moved to.
        var victim = CreateTestFilePath("handle_delete_victim.txt");
        var movedAside = CreateTestFilePath("handle_delete_moved.txt");
        File.WriteAllText(victim, "the file that was checked");

        HandleBoundDelete.Delete(victim, () =>
        {
            File.Move(victim, movedAside);
            File.WriteAllText(victim, "the replacement");
        });

        Assert.False(File.Exists(movedAside), "the checked file survived under its new name");
        Assert.True(File.Exists(victim), "the replacement was deleted instead of the checked file");
        Assert.Equal("the replacement", File.ReadAllText(victim));
    }

    [Fact]
    public void ThePathIsStillResolvedOnce_SoASwapBeforeTheOpenIsNotCovered()
    {
        // The boundary, written down rather than left to the class summary. What is bound to the
        // handle is everything after the open; the caller's own check happened earlier and on a
        // path, so a name that changed meaning in between is opened as it now reads. Closing that
        // half means identifying the object in hand, which .NET does not expose portably — so the
        // limit is recorded here instead of being described as though it were not there.
        var checkedPath = CreateTestFilePath("handle_delete_swapped_early.txt");
        var actual = CreateTestFilePath("handle_delete_actual.txt");
        File.WriteAllText(actual, "what the name means by the time it is opened");

        // The swap: between a caller validating `checkedPath` and this call, the name came to mean
        // a different file.
        File.Move(actual, checkedPath);

        HandleBoundDelete.Delete(checkedPath);

        Assert.False(File.Exists(checkedPath),
            "the file the name resolved to at open time is what goes");
    }

    [Fact]
    public void AFileHeldOpenByAReader_ShouldStillBeDeleted()
    {
        Skip.IfNot(OperatingSystem.IsWindows(), "Disposition semantics are the Windows ones.");

        // A superseded document is exactly the sort of file something else may still have open —
        // that is why the delete failed in the first place. The disposition takes effect when the
        // last handle closes, so this must not throw.
        var path = CreateTestFilePath("handle_delete_open.txt");
        File.WriteAllText(path, "content");

        using (var reader = new FileStream(path, FileMode.Open, FileAccess.Read,
                   FileShare.ReadWrite | FileShare.Delete))
        {
            HandleBoundDelete.Delete(path);
            Assert.Equal(7, reader.Length);
        }

        Assert.False(File.Exists(path), "the file survived the last handle closing");
    }

    [Fact]
    public void AMissingFile_ShouldReportAFailureRatherThanSucceedQuietly()
    {
        // The queue treats a throw as "still owed" and retries; reporting success for a file that
        // was never opened would drop the debt.
        var path = CreateTestFilePath("handle_delete_missing.txt");

        Assert.ThrowsAny<IOException>(() => HandleBoundDelete.Delete(path));
    }

    [Fact]
    public void TheQueue_ShouldUseTheHandleBoundDeleteByDefault()
    {
        // The seam exists for fixtures; production must not be left on the path-based one.
        var queue = new CleanupDebtQueue(CreateTestFilePath("handle_delete_queue.json"), [TestDir], Recovery);

        var target = CreateTestFilePath("handle_delete_swept.txt");
        File.WriteAllText(target, "content");
        queue.Record(target, "locked");

        Assert.Contains(target, queue.Sweep().Deleted);
        Assert.False(File.Exists(target));
    }

    [Fact]
    public void DeleteIf_WhenThePredicateRefuses_ShouldLeaveTheFileExactlyAsItWas()
    {
        // R20-REC02, and the measurement behind its Windows shape: the first version opened with
        // delete-on-close and tried to withdraw the disposition on refusal. The flag is a property
        // of the open and cannot be withdrawn — the refused file was gone anyway.
        var path = Path.Combine(Path.GetTempPath(), "hbd-refuse-" + Guid.NewGuid().ToString("N"));
        File.WriteAllText(path, "not the one you meant");
        try
        {
            var asked = false;
            var deleted = HandleBoundDelete.DeleteIf(path, stream =>
            {
                asked = true;
                Assert.True(stream.CanRead);
                return false;
            });

            Assert.True(asked, "the predicate was never consulted");
            Assert.False(deleted);
            Assert.True(File.Exists(path), "a refused delete removed the file");
            Assert.Equal("not the one you meant", File.ReadAllText(path));
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void DeleteIf_WhenThePredicatePasses_ShouldDeleteTheFileItJudged()
    {
        var path = Path.Combine(Path.GetTempPath(), "hbd-pass-" + Guid.NewGuid().ToString("N"));
        File.WriteAllText(path, "exactly the one");

        var expectedDigest = Convert.ToHexString(SHA256.HashData("exactly the one"u8));
        var deleted = HandleBoundDelete.DeleteIf(path,
            stream => string.Equals(HandleBoundDelete.DigestOf(stream), expectedDigest,
                StringComparison.OrdinalIgnoreCase));

        Assert.True(deleted);
        Assert.False(File.Exists(path));
    }
}
