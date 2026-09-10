using Aspose.Words;
using AsposeMcpServer.Core.Session;

namespace AsposeMcpServer.Tests.Core.Session;

/// <summary>
///     Saving and closing must exclude new operations, not merely wait for the current ones.
///     <para>
///         Draining spun until the active-user count reached zero and then proceeded. Nothing
///         stopped a new request from acquiring the instant it did, so a save could still capture a
///         document another operation had just started mutating, and a close could dispose one
///         still in use (R2-S09).
///     </para>
/// </summary>
public class SessionExclusiveAccessTests : IDisposable
{
    private readonly string _filePath;
    private readonly DocumentSessionManager _manager;
    private readonly string _testDir;

    /// <summary>
    ///     Creates a manager with sessions enabled and one document to open.
    /// </summary>
    public SessionExclusiveAccessTests()
    {
        _testDir = Path.Combine(Path.GetTempPath(), "SessionExclusive_" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(_testDir);

        _filePath = Path.Combine(_testDir, "fixture.docx");
        var doc = new Document();
        new DocumentBuilder(doc).Writeln("exclusive access fixture");
        doc.Save(_filePath);

        _manager = new DocumentSessionManager(new SessionConfig
        {
            Enabled = true,
            MaxSessions = 5,
            IdleTimeoutMinutes = 0,
            TempDirectory = Path.Combine(_testDir, "temp")
        });
    }

    /// <inheritdoc />
    public void Dispose()
    {
        _manager.Dispose();
        if (Directory.Exists(_testDir))
            try
            {
                Directory.Delete(_testDir, true);
            }
            catch (IOException)
            {
                // Best-effort cleanup; a locked file must not fail the run.
            }

        GC.SuppressFinalize(this);
    }

    [Fact]
    public void BeginExclusive_ShouldRefuseNewUsageUntilReleased()
    {
        var sessionId = _manager.OpenDocument(_filePath);
        var session = _manager.GetSession(sessionId);

        Assert.True(session.BeginExclusive());

        // This is the window the drain alone left open.
        Assert.Throws<InvalidOperationException>(() => session.AcquireUsage());

        session.EndExclusive();

        using var scope = session.AcquireUsage();
        Assert.NotNull(scope);
    }

    [Fact]
    public void BeginExclusive_WithAnOperationInFlight_ShouldNotComplete()
    {
        var sessionId = _manager.OpenDocument(_filePath);
        var session = _manager.GetSession(sessionId);

        using (session.AcquireUsage())
        {
            Assert.False(session.BeginExclusive(50));
        }

        // A failed attempt must leave the session usable rather than permanently closed.
        Assert.True(session.BeginExclusive());
        session.EndExclusive();
        using var scope = session.AcquireUsage();
        Assert.NotNull(scope);
    }

    [Fact]
    public void SaveDocument_ShouldLeaveTheSessionUsable()
    {
        // The save path holds the session exclusively; if it failed to release, every later
        // operation on that session would be refused.
        var sessionId = _manager.OpenDocument(_filePath);
        var outputPath = Path.Combine(_testDir, "saved.docx");

        _manager.SaveDocument(sessionId, outputPath);

        var session = _manager.GetSession(sessionId);
        using var scope = session.AcquireUsage();
        Assert.NotNull(scope);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void BeginExclusive_ShouldBeHeldByExactlyOneCaller()
    {
        // Acquisition used an unconditional write, so two racers both believed they held the
        // session and the first to finish released the other's exclusivity (R3-S04).
        var sessionId = _manager.OpenDocument(_filePath);
        var session = _manager.GetSession(sessionId);

        Assert.True(session.BeginExclusive());
        Assert.False(session.BeginExclusive(50));

        session.EndExclusive();
        Assert.True(session.BeginExclusive());
        session.EndExclusive();
    }

    [Fact]
    public async Task ConcurrentBeginExclusive_ShouldAdmitExactlyOne()
    {
        var sessionId = _manager.OpenDocument(_filePath);
        var session = _manager.GetSession(sessionId);

        // A barrier makes the interleaving deterministic rather than hoping for a collision.
        using var ready = new Barrier(4);
        var winners = 0;

        var contenders = Enumerable.Range(0, 3).Select(_ => Task.Run(() =>
        {
            ready.SignalAndWait();
            if (session.BeginExclusive(200)) Interlocked.Increment(ref winners);
        })).ToArray();

        ready.SignalAndWait();
        await Task.WhenAll(contenders);

        Assert.Equal(1, winners);
    }

    /// <summary>
    ///     The window R4-S05 names: a caller increments, sees the barrier and backs the increment
    ///     out — and the owner releases exclusivity before the caller decides what to do about it.
    ///     Re-reading the flag to decide returned a scope whose increment had already been undone,
    ///     so disposing it drove the active count below zero, and a negative count reads as idle
    ///     to the drain loop.
    ///     <para>
    ///         The window is a few instructions wide, so the release is performed from inside it
    ///         through the session's own seam rather than raced from another thread: a barrier
    ///         lines up the start of two threads and then leaves the outcome to chance, which is
    ///         not evidence either way.
    ///     </para>
    /// </summary>
    [Fact]
    public void AcquireUsage_WhenExclusivityIsReleasedInsideTheCheck_ShouldStillRefuse()
    {
        var sessionId = _manager.OpenDocument(_filePath);
        var session = _manager.GetSession(sessionId);

        // The barrier goes up after the caller has been admitted, and comes down again while it is
        // backing out - the exact interleaving the window is made of.
        session.OnUsageCheckPassed = () => Assert.True(session.BeginExclusive());
        session.OnExclusiveBarrierObserved = () => session.EndExclusive();
        try
        {
            Assert.Throws<InvalidOperationException>(() => session.AcquireUsage());
        }
        finally
        {
            session.OnUsageCheckPassed = null;
            session.OnExclusiveBarrierObserved = null;
        }

        // The refused caller must have left the count exactly as it found it.
        Assert.Equal(0, session.ActiveUsers);

        // And the session must still be usable, since exclusivity was released.
        using var scope = session.AcquireUsage();
        Assert.NotNull(scope);
        Assert.Equal(1, session.ActiveUsers);
    }

    /// <summary>
    ///     R4-S06: a request that already holds a session reference can sit between the lookup and
    ///     <c>AcquireUsage</c>. Closing only drained, so such a request could acquire after the
    ///     drain reported zero and then race the final save and dispose.
    ///     <para>
    ///         The acquire is attempted from inside the close, through the manager's seam, because
    ///         that span is the whole window: before it the session is live, after it the session
    ///         is disposed and any acquire fails for a different reason.
    ///     </para>
    /// </summary>
    [Fact]
    public void CloseDocument_ShouldRefuseAnAcquireTakenDuringTheClose()
    {
        var sessionId = _manager.OpenDocument(_filePath);
        var session = _manager.GetSession(sessionId);

        Exception? refusal = null;
        var attempted = false;

        _manager.OnSessionSealed = sealedSession =>
        {
            attempted = true;
            try
            {
                using var scope = sealedSession.AcquireUsage();
            }
            catch (Exception ex)
            {
                refusal = ex;
            }
        };

        try
        {
            _manager.CloseDocument(sessionId);
        }
        finally
        {
            _manager.OnSessionSealed = null;
        }

        Assert.True(attempted, "the fixture never reached the window it is testing");
        Assert.NotNull(refusal);
        Assert.IsType<InvalidOperationException>(refusal);

        // The reference the request was holding is no good afterwards either.
        Assert.ThrowsAny<Exception>(() => session.AcquireUsage());
    }

    /// <summary>
    ///     R4-S06: closing must tell "another operation holds this" apart from "the drain timed
    ///     out". One <c>false</c> meant both, so a close reported the second, skipped its save and
    ///     disposed the document under the save that held it.
    /// </summary>
    [Fact]
    public void BeginClosing_ShouldSayWhichThingStoppedIt()
    {
        var session = _manager.GetSession(_manager.OpenDocument(_filePath));

        Assert.True(session.BeginExclusive());
        Assert.Equal(SessionSealOutcome.HeldByAnotherOperation, session.BeginClosing(50));

        // The session is unregistered by the time a close reaches here, so a document the close
        // declined to take has to be released by whoever still holds it, or nothing ever will.
        Assert.True(session.DeferredReleaseArmed);
        session.EndExclusive();

        var busy = _manager.GetSession(_manager.OpenDocument(_filePath));
        using (busy.AcquireUsage())
        {
            Assert.Equal(SessionSealOutcome.ActiveOperationsRemain, busy.BeginClosing(50));
        }

        var idle = _manager.GetSession(_manager.OpenDocument(_filePath));
        Assert.Equal(SessionSealOutcome.Sealed, idle.BeginClosing(50));

        // A close that did take the document owns it; a holder that released it on the way out
        // would leave that close writing a disposed session.
        Assert.False(idle.DeferredReleaseArmed);

        // Sealing is permanent, and only the caller that sealed it owns the document.
        Assert.Equal(SessionSealOutcome.AlreadyClosing, idle.BeginClosing(50));
        Assert.False(idle.BeginExclusive(50));
        Assert.Throws<InvalidOperationException>(() => idle.AcquireUsage());
    }

    /// <summary>
    ///     R4-S06: a close that arrives while a save holds exclusivity but has not yet started
    ///     writing must wait for that save, not dispose the document out from under it.
    ///     <para>
    ///         The window is a few instructions wide, so the save is held inside it through
    ///         <c>OnExclusiveAcquired</c> and released once the close has raised its closing
    ///         barrier. Before the fix the close read the failed barrier as a drain timeout,
    ///         skipped its own save and disposed, and the waiting save then threw
    ///         <see cref="ObjectDisposedException" /> — so the edit reached neither the document
    ///         on disk nor an error the caller could act on.
    ///     </para>
    /// </summary>
    [Fact]
    public void CloseDocument_EnteringWhileASaveHoldsExclusivity_ShouldLetThatSaveFinish()
    {
        var sessionId = _manager.OpenDocument(_filePath);
        var session = _manager.GetSession(sessionId);

        const string edit = "written by the save the close raced";
        session.Execute(doc => new DocumentBuilder((Document)doc).Writeln(edit));
        session.IsDirty = true;

        using var saveIsInsideTheWindow = new ManualResetEventSlim(false);
        using var closeHasArrived = new ManualResetEventSlim(false);

        session.OnExclusiveAcquired = () =>
        {
            saveIsInsideTheWindow.Set();
            closeHasArrived.Wait(TimeSpan.FromSeconds(10));
        };
        session.OnClosingBarrierRaised = () => closeHasArrived.Set();

        Exception? saveFailure = null;
        var saving = new Thread(() =>
        {
            try
            {
                _manager.SaveDocument(sessionId);
            }
            catch (Exception ex)
            {
                saveFailure = ex;
            }
        });

        try
        {
            saving.Start();
            Assert.True(saveIsInsideTheWindow.Wait(TimeSpan.FromSeconds(10)),
                "the fixture never reached the window it is testing");

            _manager.CloseDocument(sessionId);
            Assert.True(saving.Join(TimeSpan.FromSeconds(10)), "the save never finished");
        }
        finally
        {
            closeHasArrived.Set();
            session.OnExclusiveAcquired = null;
            session.OnClosingBarrierRaised = null;
            saving.Join(TimeSpan.FromSeconds(10));
        }

        Assert.Null(saveFailure);

        // One complete save: the edit is on disk, and the session was no longer dirty by the time
        // the close looked, so it had nothing left to write.
        Assert.False(session.IsDirty);
        Assert.Contains(edit, new Document(_filePath).GetText(), StringComparison.Ordinal);

        // The close still reclaimed the document and still sealed the session.
        Assert.False(session.DisposedWithActiveUsers);
        Assert.True(session.IsDisposed);
        Assert.ThrowsAny<Exception>(() => session.AcquireUsage());
    }

    /// <summary>
    ///     R4-S06: a close that cannot wait out the operations already running unregisters the
    ///     session and leaves the document to them. Nothing could then release it — the session is
    ///     unreachable, and <c>ReleaseUsage</c> only decremented a counter — so the native document
    ///     had no owner at all. The last operation out takes it, and exactly one path may.
    /// </summary>
    [Fact]
    public void CloseDocument_WithAnOperationStillRunning_ShouldLeaveTheDocumentToTheLastOneOut()
    {
        var sessionId = _manager.OpenDocument(_filePath);
        var session = _manager.GetSession(sessionId);

        // The ordered sequence: an operation is in flight before the close begins.
        var scope = session.AcquireUsage();
        Assert.Equal(1, session.ActiveUsers);

        _manager.ClosingDrainTimeoutMs = 50;
        _manager.CloseDocument(sessionId);

        // The close has unregistered the session and returned without taking the document.
        Assert.Throws<KeyNotFoundException>(() => _manager.GetSession(sessionId));
        Assert.True(session.DeferredReleaseArmed);
        Assert.False(session.IsDisposed);

        // The operation finishes last, so it is the one that releases the document.
        scope.Dispose();
        Assert.Equal(0, session.ActiveUsers);
        Assert.True(session.IsDisposed);

        // And nobody else may: the ownership is claimed exactly once.
        Assert.False(session.TryClaimDisposeOwnership());
    }

    /// <summary>
    ///     The other two ways of being last: the close itself, and an exclusive holder the close
    ///     declined to wait for. Neither may dispose a document the other already owns.
    /// </summary>
    [Fact]
    public void DisposeOwnership_ShouldBeClaimableExactlyOnce()
    {
        var sealedSession = _manager.GetSession(_manager.OpenDocument(_filePath));
        Assert.Equal(SessionSealOutcome.Sealed, sealedSession.BeginClosing(50));
        Assert.False(sealedSession.DeferredReleaseArmed);
        Assert.True(sealedSession.TryClaimDisposeOwnership());
        Assert.False(sealedSession.TryClaimDisposeOwnership());

        var held = _manager.GetSession(_manager.OpenDocument(_filePath));
        Assert.True(held.BeginExclusive());
        Assert.Equal(SessionSealOutcome.HeldByAnotherOperation, held.BeginClosing(50));
        Assert.True(held.DeferredReleaseArmed);
        Assert.True(held.TryClaimDisposeOwnership());
        Assert.False(held.TryClaimDisposeOwnership());
        held.EndExclusive();
    }

    /// <summary>
    ///     R7-S01: the last operation can finish while the close is still deciding.
    ///     <para>
    ///         The deferred-release flag was written only after that decision, so a releaser that
    ///         got there first read it as unset and left. The close then armed a flag nobody would
    ///         read again and the document was never released. The release is driven from inside
    ///         the close's own wait, which is the only place that ordering can be fixed.
    ///     </para>
    /// </summary>
    [Fact]
    public void BeginClosing_WhenTheLastUsageEndsWhileTheCloseIsDeciding_ShouldStillReleaseIt()
    {
        var session = _manager.GetSession(_manager.OpenDocument(_filePath));
        var scope = session.AcquireUsage();

        // The exact window: the close has given up waiting and is about to hand the document
        // over, and the last operation finishes right there — before the flag is written.
        session.OnCloseAboutToDefer = () =>
        {
            session.OnCloseAboutToDefer = null;
            scope.Dispose();
        };

        try
        {
            var outcome = session.BeginClosing(50);

            Assert.Equal(SessionSealOutcome.ActiveOperationsRemain, outcome);
            Assert.Equal(0, session.ActiveUsers);
            Assert.True(session.IsDisposed,
                "the last operation had already gone, so the hand-over was to nobody");
            Assert.False(session.TryClaimDisposeOwnership());
        }
        finally
        {
            session.OnCloseAboutToDefer = null;
        }
    }

    /// <summary>
    ///     R7-S01, the other ordering: an exclusive holder the close declined to wait for releases
    ///     the barrier in the same window.
    /// </summary>
    [Fact]
    public void BeginClosing_WhenAnExclusiveHolderEndsWhileTheCloseIsDeciding_ShouldStillReleaseIt()
    {
        var session = _manager.GetSession(_manager.OpenDocument(_filePath));
        Assert.True(session.BeginExclusive());

        // The holder lets go in the same window: after the close gave up waiting for it, and
        // before the flag that would have told it to release the document is written.
        session.OnCloseAboutToDefer = () =>
        {
            session.OnCloseAboutToDefer = null;
            session.EndExclusive();
        };

        try
        {
            var outcome = session.BeginClosing(50);

            Assert.Equal(SessionSealOutcome.HeldByAnotherOperation, outcome);
            Assert.True(session.IsDisposed,
                "the holder had already released, so the hand-over was to nobody");
            Assert.False(session.TryClaimDisposeOwnership());
        }
        finally
        {
            session.OnCloseAboutToDefer = null;
        }
    }

    /// <summary>
    ///     R8-S03: the exclusive holder and the close can each be past their own write and about to
    ///     read the other's.
    ///     <para>
    ///         The close writes the deferred-release flag and then reads the counts; the holder
    ///         wrote the barrier down and the caller then read the flag. A store followed by a load
    ///         of a different location may be reordered, so each side could see the other's old
    ///         value and neither would dispose — an armed session with no owner. Both sides now use
    ///         an interlocked exchange, which carries a full fence, and the holder's settle lives
    ///         inside <c>EndExclusive</c> rather than at each call site.
    ///     </para>
    ///     <para>
    ///         A fence cannot be observed from a test; what this drives is the interleaving the
    ///         fence exists to make safe, and the rule that exactly one side disposes.
    ///     </para>
    /// </summary>
    [Fact]
    public async Task AnExclusiveReleaseInsideTheClosesArming_ShouldStillReleaseTheDocument()
    {
        var session = _manager.GetSession(_manager.OpenDocument(_filePath));
        Assert.True(session.BeginExclusive(50));

        using var closeIsAboutToArm = new ManualResetEventSlim(false);
        using var holderHasReleased = new ManualResetEventSlim(false);
        using var closeHasArmed = new ManualResetEventSlim(false);

        // The close has given up waiting for the barrier and is at the point of handing the
        // document over. It says so, then waits for the holder to put the barrier down — so the
        // arming and the release are both in flight at once, which is the window.
        session.OnCloseAboutToDefer = () =>
        {
            session.OnCloseAboutToDefer = null;
            closeIsAboutToArm.Set();
            holderHasReleased.Wait(TimeSpan.FromSeconds(5));
        };

        // The holder pauses immediately after putting the barrier down, until the close has armed
        // and settled, so its own settle runs last.
        session.OnExclusiveReleased = () =>
        {
            session.OnExclusiveReleased = null;
            holderHasReleased.Set();
            closeHasArmed.Wait(TimeSpan.FromSeconds(5));
        };

        var closing = Task.Run(() =>
        {
            var outcome = session.BeginClosing(50);
            closeHasArmed.Set();
            return outcome;
        });

        // Releasing before the close reaches its arming would simply let the close take the
        // barrier, which is a different story than the one under test.
        Assert.True(closeIsAboutToArm.Wait(TimeSpan.FromSeconds(10)),
            "the close never reached the point where it hands the document over");

        session.EndExclusive();

        var outcome = await closing.WaitAsync(TimeSpan.FromSeconds(10));

        Assert.Equal(SessionSealOutcome.HeldByAnotherOperation, outcome);
        Assert.True(session.DeferredReleaseArmed);
        Assert.True(session.IsDisposed,
            "both sides were past their own write and read the other's, so one of them had to be "
            + "the last one out");
    }

    /// <summary>
    ///     The same race without seams, run repeatedly: no session may be left armed and alive.
    /// </summary>
    /// <summary>
    ///     R8-S03, the other ordering: the close arms and looks while the barrier is still up, so
    ///     only the holder can be the last one out.
    ///     <para>
    ///         Both release sites used to do this settle themselves, which worked only for as long
    ///         as every future caller remembered to copy it. It now lives in
    ///         <c>EndExclusive</c>, and this is the interleaving that needs it.
    ///     </para>
    /// </summary>
    [Fact]
    public async Task AnExclusiveReleaseAfterTheCloseHasLooked_ShouldBeTheOneThatDisposes()
    {
        var session = _manager.GetSession(_manager.OpenDocument(_filePath));
        Assert.True(session.BeginExclusive(50));

        var closing = Task.Run(() => session.BeginClosing(50));

        Assert.Equal(SessionSealOutcome.HeldByAnotherOperation,
            await closing.WaitAsync(TimeSpan.FromSeconds(10)));

        // The close has armed and seen the barrier still up, so it left the document here. Nothing
        // else can reach the session now.
        Assert.True(session.DeferredReleaseArmed);
        Assert.False(session.IsDisposed);

        session.EndExclusive();

        Assert.True(session.IsDisposed,
            "the close handed the document to the exclusive holder, so putting the barrier down "
            + "is what has to release it");
    }

    [Fact]
    public async Task RepeatedExclusiveReleasesRacingACloses_ShouldNeverLeaveAnArmedSession()
    {
        for (var attempt = 0; attempt < 60; attempt++)
        {
            var sessionId = _manager.OpenDocument(_filePath);
            var session = _manager.GetSession(sessionId);
            Assert.True(session.BeginExclusive(50));

            using var ready = new Barrier(2);
            var closing = Task.Run(() =>
            {
                ready.SignalAndWait();
                return session.BeginClosing(20);
            });

            ready.SignalAndWait();
            session.EndExclusive();
            await closing.WaitAsync(TimeSpan.FromSeconds(10));

            Assert.False(session is { DeferredReleaseArmed: true, IsDisposed: false },
                $"attempt {attempt} left an armed session that nothing disposed");

            // The registry keeps a session until it is closed through the manager, and there is a
            // per-user limit; freeing the slot is what lets this run more than a handful of times.
            try
            {
                _manager.CloseDocument(sessionId, true);
            }
            catch (Exception error) when (error is KeyNotFoundException or ObjectDisposedException
                                              or InvalidOperationException)
            {
                // Already gone, which is the outcome under test.
            }
        }
    }
}
