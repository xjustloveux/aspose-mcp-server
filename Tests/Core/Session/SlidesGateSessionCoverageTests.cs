using System.Reflection;
using Aspose.Slides;
using Aspose.Words;
using AsposeMcpServer.Core.Conversion;
using AsposeMcpServer.Core.Session;
using AsposeMcpServer.Helpers;
using AsposeMcpServer.Tests.Infrastructure;
using SaveFormat = Aspose.Slides.Export.SaveFormat;

namespace AsposeMcpServer.Tests.Core.Session;

/// <summary>
///     R9-S01: the session paths into Aspose.Slides must be behind the gate too.
///     <para>
///         The first gate was taken where a presentation is <em>constructed</em>, which left every
///         path that works on one it was handed: converting a session document, saving it —
///         explicitly, on close, on any of the three disconnect behaviours, or from the auto-save
///         timer — and releasing it on dispose. Two sessions could therefore be inside the library
///         at the same moment, which is the interleaving that made it raise
///         <c>Nullable object must have a value</c> from its own code.
///     </para>
///     <para>
///         Each test here holds the gate on this thread and shows the operation cannot proceed
///         until it is released, against a control that shows the same operation finishes promptly
///         when the gate is free. Without the control, an operation that was merely slow would look
///         gated.
///     </para>
/// </summary>
[Collection("SerialSlides")]
public class SlidesGateSessionCoverageTests : TestBase
{
    /// <summary>How long an ungated operation is given to finish.</summary>
    private static readonly TimeSpan Promptly = TimeSpan.FromSeconds(20);

    /// <summary>How long a gated operation is watched for, to show it does not proceed.</summary>
    private static readonly TimeSpan LongEnoughToShowItIsBlocked = TimeSpan.FromMilliseconds(400);

    /// <summary>Writes a small presentation for a session to open.</summary>
    /// <param name="name">Fixture file name.</param>
    /// <returns>The presentation path.</returns>
    private string APresentation(string name)
    {
        var path = CreateTestFilePath(name);
        using var slidesGate = SlidesGate.Enter();
        using var presentation = new Presentation();
        presentation.Slides.AddEmptySlide(presentation.LayoutSlides[0]);
        presentation.Save(path, SaveFormat.Pptx);
        return path;
    }

    /// <summary>
    ///     Runs an operation on its own thread and reports when it finished.
    /// </summary>
    /// <param name="operation">The operation to run.</param>
    /// <returns>An event set once the operation returns, and the thread running it.</returns>
    private static (ManualResetEventSlim Finished, Thread Worker, Func<Exception?> Failure)
        RunOnItsOwnThread(Action operation)
    {
        var finished = new ManualResetEventSlim(false);
        Exception? failure = null;

        var worker = new Thread(() =>
        {
            try
            {
                operation();
            }
            catch (Exception error)
            {
                failure = error;
            }
            finally
            {
                finished.Set();
            }
        }) { IsBackground = true };

        worker.Start();
        return (finished, worker, () => failure);
    }

    /// <summary>
    ///     Asserts an operation waits for the gate: blocked while this thread holds it, and
    ///     finished once it is released.
    /// </summary>
    /// <param name="what">What the operation is, for the failure message.</param>
    /// <param name="operation">The operation, which must be safe to run once.</param>
    private static void ShouldWaitForTheGate(string what, Action operation)
    {
        ManualResetEventSlim finished;
        Thread worker;
        Func<Exception?> failure;

        using (SlidesGate.Enter())
        {
            (finished, worker, failure) = RunOnItsOwnThread(operation);

            Assert.False(finished.Wait(LongEnoughToShowItIsBlocked),
                $"{what} entered Aspose.Slides while another thread held the gate; "
                + "it is not behind the gate.");
        }

        Assert.True(finished.Wait(Promptly),
            $"{what} did not finish after the gate was released.");
        worker.Join(Promptly);

        Assert.Null(failure());
        finished.Dispose();
    }

    /// <summary>
    ///     The control: the same operation finishes promptly when nothing holds the gate, so
    ///     "blocked" above means the gate and not slowness.
    /// </summary>
    /// <param name="what">What the operation is, for the failure message.</param>
    /// <param name="operation">The operation, which must be safe to run once.</param>
    private static void ShouldFinishPromptlyWhenTheGateIsFree(string what, Action operation)
    {
        var (finished, worker, failure) = RunOnItsOwnThread(operation);

        Assert.True(finished.Wait(Promptly), $"{what} did not finish with the gate free.");
        worker.Join(Promptly);

        Assert.Null(failure());
        finished.Dispose();
    }

    [Fact]
    public void SavingAPresentationSession_ShouldWaitForTheGate()
    {
        // One gate in SaveDocumentToFile covers explicit save, save on close, all three
        // disconnect behaviours and auto-save, which all funnel through it.
        var manager = SessionManager;
        var blocked = manager.OpenDocument(APresentation("gate_save_blocked.pptx"));
        var control = manager.OpenDocument(APresentation("gate_save_control.pptx"));

        ShouldWaitForTheGate("saving a presentation session",
            () => manager.SaveDocument(blocked));
        ShouldFinishPromptlyWhenTheGateIsFree("saving a presentation session",
            () => manager.SaveDocument(control));
    }

    [Fact]
    public void ConvertingASessionPresentation_ShouldWaitForTheGate()
    {
        // The session conversion path takes a presentation it was handed, so nothing about it
        // constructs one — which is exactly why the constructor-shaped guard could not see it.
        var manager = SessionManager;
        var blocked = manager.GetDocument<Presentation>(
            manager.OpenDocument(APresentation("gate_convert_blocked.pptx")));
        var control = manager.GetDocument<Presentation>(
            manager.OpenDocument(APresentation("gate_convert_control.pptx")));

        var options = new ConversionOptions { RecoveryDirectory = Path.GetTempPath(), AllowedBasePaths = [TestDir] };

        ShouldWaitForTheGate("converting a session presentation",
            () => DocumentConverter.ConvertPowerPointDocument(
                blocked, CreateTestFilePath("gate_convert_blocked.pdf"), "pdf", null, options));
        ShouldFinishPromptlyWhenTheGateIsFree("converting a session presentation",
            () => DocumentConverter.ConvertPowerPointDocument(
                control, CreateTestFilePath("gate_convert_control.pdf"), "pdf", null, options));
    }

    [Fact]
    public void TheAutoSaveTimer_ShouldWaitForTheGate()
    {
        // Driven through the timer's own callback rather than through SaveDocument, so this shows
        // the auto-save path itself is gated and not merely a path that resembles it.
        var manager = SessionManager;
        var sessionId = manager.OpenDocument(APresentation("gate_autosave.pptx"));
        manager.MarkDirty(sessionId);

        var callback = typeof(DocumentSessionManager)
            .GetMethod("AutoSaveDirtySessions", BindingFlags.Instance | BindingFlags.NonPublic);
        Assert.NotNull(callback);

        void AutoSave()
        {
            manager.MarkDirty(sessionId);
            callback.Invoke(manager, [null]);
        }

        ShouldWaitForTheGate("the auto-save timer", AutoSave);
        ShouldFinishPromptlyWhenTheGateIsFree("the auto-save timer", AutoSave);
    }

    [Fact]
    public void DiscardingAPresentationSession_ShouldWaitForTheGate()
    {
        // Discard skips the save and goes straight to disposing the presentation, which is itself
        // a re-entry into the library.
        var manager = SessionManager;
        var blocked = manager.OpenDocument(APresentation("gate_discard_blocked.pptx"));
        var control = manager.OpenDocument(APresentation("gate_discard_control.pptx"));

        ShouldWaitForTheGate("discarding a presentation session",
            () => manager.CloseDocument(blocked, true));
        ShouldFinishPromptlyWhenTheGateIsFree("discarding a presentation session",
            () => manager.CloseDocument(control, true));
    }

    [Fact]
    public void SavingAWordSession_ShouldNotWaitForTheGate()
    {
        // The gate exists for one library's defect. Word, Excel and PDF sessions must be
        // unaffected, or the cost of the mitigation is far larger than the defect.
        var manager = SessionManager;
        var path = CreateTestFilePath("gate_word.docx");
        var document = new Document();
        new DocumentBuilder(document).Writeln("body");
        document.Save(path);

        var sessionId = manager.OpenDocument(path);

        using (SlidesGate.Enter())
        {
            var (finished, worker, failure) = RunOnItsOwnThread(() => manager.SaveDocument(sessionId));

            Assert.True(finished.Wait(Promptly),
                "saving a Word session waited for the Slides gate; the gate is too wide.");
            worker.Join(Promptly);
            Assert.Null(failure());
            finished.Dispose();
        }
    }

    [Fact]
    public void AScopeReleasedOnAnotherThread_ShouldBeRefused()
    {
        // The nesting count is thread-static, so a hold that ends on a different thread would
        // decrement one that never entered and leave it unable to nest afterwards.
        var scope = SlidesGate.Enter();
        try
        {
            Exception? failure = null;
            var worker = new Thread(() =>
            {
                try
                {
                    scope.Dispose();
                }
                catch (Exception error)
                {
                    failure = error;
                }
            }) { IsBackground = true };

            worker.Start();
            Assert.True(worker.Join(Promptly));

            Assert.IsType<InvalidOperationException>(failure);
            Assert.Contains("different thread", failure!.Message, StringComparison.Ordinal);
        }
        finally
        {
            scope.Dispose();
        }
    }
}
