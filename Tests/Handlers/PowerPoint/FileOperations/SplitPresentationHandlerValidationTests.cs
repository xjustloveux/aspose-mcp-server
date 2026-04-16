using System.Runtime.Versioning;
using Aspose.Slides;
using Aspose.Slides.Export;
using AsposeMcpServer.Handlers.PowerPoint.FileOperations;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.FileOperations;

/// <summary>
///     Repro tests for bug 20260415-ppt-split-dos.
///     Demonstrates that SplitPresentationHandler enters an unbounded loop
///     (CPU + unbounded Presentation allocations) when <c>slidesPerFile</c>
///     is 0 or negative, because the handler never validates the step value
///     used by <c>for (var i = start; i &lt;= end; i += p.SlidesPerFile)</c>
///     at SplitPresentationHandler.cs:72.
///     Expected (per charter §6 correctness + §5 input validation): a
///     fast-fail <see cref="ArgumentException" />.
///     Actual: the call hangs indefinitely.
///     NOTE: These tests are authored only — they have not been executed.
///     The build is currently broken by bug 20260415-build-msb3552
///     (MSB3552 on **/*.resx). Once the build is restored, running these
///     tests should make the bug observable as a timeout.
/// </summary>
[SupportedOSPlatform("windows")]
public class SplitPresentationHandlerValidationTests : PptHandlerTestBase
{
    private readonly SplitPresentationHandler _handler = new();
    private readonly string _inputPath;

    public SplitPresentationHandlerValidationTests()
    {
        _inputPath = Path.Combine(TestDir, "dos_input.pptx");

        using var pres = new Presentation();
        pres.Slides.AddEmptySlide(pres.Slides[0].LayoutSlide);
        pres.Slides.AddEmptySlide(pres.Slides[0].LayoutSlide);
        pres.Save(_inputPath, SaveFormat.Pptx);
    }

    [SkippableFact]
    public void Execute_WithSlidesPerFileZero_ShouldThrow_NotHang()
    {
        SkipIfNotWindows();

        var outputDir = Path.Combine(TestDir, "split_dos_zero");
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "inputPath", _inputPath },
            { "outputDirectory", outputDir },
            { "slidesPerFile", 0 }
        });

        // Run on a background task so we can bound wall-clock time. If the
        // handler's loop is infinite the task never completes and the Wait
        // returns false, which we treat as a test failure (the DoS repro).
        var task = Task.Run(() => _handler.Execute(context, parameters));

        // Task.Wait throws AggregateException when the task is faulted
        // (i.e. the handler threw synchronously). A faulted task still
        // means it completed — it did not hang.
        bool completed;
        try
        {
            completed = task.Wait(TimeSpan.FromSeconds(5));
        }
        catch (AggregateException)
        {
            completed = true;
        }

        Assert.True(completed,
            "SplitPresentationHandler did not return within 5s with slidesPerFile=0 — " +
            "infinite-loop DoS confirmed (bug 20260415-ppt-split-dos).");

        // The expected terminal state is a thrown ArgumentException
        // surfaced through the task.
        var ex = Record.Exception(() =>
        {
            try
            {
                task.GetAwaiter().GetResult();
            }
            catch (AggregateException ae) when (ae.InnerException != null)
            {
                throw ae.InnerException;
            }
        });
        Assert.IsType<ArgumentException>(ex);
    }

    [SkippableFact]
    public void Execute_WithSlidesPerFileNegative_ShouldThrow_NotHang()
    {
        SkipIfNotWindows();

        var outputDir = Path.Combine(TestDir, "split_dos_neg");
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "inputPath", _inputPath },
            { "outputDirectory", outputDir },
            { "slidesPerFile", -1 }
        });

        var task = Task.Run(() => _handler.Execute(context, parameters));

        bool completed;
        try
        {
            completed = task.Wait(TimeSpan.FromSeconds(5));
        }
        catch (AggregateException)
        {
            completed = true;
        }

        Assert.True(completed,
            "SplitPresentationHandler did not return within 5s with slidesPerFile=-1 — " +
            "negative step causes non-terminating loop (bug 20260415-ppt-split-dos).");

        var ex = Record.Exception(() =>
        {
            try
            {
                task.GetAwaiter().GetResult();
            }
            catch (AggregateException ae) when (ae.InnerException != null)
            {
                throw ae.InnerException;
            }
        });
        Assert.IsType<ArgumentException>(ex);
    }

    // Edge-case: upper-bound guard — slidesPerFile=1001 is just above the 1000
    // ceiling enforced by the fix. Mirrors the PDF sibling's range validation.
    [SkippableFact]
    public void Execute_WithSlidesPerFileAboveCeiling_ShouldThrow()
    {
        SkipIfNotWindows();

        var outputDir = Path.Combine(TestDir, "split_dos_above_ceiling");
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "inputPath", _inputPath },
            { "outputDirectory", outputDir },
            { "slidesPerFile", 1001 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("slidesPerFile", ex.Message);
        Assert.Contains("1000", ex.Message);
    }

    // Edge-case: lower-bound guard — slidesPerFile=int.MinValue is the extreme
    // negative case; must fail fast with ArgumentException (not hang, not
    // overflow in the step arithmetic).
    [SkippableFact]
    public void Execute_WithSlidesPerFileIntMinValue_ShouldThrow()
    {
        SkipIfNotWindows();

        var outputDir = Path.Combine(TestDir, "split_dos_intmin");
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "inputPath", _inputPath },
            { "outputDirectory", outputDir },
            { "slidesPerFile", int.MinValue }
        });

        var task = Task.Run(() => _handler.Execute(context, parameters));

        bool completed;
        try
        {
            completed = task.Wait(TimeSpan.FromSeconds(5));
        }
        catch (AggregateException)
        {
            completed = true;
        }

        Assert.True(completed,
            "SplitPresentationHandler did not return within 5s with slidesPerFile=int.MinValue.");

        var ex = Record.Exception(() =>
        {
            try
            {
                task.GetAwaiter().GetResult();
            }
            catch (AggregateException ae) when (ae.InnerException != null)
            {
                throw ae.InnerException;
            }
        });
        Assert.IsType<ArgumentException>(ex);
    }
}
