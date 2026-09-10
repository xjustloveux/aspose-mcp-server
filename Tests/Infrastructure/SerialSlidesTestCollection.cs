namespace AsposeMcpServer.Tests.Infrastructure;

/// <summary>
///     Collection for Aspose.Slides fixtures heavy enough to be worth keeping away from everything
///     else while they run.
///     <para>
///         One full licensed run died in
///         <c>SplitPresentationHandlerTests.Execute_OutputFileCountAtOrBelowCap_DoesNotTriggerGuard</c>
///         with <c>InvalidOperationException: Nullable object must have a value</c>. The TRX stack
///         put the dereference inside Aspose's own obfuscated licence and metering path, entered
///         from an ordinary <c>Shapes.AddAutoShape</c> call — nothing in this repository is on that
///         stack. It is a defect in the library, and it is not ours to fix; what is ours is the
///         concurrency we point at it. That test builds a 1,001-slide presentation while the rest
///         of the suite is exercising the same library from other threads, so it runs alone
///         (§19.10.1).
///     </para>
///     <para>
///         It recurred, in a different test and with a readable stack:
///         <c>ListPptFontsHandlerTests.Execute_ReturnsFontWithName</c> failed with
///         <c>Aspose.Slides.PptxReadException : Nullable object must have a value</c> raised from
///         <c>Aspose.Slides.Presentation..ctor(Stream)</c>. So it is not the licence path in
///         particular and not one heavy fixture: reading a presentation while other threads are
///         inside the same library is enough. Every Aspose.Slides test therefore lives here, which
///         means none of them runs beside another or beside the rest of the suite.
///     </para>
///     <para>
///         Still a mitigation, not a fix — the defect is in the library. What it costs is the
///         PowerPoint surface running serially; what it buys is a suite whose failures are the
///         code's rather than the vendor's.
///     </para>
/// </summary>
[CollectionDefinition("SerialSlides", DisableParallelization = true)]
public class SerialSlidesTestCollection;
