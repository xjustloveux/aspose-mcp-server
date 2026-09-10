using AsposeMcpServer.Core.Conversion;
using AsposeMcpServer.Helpers;
using AsposeMcpServer.Tests.Infrastructure;
using Xunit.Sdk;

namespace AsposeMcpServer.Tests.Handlers.Conversion;

/// <summary>
///     R19-CNV01: the scan and the parse have to be about the same bytes.
///     <para>
///         <see cref="MhtExternalReferenceScanner" /> opened the caller's path to decide whether
///         converting the document would make this server fetch something, and then every loader
///         opened that same path again. For these formats the scan is the only control there is —
///         the code says so — so a file replaced between the two opens was converted having never
///         been scanned.
///     </para>
///     <para>
///         The swap is driven through <c>ImmutableInputCopy.AfterCopy</c>, which fires at exactly
///         the moment the race needs: the copy has been taken and nothing has read anything yet.
///         The first test proves the seam is real by swapping in the other direction — if the
///         conversion still read the caller's path, that swap would be seen.
///     </para>
/// </summary>
[Collection("SerialStaticSeams")]
public class ConversionInputSwapTests : TestBase
{
    /// <summary>Writes an MHT archive holding the supplied markup.</summary>
    /// <param name="fileName">File name under the test directory.</param>
    /// <param name="html">Markup for the single HTML part.</param>
    /// <returns>The path written.</returns>
    private string WriteMht(string fileName, string html)
    {
        var path = CreateTestFilePath(fileName);
        File.WriteAllText(path, string.Join("\r\n",
            "From: <saved by test>",
            "Subject: fixture",
            "MIME-Version: 1.0",
            "Content-Type: multipart/related; boundary=\"----=_Boundary\"",
            "",
            "------=_Boundary",
            "Content-Type: text/html; charset=\"utf-8\"",
            "Content-Transfer-Encoding: 8bit",
            "Content-Location: http://localhost/fixture.html",
            "",
            html,
            "",
            "------=_Boundary--",
            ""));
        return path;
    }

    [Fact]
    public void TheSwapSeam_ShouldFireBetweenTheCopyAndTheScan()
    {
        // The positive control for the mechanism the two tests below rely on. Without this, a
        // fixture that "proves" the swap has no effect proves only that the swap never happened.
        var input = WriteMht("seam_probe.mht", "<html><body><p>hello</p></body></html>");
        var output = CreateTestFilePath("seam_probe.pdf");

        var fired = new List<string>();
        var previous = ImmutableInputCopy.AfterCopy;
        ImmutableInputCopy.AfterCopy = source => fired.Add(source);

        try
        {
            DocumentConverter.ConvertToPdfFromSpecialFormat(input, output, [TestDir], false,
                TestDir);
        }
        catch (Exception exception) when (exception is not XunitException)
        {
            // Whether this machine can render the archive is beside the point; the seam fires
            // before any of that.
        }
        finally
        {
            ImmutableInputCopy.AfterCopy = previous;
        }

        Assert.Single(fired);
        Assert.Equal(input, fired[0]);
    }

    [Fact]
    public void AnInputSwappedAfterItIsCopied_ShouldNotBeTheOneConverted()
    {
        // The attack. The caller offers an archive that scans clean, and replaces it with one
        // naming a remote resource the moment the scan is over. Measured on the previous version
        // the loader opened the caller's path and got the second file.
        var input = WriteMht("swap_benign.mht", "<html><body><p>nothing to fetch</p></body></html>");
        var output = CreateTestFilePath("swap_benign.pdf");

        var hostile = WriteMht("swap_hostile.mht",
            "<html><body><img src=\"https://example.invalid/tracker.png\"></body></html>");

        var previous = ImmutableInputCopy.AfterCopy;
        ImmutableInputCopy.AfterCopy = source => File.Copy(hostile, source, true);

        Exception? thrown = null;
        try
        {
            DocumentConverter.ConvertToPdfFromSpecialFormat(input, output, [TestDir], false,
                TestDir);
        }
        catch (Exception exception) when (exception is not XunitException)
        {
            thrown = exception;
        }
        finally
        {
            ImmutableInputCopy.AfterCopy = previous;
        }

        // The swap really happened: the caller's file is the hostile one now.
        Assert.Contains("example.invalid", File.ReadAllText(input), StringComparison.Ordinal);

        // And it changed nothing. A refusal naming that host would mean the scanner had read the
        // caller's path rather than the copy — which is the defect. Any other failure is this
        // machine being unable to render an archive, which this test does not care about.
        Assert.DoesNotContain("example.invalid", thrown?.Message ?? string.Empty,
            StringComparison.Ordinal);
    }

    [Fact]
    public void AnInputThatNamesARemoteResource_ShouldStillBeRefused()
    {
        // The control that keeps the test above honest: reading a copy must not stop the scanner
        // seeing what is in it.
        var input = WriteMht("swap_control.mht",
            "<html><body><img src=\"https://example.invalid/tracker.png\"></body></html>");
        var output = CreateTestFilePath("swap_control.pdf");

        var failure = Assert.ThrowsAny<ArgumentException>(() =>
            DocumentConverter.ConvertToPdfFromSpecialFormat(input, output, [TestDir], false,
                TestDir));

        Assert.Contains("example.invalid", failure.Message, StringComparison.Ordinal);
        Assert.False(File.Exists(output));
    }
}
