using AsposeMcpServer.Core.Session;
using AsposeMcpServer.Helpers.PowerPoint;

namespace AsposeMcpServer.Tests.Core.Session;

/// <summary>
///     Periodic auto-save writes to a staging file and then moves it over the recovery slot, so
///     the staging file has to be saveable in the session's own format.
///     <para>
///         The suffix used to be appended after the extension, giving <c>session.pptx.new</c>.
///         Save format is resolved from the extension, and
///         <see cref="PptSaveFormatResolver" /> refuses an unknown one, so every periodic
///         auto-save of a PowerPoint session threw. The throw was caught by the loop's own
///         <c>catch</c> and written to the log, so the feature reported nothing and simply never
///         worked (R2-C01).
///     </para>
/// </summary>
public class AutoSaveStagingPathTests
{
    /// <summary>
    ///     Every format a presentation session can hold must survive the staging round trip.
    /// </summary>
    /// <param name="extension">Extension of the session's recovery file.</param>
    [Theory]
    [InlineData(".pptx")]
    [InlineData(".pptm")]
    [InlineData(".ppt")]
    [InlineData(".potx")]
    public void StagingPath_ForAPresentation_ShouldStillResolveToASaveFormat(string extension)
    {
        var finalPath = Path.Combine(Path.GetTempPath(), "session_abc" + extension);

        var staging = DocumentSessionManager.BuildStagingPath(finalPath);

        Assert.Equal(extension, Path.GetExtension(staging));
        Assert.NotEqual(finalPath, staging);

        // The staging file is written through the same resolver as any other save.
        var resolved = PptSaveFormatResolver.Resolve(staging);
        Assert.Equal(PptSaveFormatResolver.Resolve(finalPath), resolved);
    }

    [Fact]
    public void StagingPath_ShouldStayBesideTheFileItReplaces()
    {
        var finalPath = Path.Combine(Path.GetTempPath(), "sessions", "session_abc.docx");

        var staging = DocumentSessionManager.BuildStagingPath(finalPath);

        Assert.Equal(Path.GetDirectoryName(finalPath), Path.GetDirectoryName(staging));
        Assert.Equal("session_abc.new.docx", Path.GetFileName(staging));
    }
}
