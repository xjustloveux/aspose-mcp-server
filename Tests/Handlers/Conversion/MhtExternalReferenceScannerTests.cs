using System.Text;
using System.Text.RegularExpressions;
using AsposeMcpServer.Helpers;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Conversion;

/// <summary>
///     Guards the residual half of RB-30. Aspose.Pdf 23.10.0 exposes an external-resource hook on
///     <c>HtmlLoadOptions</c> only, so an MHT archive is the one conversion input whose fetching
///     cannot be intercepted. The scanner refuses an archive that references anything it does not
///     carry, before the converter opens it.
/// </summary>
public class MhtExternalReferenceScannerTests : TestBase
{
    /// <summary>
    ///     Writes an MHT archive whose single HTML part holds the supplied markup.
    /// </summary>
    /// <param name="fileName">File name to create under the test directory.</param>
    /// <param name="html">Markup for the HTML part.</param>
    /// <param name="base64">Whether to encode the part as base64 rather than 8bit.</param>
    /// <returns>The path that was written.</returns>
    private string WriteMht(string fileName, string html, bool base64 = false)
    {
        var path = CreateTestFilePath(fileName);
        var body = base64
            ? Convert.ToBase64String(Encoding.UTF8.GetBytes(html))
            : html;
        var encoding = base64 ? "base64" : "8bit";

        var mht = string.Join("\r\n",
            "From: <saved by test>",
            "Subject: fixture",
            "MIME-Version: 1.0",
            "Content-Type: multipart/related; boundary=\"----=_Boundary\"",
            "",
            "------=_Boundary",
            "Content-Type: text/html; charset=\"utf-8\"",
            $"Content-Transfer-Encoding: {encoding}",
            "Content-Location: http://localhost/fixture.html",
            "",
            body,
            "",
            "------=_Boundary--",
            "");
        File.WriteAllText(path, mht, Encoding.UTF8);
        return path;
    }

    /// <summary>
    ///     R5-S01: the parser's message owns the streams and buffers it filled from the archive,
    ///     and the scan used to return from it without releasing them, so they accumulated across
    ///     attacker-supplied inputs until a collection happened to run.
    ///     <para>
    ///         This is a source guard rather than a behavioural one, because the release is not
    ///         observable from outside: the parser closes the file itself, so a scan leaves the
    ///         archive replaceable whether the message was disposed or not — measured, with the
    ///         disposal removed. What can be checked is that every message this file loads is bound
    ///         to a <c>using</c>, which is what makes the release cover the refusal paths too:
    ///         <c>Take</c> throws on any cap, and every cap is reached inside the block.
    ///     </para>
    /// </summary>
    [Fact]
    public void EveryMessageTheScannerLoads_ShouldBeBoundToAUsing()
    {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);
        while (directory != null && !File.Exists(Path.Combine(directory.FullName, "AsposeMcpServer.csproj")))
            directory = directory.Parent;

        Assert.NotNull(directory);
        var source = File.ReadAllText(
            Path.Combine(directory.FullName, "Helpers", "MhtExternalReferenceScanner.cs"), Encoding.UTF8);

        var loads = Regex.Matches(source, @"(\w+)\s*=\s*MailMessage\.Load\(")
            .Select(m => m.Groups[1].Value)
            .ToList();
        var unreleased = loads
            .Where(name => !Regex.IsMatch(source, @"using\s*\(\s*" + Regex.Escape(name) + @"\s*\)")
                           && !Regex.IsMatch(source, @"using\s+var\s+" + Regex.Escape(name) + @"\s*="))
            .ToList();

        Assert.True(loads.Count > 0,
            "the scanner no longer loads a message; this guard is checking nothing");
        Assert.True(unreleased.Count == 0,
            "these loaded messages are never bound to a using, so they keep their streams and " +
            "buffers after the scan returns: " + string.Join(", ", unreleased));
    }

    [Fact]
    public void SelfContainedArchive_ShouldBeAccepted()
    {
        var path = WriteMht("selfcontained.mht",
            "<html><body><img src=\"cid:logo\"><p>hello</p></body></html>");

        MhtExternalReferenceScanner.EnsureSelfContained(path);

        Assert.Empty(MhtExternalReferenceScanner.FindRemoteReferences(path));
    }

    [Theory]
    [InlineData("<img src=\"http://169.254.169.254/latest/meta-data/\">")]
    [InlineData("<img src='https://internal.example.com/logo.png'>")]
    [InlineData("<link href=\"https://cdn.example.com/site.css\" rel=\"stylesheet\">")]
    [InlineData("<div style=\"background:url(https://example.com/bg.png)\"></div>")]
    [InlineData("<style>@import \"https://example.com/more.css\";</style>")]
    public void RemoteReference_ShouldBeRefused(string markup)
    {
        var path = WriteMht("remote.mht", $"<html><body>{markup}</body></html>");

        var ex = Assert.Throws<ArgumentException>(() =>
            MhtExternalReferenceScanner.EnsureSelfContained(path));

        Assert.Contains("references resources that are not contained", ex.Message);
    }

    [Fact]
    public void Base64EncodedPart_ShouldStillBeScanned()
    {
        var path = WriteMht("encoded.mht",
            "<html><body><img src=\"https://example.com/tracker.gif\"></body></html>", true);

        Assert.Throws<ArgumentException>(() => MhtExternalReferenceScanner.EnsureSelfContained(path));
    }

    [Fact]
    public void OptIn_ShouldSkipTheScanEntirely()
    {
        var path = WriteMht("optin.mht",
            "<html><body><img src=\"https://example.com/a.png\"></body></html>");

        var exception = Record.Exception(() =>
            MhtExternalReferenceScanner.EnsureSelfContained(path, true));

        Assert.Null(exception);
    }

    [Theory]
    [InlineData("cid:embedded-part")]
    [InlineData("data:image/png;base64,AAAA")]
    public void LocalSchemes_ShouldNotCountAsRemote(string uri)
    {
        var path = WriteMht("local.mht", $"<html><body><img src=\"{uri}\"></body></html>");

        Assert.Empty(MhtExternalReferenceScanner.FindRemoteReferences(path));
    }

    [Fact]
    public void EvaluationUriInSubject_ShouldNotBecomeAResourceReference()
    {
        const string uri =
            "http://www.aspose.com/corporate/purchase/end-user-license-agreement.aspx";
        var path = CreateTestFilePath("evaluation_uri_subject.mht");
        File.WriteAllText(path, string.Join("\r\n",
            $"Subject: {uri}",
            "MIME-Version: 1.0",
            "Content-Type: text/html; charset=utf-8",
            "Content-Transfer-Encoding: 8bit",
            "",
            "<html><body>harmless</body></html>",
            ""));

        Assert.Empty(MhtExternalReferenceScanner.FindRemoteReferences(path));
        MhtExternalReferenceScanner.EnsureSelfContained(path);
    }

    [Theory]
    [InlineData("X_Custom")]
    [InlineData("X.Test")]
    [InlineData("X!Test")]
    public void PrintableMimeFieldName_ShouldNotInvalidateAnOtherwiseSafeArchive(string fieldName)
    {
        var path = CreateTestFilePath("printable_mime_field_name.mht");
        File.WriteAllText(path, string.Join("\r\n",
            $"{fieldName}: harmless",
            "MIME-Version: 1.0",
            "Content-Type: text/html; charset=utf-8",
            "Content-Transfer-Encoding: 8bit",
            "",
            "<html><body>harmless</body></html>",
            ""));

        Assert.Empty(MhtExternalReferenceScanner.FindRemoteReferences(path));
        MhtExternalReferenceScanner.EnsureSelfContained(path);
    }

    [Fact]
    public void MultipleReferences_ShouldBeReportedWithoutDuplicates()
    {
        var path = WriteMht("multi.mht",
            "<html><body>" +
            "<img src=\"https://example.com/a.png\">" +
            "<img src=\"https://example.com/a.png\">" +
            "<img src=\"https://example.com/b.png\">" +
            "</body></html>");

        var found = MhtExternalReferenceScanner.FindRemoteReferences(path);

        Assert.Equal(2, found.Count);
    }

    [Fact]
    public void UnparsableFile_ShouldBeRefusedBecauseItCannotBeChecked()
    {
        // This assertion was the other way round until A-03. Letting an unreadable archive
        // through assumed the converter would reject it too, but the converter runs its own
        // parser over the same bytes and may well read something this one could not.
        var path = CreateTestFilePath("broken.mht");
        File.WriteAllText(path, "this is not an MHT archive");

        Assert.ThrowsAny<ArgumentException>(() => MhtExternalReferenceScanner.EnsureSelfContained(path));
        Assert.Empty(MhtExternalReferenceScanner.FindRemoteReferences(path));
    }

    [Fact]
    public void HeaderNamesAfterLeadingGarbage_ShouldNotMakeAFileAMimeArchive()
    {
        var path = CreateTestFilePath("smuggled_headers.mht");
        File.WriteAllText(path, string.Join("\r\n",
            "this leading line is not a MIME header",
            "MIME-Version: 1.0",
            "Content-Type: text/html; charset=utf-8",
            "",
            "<html><body>harmless</body></html>"));

        Assert.ThrowsAny<ArgumentException>(() => MhtExternalReferenceScanner.EnsureSelfContained(path));
        Assert.Empty(MhtExternalReferenceScanner.FindRemoteReferences(path));
    }

    [Fact]
    public void AValidFoldedContentTypeHeader_ShouldRemainAccepted()
    {
        var path = CreateTestFilePath("folded_header.mht");
        File.WriteAllText(path, string.Join("\r\n",
            "From: <saved by test>",
            "MIME-Version: 1.0",
            "Content-Type: multipart/related;",
            " boundary=\"----=_Boundary\"",
            "",
            "------=_Boundary",
            "Content-Type: text/html; charset=utf-8",
            "Content-Transfer-Encoding: 8bit",
            "",
            "<html><body>harmless</body></html>",
            "------=_Boundary--",
            ""));

        var exception = Record.Exception(() => MhtExternalReferenceScanner.EnsureSelfContained(path));

        Assert.Null(exception);
        Assert.Empty(MhtExternalReferenceScanner.FindRemoteReferences(path));
    }
}
