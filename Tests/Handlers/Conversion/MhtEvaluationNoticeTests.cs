using System.Text;
using AsposeMcpServer.Helpers;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Conversion;

/// <summary>
///     Guards a defect found while running the suite without a licence. The scanner reads the
///     decoded message body, and an unlicensed Aspose.Email injects its evaluation notice into that
///     body with a link to the Aspose licence page. Because the guard is fail-closed, the injected
///     link made every MHT conversion refuse on a server without a licence, naming a URL the caller
///     never wrote. The notice is library output rather than archive content, so it must not count
///     as an external reference in either mode.
/// </summary>
public class MhtEvaluationNoticeTests : TestBase
{
    /// <summary>Writes a self-contained MHT archive holding the supplied markup.</summary>
    /// <param name="fileName">File name under the test directory.</param>
    /// <param name="html">Markup for the single HTML part.</param>
    /// <param name="transferEncoding">MIME transfer encoding used for the HTML payload.</param>
    /// <returns>The path written.</returns>
    private string WriteMht(string fileName, string html, string transferEncoding = "8bit")
    {
        var path = CreateTestFilePath(fileName);
        var payload = transferEncoding switch
        {
            "base64" => Convert.ToBase64String(Encoding.UTF8.GetBytes(html)),
            "quoted-printable" => string.Concat(Encoding.UTF8.GetBytes(html)
                .Select(value => $"={value:X2}")),
            _ => html
        };
        File.WriteAllText(path, string.Join("\r\n",
            "From: <saved by test>",
            "Subject: fixture",
            "MIME-Version: 1.0",
            "Content-Type: multipart/related; boundary=\"----=_Boundary\"",
            "",
            "------=_Boundary",
            "Content-Type: text/html; charset=\"utf-8\"",
            $"Content-Transfer-Encoding: {transferEncoding}",
            "Content-Location: http://localhost/fixture.html",
            "",
            payload,
            "",
            "------=_Boundary--",
            ""));
        return path;
    }

    [Fact]
    public void SelfContainedArchive_ShouldStayAcceptedWithoutALicence()
    {
        var path = WriteMht("eval_selfcontained.mht",
            "<html><body><img src=\"cid:logo\"><p>hello</p></body></html>");

        var found = MhtExternalReferenceScanner.FindRemoteReferences(path);

        Assert.Empty(found);
        MhtExternalReferenceScanner.EnsureSelfContained(path);
    }

    [Fact]
    public void GenuineRemoteReference_ShouldStillBeRefused()
    {
        var path = WriteMht("eval_remote.mht",
            "<html><body><img src=\"https://example.com/tracker.png\"></body></html>");

        var found = MhtExternalReferenceScanner.FindRemoteReferences(path);

        Assert.Equal(["https://example.com/tracker.png"], found);
        Assert.ThrowsAny<ArgumentException>(() => MhtExternalReferenceScanner.EnsureSelfContained(path));
    }

    [Fact]
    public void AnAsposeLinkWrittenByTheArchive_ShouldStillCountAsARemoteReference()
    {
        // The exclusion covers the one measured notice URL only. An aspose.com link the archive
        // itself carries is ordinary remote content and is reported like any other.
        var path = WriteMht("eval_mixed.mht",
            "<html><body>"
            + "<a href=\"http://www.aspose.com/purchase/default.aspx\">buy</a>"
            + "<img src=\"https://example.com/tracker.png\">"
            + "</body></html>");

        var found = MhtExternalReferenceScanner.FindRemoteReferences(path);

        Assert.Equal(["http://www.aspose.com/purchase/default.aspx", "https://example.com/tracker.png"], found);
        Assert.ThrowsAny<ArgumentException>(() => MhtExternalReferenceScanner.EnsureSelfContained(path));
    }

    [Fact]
    public void TheExactEvaluationNoticeUriWrittenByTheArchive_ShouldStillBeRefused()
    {
        // The library may add this URI after parsing, but the same bytes supplied by the caller
        // are an ordinary remote reference. Treating the value alone as provenance lets an input
        // opt itself out of the external-resource guard.
        const string uri =
            "http://www.aspose.com/corporate/purchase/end-user-license-agreement.aspx";
        var path = WriteMht("eval_exact_uri.mht",
            $"<html><body><img src=\"{uri}\"></body></html>");

        var found = MhtExternalReferenceScanner.FindRemoteReferences(path);

        Assert.Contains(uri, found);
        Assert.ThrowsAny<ArgumentException>(() => MhtExternalReferenceScanner.EnsureSelfContained(path));
    }

    [Theory]
    [InlineData("8bit")]
    [InlineData("base64")]
    [InlineData("quoted-printable")]
    public void TheCompleteEvaluationNoticeWrittenByTheArchive_ShouldStillBeRefused(string transferEncoding)
    {
        const string uri =
            "http://www.aspose.com/corporate/purchase/end-user-license-agreement.aspx";
        const string notice =
            "<br><center><span style=\"color:red\">Evaluation Only. Created with Aspose.Email for .NET. "
            + "Copyright 2002-2022 Aspose Pty Ltd.</span></center><br><center><a href=\""
            + uri
            + "\">View EULA Online</a></center><hr><br>";
        var path = WriteMht($"eval_complete_notice_{transferEncoding}.mht",
            $"<html><body>{notice}</body></html>", transferEncoding);

        var found = MhtExternalReferenceScanner.FindRemoteReferences(path);

        Assert.Contains(uri, found);
        Assert.ThrowsAny<ArgumentException>(() => MhtExternalReferenceScanner.EnsureSelfContained(path));
    }

    [Fact]
    public void ACompleteNoticeInAnotherHtmlView_ShouldStillBeRefused()
    {
        const string uri =
            "http://www.aspose.com/corporate/purchase/end-user-license-agreement.aspx";
        const string notice =
            "<br><center><span style=\"color:red\">Evaluation Only. Created with Aspose.Email for .NET. "
            + "Copyright 2002-2022 Aspose Pty Ltd.</span></center><br><center><a href=\""
            + uri
            + "\">View EULA Online</a></center><hr><br>";
        var path = CreateTestFilePath("eval_notice_alternate_view.mht");
        File.WriteAllText(path, string.Join("\r\n",
            "From: <saved by test>",
            "Subject: fixture",
            "MIME-Version: 1.0",
            "Content-Type: multipart/related; boundary=\"----=_Boundary\"",
            "",
            "------=_Boundary",
            "Content-Type: text/html; charset=\"utf-8\"",
            "Content-Transfer-Encoding: 8bit",
            "Content-Location: http://localhost/first.html",
            "",
            "<html><body>harmless</body></html>",
            "",
            "------=_Boundary",
            "Content-Type: text/html; charset=\"utf-8\"",
            "Content-Transfer-Encoding: 8bit",
            "Content-Location: http://localhost/second.html",
            "",
            $"<html><body>{notice}</body></html>",
            "",
            "------=_Boundary--",
            ""));

        var found = MhtExternalReferenceScanner.FindRemoteReferences(path);

        Assert.Contains(uri, found);
        Assert.ThrowsAny<ArgumentException>(() => MhtExternalReferenceScanner.EnsureSelfContained(path));
    }

    [Fact]
    public void AQuotedPrintablePseudoBoundaryBeforeTheNotice_ShouldNotHideCallerProvenance()
    {
        const string uri =
            "http://www.aspose.com/corporate/purchase/end-user-license-agreement.aspx";
        const string notice =
            "<br><center><span style=\"color:red\">Evaluation Only. Created with Aspose.Email for .NET. "
            + "Copyright 2002-2022 Aspose Pty Ltd.</span></center><br><center><a href=\""
            + uri
            + "\">View EULA Online</a></center><hr><br>";
        var encodedNotice = string.Concat(Encoding.UTF8.GetBytes(notice)
            .Select(value => $"={value:X2}"));
        var path = CreateTestFilePath("eval_notice_pseudo_boundary.mht");
        File.WriteAllText(path, string.Join("\r\n",
            "From: <saved by test>",
            "MIME-Version: 1.0",
            "Content-Type: multipart/related; boundary=\"----=_RealBoundary\"",
            "",
            "------=_RealBoundary",
            "Content-Type: text/html; charset=utf-8",
            "Content-Transfer-Encoding: quoted-printable",
            "",
            "<html><body>",
            "--not-the-real-boundary",
            encodedNotice,
            "</body></html>",
            "------=_RealBoundary--",
            ""));

        var found = MhtExternalReferenceScanner.FindRemoteReferences(path);

        Assert.Contains(uri, found);
        Assert.ThrowsAny<ArgumentException>(() => MhtExternalReferenceScanner.EnsureSelfContained(path));
    }

    [Fact]
    public void APseudoBoundaryRepeatedBeforeAndInsideThePart_ShouldNotHideCallerProvenance()
    {
        const string uri =
            "http://www.aspose.com/corporate/purchase/end-user-license-agreement.aspx";
        const string notice =
            "<br><center><span style=\"color:red\">Evaluation Only. Created with Aspose.Email for .NET. "
            + "Copyright 2002-2022 Aspose Pty Ltd.</span></center><br><center><a href=\""
            + uri
            + "\">View EULA Online</a></center><hr><br>";
        var encodedNotice = string.Concat(Encoding.UTF8.GetBytes(
                $"<html><body>BEFORE{notice}AFTER</body></html>")
            .Select(value => $"={value:X2}"));
        var path = CreateTestFilePath("eval_notice_repeated_pseudo_boundary.mht");
        File.WriteAllText(path, string.Join("\r\n",
            "From: <saved by test>",
            "MIME-Version: 1.0",
            "Content-Type: multipart/related; boundary=\"----=_RealBoundary\"",
            "",
            "------=_RealBoundary",
            "Content-Type: text/html; charset=utf-8",
            "--fake",
            "Content-Transfer-Encoding: quoted-printable",
            "",
            "harmless",
            "--fake",
            encodedNotice,
            "------=_RealBoundary--",
            ""));

        var found = MhtExternalReferenceScanner.FindRemoteReferences(path);

        Assert.Contains(uri, found);
        Assert.ThrowsAny<ArgumentException>(() => MhtExternalReferenceScanner.EnsureSelfContained(path));
    }

    [Fact]
    public void ATransportPaddedBoundary_ShouldNotHideCallerProvenance()
    {
        const string uri =
            "http://www.aspose.com/corporate/purchase/end-user-license-agreement.aspx";
        const string notice =
            "<br><center><span style=\"color:red\">Evaluation Only. Created with Aspose.Email for .NET. "
            + "Copyright 2002-2022 Aspose Pty Ltd.</span></center><br><center><a href=\""
            + uri
            + "\">View EULA Online</a></center><hr><br>";
        var payload = Convert.ToBase64String(Encoding.UTF8.GetBytes(
            $"<html><body>{notice}</body></html>"));
        var path = CreateTestFilePath("eval_notice_transport_padding.mht");
        File.WriteAllText(path, string.Join("\r\n",
            "From: <saved by test>",
            "MIME-Version: 1.0",
            "Content-Type: multipart/related; boundary=b",
            "",
            "--b   ",
            "Content-Type: text/html; charset=utf-8",
            "Content-Transfer-Encoding: base64",
            "",
            payload,
            "--b--\t",
            ""));

        var found = MhtExternalReferenceScanner.FindRemoteReferences(path);

        Assert.Contains(uri, found);
        Assert.ThrowsAny<ArgumentException>(() => MhtExternalReferenceScanner.EnsureSelfContained(path));
    }

    [Fact]
    public void AUtf16Base64Notice_ShouldNotHideCallerProvenance()
    {
        const string uri =
            "http://www.aspose.com/corporate/purchase/end-user-license-agreement.aspx";
        var payload = Convert.ToBase64String(Encoding.Unicode.GetBytes(
            $"<html><body><a href=\"{uri}\">View EULA Online</a></body></html>"));
        var path = CreateTestFilePath("eval_notice_utf16_base64.mht");
        File.WriteAllText(path, string.Join("\r\n",
            "From: <saved by test>",
            "MIME-Version: 1.0",
            "Content-Type: text/html; charset=utf-16le",
            "Content-Transfer-Encoding: base64",
            "",
            payload,
            ""));

        var found = MhtExternalReferenceScanner.FindRemoteReferences(path);

        Assert.Contains(uri, found);
        Assert.ThrowsAny<ArgumentException>(() => MhtExternalReferenceScanner.EnsureSelfContained(path));
    }

    [Fact]
    public void AUtf32Base64Notice_ShouldFailClosedWhenCallerProvenanceCannotBeProved()
    {
        const string uri =
            "http://www.aspose.com/corporate/purchase/end-user-license-agreement.aspx";
        const string notice =
            "<br><center><span style=\"color:red\">Evaluation Only. Created with Aspose.Email for .NET. "
            + "Copyright 2002-2022 Aspose Pty Ltd.</span></center><br><center><a href=\""
            + uri
            + "\">View EULA Online</a></center><hr><br>";
        var payload = Convert.ToBase64String(Encoding.UTF32.GetBytes(
            $"<html><body>{notice}</body></html>"));
        var path = CreateTestFilePath("eval_notice_utf32_base64.mht");
        File.WriteAllText(path, string.Join("\r\n",
            "From: <saved by test>",
            "MIME-Version: 1.0",
            "Content-Type: text/html; charset=utf-32",
            "Content-Transfer-Encoding: base64",
            "",
            payload,
            ""));

        var found = MhtExternalReferenceScanner.FindRemoteReferences(path);

        Assert.Contains(uri, found);
        Assert.ThrowsAny<ArgumentException>(() => MhtExternalReferenceScanner.EnsureSelfContained(path));
    }

    [SkippableFact]
    public void AnArchiveWithExcessiveLineCount_ShouldFailClosedDuringProvenanceInspection()
    {
        Skip.IfNot(IsEvaluationMode(AsposeLibraryType.Email),
            "The fail-closed result is observable only when Aspose.Email injects its notice");

        const string uri =
            "http://www.aspose.com/corporate/purchase/end-user-license-agreement.aspx";
        var path = CreateTestFilePath("eval_notice_excessive_lines.mht");
        File.WriteAllText(path, string.Join("\r\n",
            "From: <saved by test>",
            "MIME-Version: 1.0",
            "Content-Type: text/html; charset=utf-8",
            "Content-Transfer-Encoding: 8bit",
            "",
            "<html><body>",
            new string('\n', 100_001),
            "</body></html>",
            ""));

        var found = MhtExternalReferenceScanner.FindRemoteReferences(path);

        Assert.Contains(uri, found);
        Assert.ThrowsAny<ArgumentException>(() => MhtExternalReferenceScanner.EnsureSelfContained(path));
    }
}
