using Aspose.Pdf;
using Aspose.Pdf.Annotations;
using AsposeMcpServer.Tests.Infrastructure;
using AsposeMcpServer.Tools.Pdf;

namespace AsposeMcpServer.Tests.Tools.Pdf;

/// <summary>
///     Covers the PDF-link half of NEW-CONTRACT-02. <c>AddPdfLinkHandler</c> documents its geometry
///     as optional and asks for <c>GetOptional("width", 100.0)</c> and
///     <c>
///         GetOptional("height",
///         20.0)
///     </c>
///     , but the tool declared <c>double width = 0</c> and forwarded it unconditionally, so
///     the handler always received 0 and its own defaults could never be reached. Omitting the
///     geometry therefore produced a link with no area, which nothing can click, and reported
///     success.
/// </summary>
public class PdfLinkGeometryDefaultTests : TestBase
{
    /// <summary>Creates a one-page PDF to attach links to.</summary>
    /// <param name="fileName">File name under the test directory.</param>
    /// <returns>The path written.</returns>
    private string CreatePdf(string fileName)
    {
        var path = CreateTestFilePath(fileName);
        using var document = new Document();
        document.Pages.Add();
        document.Save(path);
        return path;
    }

    /// <summary>Reads the rectangle of the first link annotation on page 1.</summary>
    /// <param name="path">PDF to inspect.</param>
    /// <returns>The link's rectangle.</returns>
    private static Rectangle FirstLinkRect(string path)
    {
        using var document = new Document(path);
        var annotation = document.Pages[1].Annotations
            .OfType<LinkAnnotation>()
            .First();
        return annotation.Rect;
    }

    [Fact]
    public void Add_WithoutGeometry_ShouldUseTheDocumentedDefaultSize()
    {
        var tool = new PdfLinkTool();
        var path = CreatePdf("link_default.pdf");
        var outputPath = CreateTestFilePath("link_default_out.pdf");

        tool.Execute("add", path, outputPath: outputPath, pageIndex: 1, url: "https://example.com");

        var rect = FirstLinkRect(outputPath);
        Assert.Equal(100d, rect.Width, 3);
        Assert.Equal(20d, rect.Height, 3);
    }

    [Fact]
    public void Add_WithExplicitGeometry_ShouldUseIt()
    {
        var tool = new PdfLinkTool();
        var path = CreatePdf("link_explicit.pdf");
        var outputPath = CreateTestFilePath("link_explicit_out.pdf");

        tool.Execute("add", path, outputPath: outputPath, pageIndex: 1, url: "https://example.com",
            x: 10, y: 20, width: 200, height: 40);

        var rect = FirstLinkRect(outputPath);
        Assert.Equal(200d, rect.Width, 3);
        Assert.Equal(40d, rect.Height, 3);
        Assert.Equal(10d, rect.LLX, 3);
        Assert.Equal(20d, rect.LLY, 3);
    }

    [Fact]
    public void Add_WithOnlyWidth_ShouldKeepTheDefaultHeight()
    {
        var tool = new PdfLinkTool();
        var path = CreatePdf("link_partial.pdf");
        var outputPath = CreateTestFilePath("link_partial_out.pdf");

        tool.Execute("add", path, outputPath: outputPath, pageIndex: 1, url: "https://example.com",
            width: 300);

        var rect = FirstLinkRect(outputPath);
        Assert.Equal(300d, rect.Width, 3);
        Assert.Equal(20d, rect.Height, 3);
    }
}
