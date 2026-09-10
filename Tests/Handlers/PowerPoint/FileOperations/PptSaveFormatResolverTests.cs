using Aspose.Slides.Export;
using AsposeMcpServer.Helpers.PowerPoint;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.FileOperations;

/// <summary>
///     Guards RB-10: every presentation save path used to pass <see cref="SaveFormat.Pptx" />
///     regardless of the destination extension, so editing a macro-enabled or legacy presentation
///     wrote PPTX content into a file that still claimed the original format.
/// </summary>
public class PptSaveFormatResolverTests
{
    [Theory]
    [InlineData("deck.pptx", SaveFormat.Pptx)]
    [InlineData("deck.ppt", SaveFormat.Ppt)]
    [InlineData("deck.pptm", SaveFormat.Pptm)]
    [InlineData("template.potx", SaveFormat.Potx)]
    [InlineData("template.potm", SaveFormat.Potm)]
    [InlineData("template.pot", SaveFormat.Pot)]
    [InlineData("deck.odp", SaveFormat.Odp)]
    [InlineData("show.ppsx", SaveFormat.Ppsx)]
    public void EditableExtensions_ShouldKeepTheirOwnFormat(string fileName, SaveFormat expected)
    {
        Assert.Equal(expected, PptSaveFormatResolver.Resolve(fileName));
    }

    [Fact]
    public void ExtensionMatching_ShouldBeCaseInsensitive()
    {
        Assert.Equal(SaveFormat.Pptm, PptSaveFormatResolver.Resolve("DECK.PPTM"));
    }

    [Fact]
    public void FullPath_ShouldResolveFromItsExtension()
    {
        var path = Path.Combine(Path.GetTempPath(), "sub dir", "quarterly.potx");

        Assert.Equal(SaveFormat.Potx, PptSaveFormatResolver.Resolve(path));
    }

    [Fact]
    public void MissingExtension_ShouldBeRejected()
    {
        var ex = Assert.Throws<ArgumentException>(() => PptSaveFormatResolver.Resolve("deck"));

        Assert.Contains("no extension", ex.Message);
    }

    [Theory]
    [InlineData("deck.docx")]
    [InlineData("deck.xlsx")]
    [InlineData("deck.txt")]
    public void UnsupportedExtension_ShouldBeRejectedRatherThanSilentlyRewritten(string fileName)
    {
        var ex = Assert.Throws<ArgumentException>(() => PptSaveFormatResolver.Resolve(fileName));

        Assert.Contains("cannot be saved as", ex.Message);
    }
}
