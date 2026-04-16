using AsposeMcpServer.Errors.Pdf;

namespace AsposeMcpServer.Tests.Errors.Pdf;

/// <summary>
///     Unit tests for <see cref="PdfErrorMessageBuilder" />. Verifies that every sentinel
///     method returns a non-empty fixed string and that no variable content leaks through.
/// </summary>
public class PdfErrorMessageBuilderTests
{
    [Fact]
    public void ImageAccessError_IsNonEmpty()
    {
        Assert.NotEmpty(PdfErrorMessageBuilder.ImageAccessError());
    }

    [Fact]
    public void ImageAccessError_ReturnsExpectedFixedSentinel()
    {
        Assert.Equal("Image could not be accessed or decoded.", PdfErrorMessageBuilder.ImageAccessError());
    }

    [Fact]
    public void ImageAccessError_DoesNotContainExceptionOrStackTrace()
    {
        var msg = PdfErrorMessageBuilder.ImageAccessError();
        Assert.DoesNotContain("Exception", msg, StringComparison.Ordinal);
        Assert.DoesNotContain("   at ", msg, StringComparison.Ordinal);
    }

    [Fact]
    public void ImageAccessError_DoesNotContainPathSeparator()
    {
        var msg = PdfErrorMessageBuilder.ImageAccessError();
        Assert.DoesNotContain("/", msg, StringComparison.Ordinal);
        Assert.DoesNotContain(@"\", msg, StringComparison.Ordinal);
    }

    [Fact]
    public void ImageAccessError_IsStable_ReturnsSameValueOnRepeatCall()
    {
        Assert.Equal(PdfErrorMessageBuilder.ImageAccessError(), PdfErrorMessageBuilder.ImageAccessError());
    }
}
