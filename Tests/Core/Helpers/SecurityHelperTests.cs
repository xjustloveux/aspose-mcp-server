using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Tests.Core.Helpers;

/// <summary>
///     Unit tests for SecurityHelper class
/// </summary>
public class SecurityHelperTests
{
    #region SanitizeFileName Tests

    [Fact]
    public void SanitizeFileName_WithValidName_ShouldReturnSame()
    {
        var result = SecurityHelper.SanitizeFileName("document.docx");

        Assert.Equal("document.docx", result);
    }

    [Fact]
    public void SanitizeFileName_WithNullOrEmpty_ShouldReturnFile()
    {
        Assert.Equal("file", SecurityHelper.SanitizeFileName(null!));
        Assert.Equal("file", SecurityHelper.SanitizeFileName(""));
        Assert.Equal("file", SecurityHelper.SanitizeFileName("   "));
    }

    [Fact]
    public void SanitizeFileName_WithPathTraversal_ShouldRemove()
    {
        var result = SecurityHelper.SanitizeFileName("..\\..\\etc\\passwd");

        Assert.DoesNotContain("..", result);
        Assert.DoesNotContain("\\", result);
    }

    [Fact]
    public void SanitizeFileName_WithForwardSlash_ShouldReplace()
    {
        var result = SecurityHelper.SanitizeFileName("path/to/file.txt");

        Assert.DoesNotContain("/", result);
    }

    [Fact]
    public void SanitizeFileName_WithColon_ShouldReplace()
    {
        var result = SecurityHelper.SanitizeFileName("C:file.txt");

        Assert.DoesNotContain(":", result);
    }

    [Fact]
    public void SanitizeFileName_WithLongName_ShouldTruncate()
    {
        var longName = new string('a', 300) + ".txt";

        var result = SecurityHelper.SanitizeFileName(longName);

        Assert.True(result.Length <= 255);
    }

    [Fact]
    public void SanitizeFileName_WithLeadingTrailingDots_ShouldTrim()
    {
        var result = SecurityHelper.SanitizeFileName("...file...");

        Assert.False(result.StartsWith('.'));
        Assert.False(result.EndsWith('.'));
    }

    [Fact]
    public void SanitizeFileName_WithOnlyInvalidChars_ShouldReturnFile()
    {
        var result = SecurityHelper.SanitizeFileName("..../////");

        Assert.Equal("file", result);
    }

    #endregion

    #region IsSafeFilePath Tests

    [Fact]
    public void IsSafeFilePath_WithRelativePath_ShouldReturnTrue()
    {
        var result = SecurityHelper.IsSafeFilePath("documents/file.docx");

        Assert.True(result);
    }

    [Fact]
    public void IsSafeFilePath_WithNullOrEmpty_ShouldReturnFalse()
    {
        Assert.False(SecurityHelper.IsSafeFilePath(null!));
        Assert.False(SecurityHelper.IsSafeFilePath(""));
        Assert.False(SecurityHelper.IsSafeFilePath("   "));
    }

    [Fact]
    public void IsSafeFilePath_WithPathTraversal_ShouldReturnFalse()
    {
        Assert.False(SecurityHelper.IsSafeFilePath("../etc/passwd"));
        Assert.False(SecurityHelper.IsSafeFilePath("..\\windows\\system32"));
    }

    [Fact]
    public void IsSafeFilePath_WithTilde_ShouldReturnFalse()
    {
        Assert.False(SecurityHelper.IsSafeFilePath("~/documents/file.txt"));
    }

    [Fact]
    public void IsSafeFilePath_WithDoubleSlash_ShouldReturnFalse()
    {
        Assert.False(SecurityHelper.IsSafeFilePath("path//to//file"));
        Assert.False(SecurityHelper.IsSafeFilePath("path\\\\to\\\\file"));
    }

    [Fact]
    public void IsSafeFilePath_WithAbsolutePath_DefaultShouldReturnFalse()
    {
        Assert.False(SecurityHelper.IsSafeFilePath("C:\\Windows\\System32\\file.txt"));
        Assert.False(SecurityHelper.IsSafeFilePath("/etc/passwd"));
    }

    [Fact]
    public void IsSafeFilePath_WithAbsolutePath_WhenAllowed_ShouldReturnTrue()
    {
        var result = SecurityHelper.IsSafeFilePath("C:\\Users\\Documents\\file.txt", true);

        Assert.True(result);
    }

    [Fact]
    public void IsSafeFilePath_WithLongPath_ShouldReturnFalse()
    {
        var longPath = new string('a', 300);

        Assert.False(SecurityHelper.IsSafeFilePath(longPath));
    }

    #endregion

    #region ValidateFilePath Tests

    [Fact]
    public void ValidateFilePath_WithValidPath_ShouldReturnPath()
    {
        var result = SecurityHelper.ValidateFilePath("documents/file.docx");

        Assert.Equal("documents/file.docx", result);
    }

    [Fact]
    public void ValidateFilePath_WithNullOrEmpty_ShouldThrow()
    {
        var ex1 = Assert.Throws<ArgumentException>(() => SecurityHelper.ValidateFilePath(null!));
        Assert.Contains("cannot be null or empty", ex1.Message);

        var ex2 = Assert.Throws<ArgumentException>(() => SecurityHelper.ValidateFilePath(""));
        Assert.Contains("cannot be null or empty", ex2.Message);
    }

    [Fact]
    public void ValidateFilePath_WithPathTraversal_ShouldThrow()
    {
        var ex = Assert.Throws<ArgumentException>(() => SecurityHelper.ValidateFilePath("../etc/passwd"));

        Assert.Contains("invalid characters or path traversal", ex.Message);
    }

    [Fact]
    public void ValidateFilePath_WithCustomParamName_ShouldIncludeInError()
    {
        var ex = Assert.Throws<ArgumentException>(() =>
            SecurityHelper.ValidateFilePath("", "filePath"));

        Assert.Contains("filePath", ex.Message);
    }

    [Fact]
    public void ValidateFilePath_WithAbsolutePath_WhenAllowed_ShouldReturnPath()
    {
        var result = SecurityHelper.ValidateFilePath("C:\\Users\\file.txt", "path", true);

        Assert.Equal("C:\\Users\\file.txt", result);
    }

    #endregion

    #region SanitizeFileNamePattern Tests

    [Fact]
    public void SanitizeFileNamePattern_WithValidPattern_ShouldReturnSame()
    {
        var result = SecurityHelper.SanitizeFileNamePattern("file_{index}.docx");

        Assert.Equal("file_{index}.docx", result);
    }

    [Fact]
    public void SanitizeFileNamePattern_WithNullOrEmpty_ShouldReturnDefault()
    {
        Assert.Equal("file_{index}", SecurityHelper.SanitizeFileNamePattern(null!));
        Assert.Equal("file_{index}", SecurityHelper.SanitizeFileNamePattern(""));
        Assert.Equal("file_{index}", SecurityHelper.SanitizeFileNamePattern("   "));
    }

    [Fact]
    public void SanitizeFileNamePattern_WithPathTraversal_ShouldRemove()
    {
        var result = SecurityHelper.SanitizeFileNamePattern("..\\..\\file_{index}");

        Assert.DoesNotContain("..", result);
    }

    [Fact]
    public void SanitizeFileNamePattern_WithSlashes_ShouldReplace()
    {
        var result = SecurityHelper.SanitizeFileNamePattern("path/to/file_{index}");

        Assert.DoesNotContain("/", result);
        Assert.DoesNotContain("\\", result);
    }

    [Fact]
    public void SanitizeFileNamePattern_WithLongPattern_ShouldTruncate()
    {
        var longPattern = new string('a', 300) + "_{index}";

        var result = SecurityHelper.SanitizeFileNamePattern(longPattern);

        Assert.True(result.Length <= 255);
    }

    #endregion

    #region ValidateArraySize Tests

    [Fact]
    public void ValidateArraySize_WithinLimit_ShouldNotThrow()
    {
        var array = Enumerable.Range(1, 100);

        var ex = Record.Exception(() => SecurityHelper.ValidateArraySize(array));

        Assert.Null(ex);
    }

    [Fact]
    public void ValidateArraySize_ExceedsLimit_ShouldThrow()
    {
        var array = Enumerable.Range(1, 1001);

        var ex = Assert.Throws<ArgumentException>(() => SecurityHelper.ValidateArraySize(array));

        Assert.Contains("exceeds maximum allowed size", ex.Message);
    }

    [Fact]
    public void ValidateArraySize_WithCustomLimit_ShouldRespectLimit()
    {
        var array = Enumerable.Range(1, 10);

        var ex = Assert.Throws<ArgumentException>(() =>
            SecurityHelper.ValidateArraySize(array, "items", 5));

        Assert.Contains("items", ex.Message);
        Assert.Contains("5", ex.Message);
    }

    [Fact]
    public void ValidateArraySize_EmptyArray_ShouldNotThrow()
    {
        var array = Array.Empty<int>();

        var ex = Record.Exception(() => SecurityHelper.ValidateArraySize(array));

        Assert.Null(ex);
    }

    #endregion

    #region ValidateStringLength Tests

    [Fact]
    public void ValidateStringLength_WithinLimit_ShouldNotThrow()
    {
        var value = new string('a', 100);

        var ex = Record.Exception(() => SecurityHelper.ValidateStringLength(value));

        Assert.Null(ex);
    }

    [Fact]
    public void ValidateStringLength_ExceedsLimit_ShouldThrow()
    {
        var value = new string('a', 10001);

        var ex = Assert.Throws<ArgumentException>(() => SecurityHelper.ValidateStringLength(value));

        Assert.Contains("exceeds maximum allowed length", ex.Message);
    }

    [Fact]
    public void ValidateStringLength_WithCustomLimit_ShouldRespectLimit()
    {
        var value = new string('a', 100);

        var ex = Assert.Throws<ArgumentException>(() =>
            SecurityHelper.ValidateStringLength(value, "content", 50));

        Assert.Contains("content", ex.Message);
        Assert.Contains("50", ex.Message);
    }

    [Fact]
    public void ValidateStringLength_EmptyString_ShouldNotThrow()
    {
        var ex = Record.Exception(() => SecurityHelper.ValidateStringLength(""));

        Assert.Null(ex);
    }

    #endregion
}