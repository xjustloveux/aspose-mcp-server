using AsposeMcpServer.Core.Extension;

namespace AsposeMcpServer.Tests.Core.Extension;

/// <summary>
///     Unit tests for ConstraintResult record.
/// </summary>
public class ConstraintResultTests
{
    [Fact]
    public void Constructor_SetsValueAndWarning()
    {
        var result = new ConstraintResult(100, "test warning");

        Assert.Equal(100, result.Value);
        Assert.Equal("test warning", result.Warning);
    }

    [Fact]
    public void HasWarning_WithWarning_ReturnsTrue()
    {
        var result = new ConstraintResult(100, "test warning");

        Assert.True(result.HasWarning);
    }

    [Fact]
    public void HasWarning_WithNullWarning_ReturnsFalse()
    {
        var result = new ConstraintResult(100, null);

        Assert.False(result.HasWarning);
    }

    [Fact]
    public void Equality_SameValues_AreEqual()
    {
        var result1 = new ConstraintResult(100, "warning");
        var result2 = new ConstraintResult(100, "warning");

        Assert.Equal(result1, result2);
    }

    [Fact]
    public void Equality_DifferentValues_AreNotEqual()
    {
        var result1 = new ConstraintResult(100, "warning");
        var result2 = new ConstraintResult(200, "warning");

        Assert.NotEqual(result1, result2);
    }

    [Fact]
    public void Deconstruction_Works()
    {
        var result = new ConstraintResult(100, "test warning");

        var (value, warning) = result;

        Assert.Equal(100, value);
        Assert.Equal("test warning", warning);
    }
}
