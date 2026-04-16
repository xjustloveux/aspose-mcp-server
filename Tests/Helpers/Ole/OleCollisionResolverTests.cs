using AsposeMcpServer.Helpers.Ole;

namespace AsposeMcpServer.Tests.Helpers.Ole;

/// <summary>
///     Unit tests for <see cref="OleCollisionResolver" />.
/// </summary>
public class OleCollisionResolverTests : IDisposable
{
    /// <summary>
    ///     Temp directory used for filesystem checks within each test; cleaned up in
    ///     <see cref="Dispose" />.
    /// </summary>
    private readonly string _tempDir;

    /// <summary>
    ///     Initializes a fresh temp directory so collision tests against the filesystem
    ///     are isolated from each other and from co-running tests.
    /// </summary>
    public OleCollisionResolverTests()
    {
        _tempDir = Path.Combine(Path.GetTempPath(), "OleColl_" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(_tempDir);
    }

    /// <summary>Disposes the temp directory.</summary>
    public void Dispose()
    {
        try
        {
            if (Directory.Exists(_tempDir)) Directory.Delete(_tempDir, true);
        }
        catch (IOException)
        {
            /* ignored — best-effort cleanup */
        }

        GC.SuppressFinalize(this);
    }

    [Fact]
    public void Reserve_FirstCall_ReturnsOriginalName()
    {
        var resolver = new OleCollisionResolver();
        var path = resolver.Reserve(_tempDir, "file.xlsx");

        Assert.Equal(Path.Combine(_tempDir, "file.xlsx"), path);
    }

    [Fact]
    public void Reserve_SecondCallWithSameInMemory_Appends2()
    {
        var resolver = new OleCollisionResolver();
        var first = resolver.Reserve(_tempDir, "file.xlsx");
        var second = resolver.Reserve(_tempDir, "file.xlsx");

        Assert.Equal(Path.Combine(_tempDir, "file.xlsx"), first);
        Assert.Equal(Path.Combine(_tempDir, "file (2).xlsx"), second);
    }

    [Fact]
    public void Reserve_ThirdCallWithSameInMemory_Appends3()
    {
        var resolver = new OleCollisionResolver();
        resolver.Reserve(_tempDir, "file.xlsx");
        resolver.Reserve(_tempDir, "file.xlsx");
        var third = resolver.Reserve(_tempDir, "file.xlsx");

        Assert.Equal(Path.Combine(_tempDir, "file (3).xlsx"), third);
    }

    [Fact]
    public void Reserve_ExistingOnDisk_SkipsToNextFree()
    {
        File.WriteAllText(Path.Combine(_tempDir, "a.bin"), "x");
        var resolver = new OleCollisionResolver();

        var path = resolver.Reserve(_tempDir, "a.bin");

        Assert.Equal(Path.Combine(_tempDir, "a (2).bin"), path);
    }

    [Fact]
    public void Reserve_NameWithoutExtension_StillAppendsBeforeExt()
    {
        var resolver = new OleCollisionResolver();
        var first = resolver.Reserve(_tempDir, "noext");
        var second = resolver.Reserve(_tempDir, "noext");

        Assert.Equal(Path.Combine(_tempDir, "noext"), first);
        Assert.Equal(Path.Combine(_tempDir, "noext (2)"), second);
    }

    [Fact]
    public void Reserve_EmptyPreferredName_Throws()
    {
        var resolver = new OleCollisionResolver();
        Assert.Throws<ArgumentException>(() => resolver.Reserve(_tempDir, string.Empty));
    }

    [Fact]
    public void Reserve_EmptyOutputDirectory_Throws()
    {
        var resolver = new OleCollisionResolver();
        Assert.Throws<ArgumentException>(() => resolver.Reserve("  ", "file.xlsx"));
    }
}
