using AsposeMcpServer.Helpers;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Helpers;

/// <summary>
///     R18-SEC02 and R18-SEC03: the two mistakes the journal and the queue each made a copy of.
///     <para>
///         Both staged through a predictable <c>&lt;name&gt;.writing</c> and wrote it with a
///         path-based call, so a link planted at that name was followed. And both measured a size
///         through one open and read the content through another, so the file the cap was measured
///         on did not have to be the file that was read.
///     </para>
/// </summary>
public class SecureFileTests : TestBase
{
    [Fact]
    public void AStagingNameIsNotPredictable()
    {
        // Two writes to the same destination must not stage through the same name, or the name is
        // one an attacker can prepare.
        var destination = CreateTestFilePath("predictable.json");
        var seen = new List<string>();

        for (var i = 0; i < 4; i++)
            SecureFile.ReplaceAtomically(destination, $"content {i}",
                (stream, content) =>
                {
                    seen.Add(((FileStream)stream).Name);
                    using var writer = new StreamWriter(stream, leaveOpen: true);
                    writer.Write(content);
                });

        Assert.Equal(4, seen.Distinct(StringComparer.Ordinal).Count());
        Assert.All(seen, path => Assert.NotEqual(destination + ".writing", path));
        Assert.Equal("content 3", File.ReadAllText(destination));
    }

    [Fact]
    public void NoStagingFileIsLeftBehind()
    {
        var destination = CreateTestFilePath("tidy.json");

        SecureFile.ReplaceAtomically(destination, "content");

        Assert.Equal("tidy.json",
            Path.GetFileName(Assert.Single(Directory.GetFiles(TestDir, "tidy.json*"))));
    }

    [Fact]
    public void AFailedWriteLeavesNeitherAStagingFileNorAChangedDestination()
    {
        var destination = CreateTestFilePath("failed.json");
        File.WriteAllText(destination, "the previous content");

        Assert.Throws<IOException>(() => SecureFile.ReplaceAtomically(destination, "new",
            (_, _) => throw new IOException("the volume is full")));

        Assert.Equal("the previous content", File.ReadAllText(destination));
        Assert.Single(Directory.GetFiles(TestDir, "failed.json*"));
    }

    [Fact]
    public void AFileWithinTheBound_IsRead()
    {
        var path = CreateTestFilePath("within.json");
        File.WriteAllText(path, "12345");

        Assert.True(SecureFile.TryReadBounded(path, 8, out var content, out var length));
        Assert.Equal("12345", content);
        Assert.Equal(5, length);
    }

    [Fact]
    public void AFileExactlyAtTheBound_IsRead()
    {
        // The boundary the "one byte past" read exists to get right.
        var path = CreateTestFilePath("exact.json");
        File.WriteAllText(path, "12345678");

        Assert.True(SecureFile.TryReadBounded(path, 8, out var content, out _));
        Assert.Equal("12345678", content);
    }

    [Fact]
    public void AFileOverTheBound_IsRefusedWithoutBeingReadWhole()
    {
        // One byte past the bound is enough to know it is over, and is all that is read of it.
        var path = CreateTestFilePath("over.json");
        File.WriteAllText(path, new string('x', 4096));

        Assert.False(SecureFile.TryReadBounded(path, 8, out var content, out var length));
        Assert.Equal(string.Empty, content);
        Assert.Equal(9, length);
    }

    [Fact]
    public void TheWriterSeam_ShouldReceiveTheHandleTheStagingFileWasCreatedWith()
    {
        // R20-REC05, the original R19-REC02. The seam used to be given a *path*: the staging file
        // was created, its handle closed, and the callback reopened the name — a second open that
        // a link planted in between would follow. Now it is given the stream, and the property
        // pinned here is that the stream *is* the file that was created: same name, still open,
        // writable, and what is written through it is what ends up at the destination.
        var destination = CreateTestFilePath("handle_bound.json");
        FileStream? received = null;

        SecureFile.ReplaceAtomically(destination, "through the handle", (stream, content) =>
        {
            received = Assert.IsType<FileStream>(stream);
            Assert.True(received.CanWrite);
            Assert.StartsWith(destination + ".writing-", received.Name, StringComparison.Ordinal);
            using var writer = new StreamWriter(stream, leaveOpen: true);
            writer.Write(content);
        });

        Assert.NotNull(received);
        Assert.Equal("through the handle", File.ReadAllText(destination));
        Assert.False(File.Exists(received!.Name), "the staging file outlived the replace");
    }
}
