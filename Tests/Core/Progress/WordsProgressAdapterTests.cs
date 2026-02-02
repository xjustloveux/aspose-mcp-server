using Aspose.Words;
using Aspose.Words.Saving;
using AsposeMcpServer.Core.Progress;
using ModelContextProtocol;

namespace AsposeMcpServer.Tests.Core.Progress;

/// <summary>
///     Unit tests for <see cref="WordsProgressAdapter" /> class.
/// </summary>
public class WordsProgressAdapterTests : IDisposable
{
    private readonly string _testDir;

    public WordsProgressAdapterTests()
    {
        _testDir = Path.Combine(Path.GetTempPath(), $"WordsProgressTests_{Guid.NewGuid():N}");
        Directory.CreateDirectory(_testDir);
    }

    public void Dispose()
    {
        if (Directory.Exists(_testDir))
            Directory.Delete(_testDir, true);
    }

    #region Constructor Tests

    [Fact]
    public void Constructor_WithNullProgress_DoesNotThrow()
    {
        var adapter = new WordsProgressAdapter(null);
        Assert.NotNull(adapter);
    }

    [Fact]
    public void Constructor_WithProgress_DoesNotThrow()
    {
        var progress = new Progress<ProgressNotificationValue>();
        var adapter = new WordsProgressAdapter(progress);
        Assert.NotNull(adapter);
    }

    #endregion

    #region Notify Tests (via actual document save)

    [Fact]
    public void Notify_DuringSave_ReportsProgress()
    {
        var reportedValues = new List<ProgressNotificationValue>();
        var progress = new Progress<ProgressNotificationValue>(v => reportedValues.Add(v));
        var adapter = new WordsProgressAdapter(progress);

        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        for (var i = 0; i < 50; i++)
            builder.Writeln($"Test paragraph {i} with some content to increase document size.");

        var outputPath = Path.Combine(_testDir, "test_progress.docx");
        var saveOptions = new OoxmlSaveOptions
        {
            ProgressCallback = adapter
        };

        doc.Save(outputPath, saveOptions);

        Thread.Sleep(100);

        Assert.NotEmpty(reportedValues);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void Notify_WithNullProgress_DuringSave_DoesNotThrow()
    {
        var adapter = new WordsProgressAdapter(null);

        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        for (var i = 0; i < 20; i++)
            builder.Writeln($"Test paragraph {i}.");

        var outputPath = Path.Combine(_testDir, "test_null_progress.docx");
        var saveOptions = new OoxmlSaveOptions
        {
            ProgressCallback = adapter
        };

        doc.Save(outputPath, saveOptions);

        Assert.True(File.Exists(outputPath));
    }

    #endregion
}
