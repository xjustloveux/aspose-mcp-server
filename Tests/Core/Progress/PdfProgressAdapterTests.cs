using AsposeMcpServer.Core.Progress;
using ModelContextProtocol;

namespace AsposeMcpServer.Tests.Core.Progress;

/// <summary>
///     Unit tests for <see cref="PdfProgressAdapter" /> class.
/// </summary>
public class PdfProgressAdapterTests
{
    #region Constructor Tests

    [Fact]
    public void Constructor_WithNullProgress_DoesNotThrow()
    {
        var adapter = new PdfProgressAdapter(null);

        var exception = Record.Exception(() =>
        {
            adapter.ReportProgress(0, 10, "test");
            adapter.ReportPercentage(50, "test");
        });

        Assert.Null(exception);
    }

    #endregion

    #region ReportProgress Tests

    [Fact]
    public void ReportProgress_WithValidItems_ReportsCorrectPercentage()
    {
        ProgressNotificationValue? reported = null;
        var progress = new Progress<ProgressNotificationValue>(v => reported = v);
        var adapter = new PdfProgressAdapter(progress);

        adapter.ReportProgress(4, 10, "Processing");

        Thread.Sleep(50);

        Assert.NotNull(reported);
        Assert.Equal(50, reported.Progress);
        Assert.Equal(100, reported.Total);
        Assert.Equal("Processing", reported.Message);
    }

    [Fact]
    public void ReportProgress_WithZeroTotalItems_DoesNotReport()
    {
        var reportCount = 0;
        var progress = new Progress<ProgressNotificationValue>(_ => reportCount++);
        var adapter = new PdfProgressAdapter(progress);

        adapter.ReportProgress(0, 0);

        Thread.Sleep(50);

        Assert.Equal(0, reportCount);
    }

    [Fact]
    public void ReportProgress_FirstItem_ReportsCorrectPercentage()
    {
        ProgressNotificationValue? reported = null;
        var progress = new Progress<ProgressNotificationValue>(v => reported = v);
        var adapter = new PdfProgressAdapter(progress);

        adapter.ReportProgress(0, 4);

        Thread.Sleep(50);

        Assert.NotNull(reported);
        Assert.Equal(25, reported.Progress);
    }

    [Fact]
    public void ReportProgress_LastItem_Reports100Percent()
    {
        ProgressNotificationValue? reported = null;
        var progress = new Progress<ProgressNotificationValue>(v => reported = v);
        var adapter = new PdfProgressAdapter(progress);

        adapter.ReportProgress(9, 10);

        Thread.Sleep(50);

        Assert.NotNull(reported);
        Assert.Equal(100, reported.Progress);
    }

    #endregion

    #region ReportPercentage Tests

    [Fact]
    public void ReportPercentage_ReportsExactPercentage()
    {
        ProgressNotificationValue? reported = null;
        var progress = new Progress<ProgressNotificationValue>(v => reported = v);
        var adapter = new PdfProgressAdapter(progress);

        adapter.ReportPercentage(75, "Almost done");

        Thread.Sleep(50);

        Assert.NotNull(reported);
        Assert.Equal(75, reported.Progress);
        Assert.Equal(100, reported.Total);
        Assert.Equal("Almost done", reported.Message);
    }

    [Fact]
    public void ReportPercentage_WithNullMessage_ReportsWithoutMessage()
    {
        ProgressNotificationValue? reported = null;
        var progress = new Progress<ProgressNotificationValue>(v => reported = v);
        var adapter = new PdfProgressAdapter(progress);

        adapter.ReportPercentage(50);

        Thread.Sleep(50);

        Assert.NotNull(reported);
        Assert.Equal(50, reported.Progress);
        Assert.Null(reported.Message);
    }

    #endregion
}
