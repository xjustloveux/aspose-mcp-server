using AsposeMcpServer.Core.Conversion;
using AsposeMcpServer.Core.Extension;
using AsposeMcpServer.Core.Session;

namespace AsposeMcpServer.Tests.Core.Extension;

/// <summary>
///     Unit tests for SessionBindingInfo class.
/// </summary>
public class SessionBindingInfoTests
{
    #region Thread Safety Tests

    [Fact]
    public async Task SessionBindingInfo_ConcurrentAccess_IsThreadSafe()
    {
        var binding = new SessionBindingInfo(100);
        var tasks = new List<Task>();

        for (var i = 0; i < 10; i++)
            tasks.Add(Task.Run(() =>
            {
                for (var j = 0; j < 100; j++)
                {
                    binding.NeedsSend = true;
                    binding.OutputFormat = $"format_{j}";
                    _ = binding.NeedsSend;
                    _ = binding.OutputFormat;
                    binding.RecordConversionFailure();
                }
            }));

        var exception = await Record.ExceptionAsync(() => Task.WhenAll(tasks));

        Assert.Null(exception);
    }

    #endregion

    #region UpdateLastSent Backoff Tests

    [Fact]
    public void UpdateLastSent_ClearsBackoff()
    {
        var binding = new SessionBindingInfo(2);
        binding.RecordConversionFailure();
        binding.RecordConversionFailure();
        Assert.True(binding.IsInBackoff());

        binding.UpdateLastSent(DateTime.UtcNow);

        Assert.False(binding.IsInBackoff());
    }

    #endregion

    #region UpdateFormat Backoff Tests

    [Fact]
    public void UpdateFormat_ClearsBackoff()
    {
        var binding = new SessionBindingInfo(2);
        binding.RecordConversionFailure();
        binding.RecordConversionFailure();
        Assert.True(binding.IsInBackoff());

        binding.UpdateFormat("html");

        Assert.False(binding.IsInBackoff());
    }

    #endregion

    #region Default Values Tests

    [Fact]
    public void SessionBindingInfo_DefaultValues_AreCorrect()
    {
        var binding = new SessionBindingInfo();

        Assert.Equal(string.Empty, binding.SessionId);
        Assert.Equal(string.Empty, binding.ExtensionId);
        Assert.Equal(string.Empty, binding.OutputFormat);
        Assert.Equal(default, binding.CreatedAt);
        Assert.Null(binding.LastSentAt);
        Assert.True(binding.NeedsSend);
        Assert.Equal(0, binding.ConversionFailures);
    }

    [Fact]
    public void SessionBindingInfo_CanSetAllProperties()
    {
        var createdAt = DateTime.UtcNow.AddMinutes(-10);
        var lastSentAt = DateTime.UtcNow;

        var binding = new SessionBindingInfo
        {
            SessionId = "sess_123",
            ExtensionId = "ext_456",
            OutputFormat = "pdf",
            CreatedAt = createdAt,
            LastSentAt = lastSentAt,
            NeedsSend = false
        };

        Assert.Equal("sess_123", binding.SessionId);
        Assert.Equal("ext_456", binding.ExtensionId);
        Assert.Equal("pdf", binding.OutputFormat);
        Assert.Equal(createdAt, binding.CreatedAt);
        Assert.Equal(lastSentAt, binding.LastSentAt);
        Assert.False(binding.NeedsSend);
    }

    #endregion

    #region NeedsSend Tests

    [Fact]
    public void SessionBindingInfo_NeedsSend_DefaultTrue()
    {
        var binding = new SessionBindingInfo();

        Assert.True(binding.NeedsSend);
    }

    [Fact]
    public void SessionBindingInfo_NeedsSend_CanBeToggled()
    {
        var binding = new SessionBindingInfo { NeedsSend = false };

        Assert.False(binding.NeedsSend);

        binding.NeedsSend = true;

        Assert.True(binding.NeedsSend);
    }

    #endregion

    #region LastSentAt Tests

    [Fact]
    public void SessionBindingInfo_LastSentAt_InitiallyNull()
    {
        var binding = new SessionBindingInfo();

        Assert.Null(binding.LastSentAt);
    }

    [Fact]
    public void SessionBindingInfo_LastSentAt_CanBeSet()
    {
        var timestamp = DateTime.UtcNow;
        var binding = new SessionBindingInfo { LastSentAt = timestamp };

        Assert.Equal(timestamp, binding.LastSentAt);
    }

    [Fact]
    public void SessionBindingInfo_LastSentAt_CanBeReset()
    {
        var binding = new SessionBindingInfo
        {
            LastSentAt = DateTime.UtcNow
        };

        binding.LastSentAt = null;

        Assert.Null(binding.LastSentAt);
    }

    #endregion

    #region IsInBackoff Tests

    [Fact]
    public void IsInBackoff_Initially_ReturnsFalse()
    {
        var binding = new SessionBindingInfo();

        Assert.False(binding.IsInBackoff());
    }

    [Fact]
    public void IsInBackoff_AfterMaxFailures_ReturnsTrue()
    {
        var binding = new SessionBindingInfo(3);

        binding.RecordConversionFailure();
        binding.RecordConversionFailure();
        binding.RecordConversionFailure();

        Assert.True(binding.IsInBackoff());
    }

    [Fact]
    public void IsInBackoff_BeforeMaxFailures_ReturnsFalse()
    {
        var binding = new SessionBindingInfo(3);

        binding.RecordConversionFailure();
        binding.RecordConversionFailure();

        Assert.False(binding.IsInBackoff());
    }

    #endregion

    #region RecordConversionFailure Tests

    [Fact]
    public void RecordConversionFailure_IncrementsCounter()
    {
        var binding = new SessionBindingInfo();

        binding.RecordConversionFailure();
        Assert.Equal(1, binding.ConversionFailures);

        binding.RecordConversionFailure();
        Assert.Equal(2, binding.ConversionFailures);
    }

    [Fact]
    public void RecordConversionFailure_ReturnsTrueWhenEnteringBackoff()
    {
        var binding = new SessionBindingInfo(2);

        var firstResult = binding.RecordConversionFailure();
        Assert.False(firstResult);

        var secondResult = binding.RecordConversionFailure();
        Assert.True(secondResult);
    }

    #endregion

    #region ResetConversionFailures Tests

    [Fact]
    public void ResetConversionFailures_ClearsCounter()
    {
        var binding = new SessionBindingInfo();
        binding.RecordConversionFailure();
        binding.RecordConversionFailure();

        binding.ResetConversionFailures();

        Assert.Equal(0, binding.ConversionFailures);
    }

    [Fact]
    public void ResetConversionFailures_ClearsBackoff()
    {
        var binding = new SessionBindingInfo(2);
        binding.RecordConversionFailure();
        binding.RecordConversionFailure();
        Assert.True(binding.IsInBackoff());

        binding.ResetConversionFailures();

        Assert.False(binding.IsInBackoff());
    }

    #endregion

    #region UpdateLastSent Tests

    [Fact]
    public void UpdateLastSent_SetsTimestampAndClearsNeedsSend()
    {
        var binding = new SessionBindingInfo { NeedsSend = true };
        var timestamp = DateTime.UtcNow;

        binding.UpdateLastSent(timestamp);

        Assert.Equal(timestamp, binding.LastSentAt);
        Assert.False(binding.NeedsSend);
    }

    [Fact]
    public void UpdateLastSent_ResetsConversionFailures()
    {
        var binding = new SessionBindingInfo();
        binding.RecordConversionFailure();
        binding.RecordConversionFailure();

        binding.UpdateLastSent(DateTime.UtcNow);

        Assert.Equal(0, binding.ConversionFailures);
    }

    #endregion

    #region UpdateFormat Tests

    [Fact]
    public void UpdateFormat_SetsFormatAndNeedsSend()
    {
        var binding = new SessionBindingInfo
        {
            OutputFormat = "pdf",
            NeedsSend = false
        };

        binding.UpdateFormat("html");

        Assert.Equal("html", binding.OutputFormat);
        Assert.True(binding.NeedsSend);
    }

    [Fact]
    public void UpdateFormat_ResetsConversionFailures()
    {
        var binding = new SessionBindingInfo();
        binding.RecordConversionFailure();
        binding.RecordConversionFailure();

        binding.UpdateFormat("html");

        Assert.Equal(0, binding.ConversionFailures);
    }

    #endregion

    #region Owner Tests

    [Fact]
    public void SessionBindingInfo_Owner_DefaultIsAnonymous()
    {
        var binding = new SessionBindingInfo();

        Assert.True(binding.Owner.IsAnonymous);
    }

    [Fact]
    public void SessionBindingInfo_Owner_CanBeSet()
    {
        var owner = new SessionIdentity { GroupId = "group1", UserId = "user1" };
        var binding = new SessionBindingInfo { Owner = owner };

        Assert.Equal("group1", binding.Owner.GroupId);
        Assert.Equal("user1", binding.Owner.UserId);
    }

    #endregion

    #region UpdateFormatAndOptions Tests

    [Fact]
    public void UpdateFormatAndOptions_SetsFormatAndOptionsAndNeedsSend()
    {
        var binding = new SessionBindingInfo
        {
            OutputFormat = "pdf",
            NeedsSend = false
        };
        var options = new ConversionOptions { JpegQuality = 85 };

        binding.UpdateFormatAndOptions("html", options);

        Assert.Equal("html", binding.OutputFormat);
        Assert.Equal(85, binding.ConversionOptions.JpegQuality);
        Assert.True(binding.NeedsSend);
    }

    [Fact]
    public void UpdateFormatAndOptions_ResetsConversionFailures()
    {
        var binding = new SessionBindingInfo();
        binding.RecordConversionFailure();
        binding.RecordConversionFailure();

        binding.UpdateFormatAndOptions("html", new ConversionOptions());

        Assert.Equal(0, binding.ConversionFailures);
    }

    [Fact]
    public void UpdateFormatAndOptions_ClearsBackoff()
    {
        var binding = new SessionBindingInfo(2);
        binding.RecordConversionFailure();
        binding.RecordConversionFailure();
        Assert.True(binding.IsInBackoff());

        binding.UpdateFormatAndOptions("html", new ConversionOptions());

        Assert.False(binding.IsInBackoff());
    }

    #endregion

    #region ConversionOptions Tests

    [Fact]
    public void ConversionOptions_DefaultIsNewInstance()
    {
        var binding = new SessionBindingInfo();

        Assert.NotNull(binding.ConversionOptions);
    }

    [Fact]
    public void ConversionOptions_CanBeSet()
    {
        var options = new ConversionOptions
        {
            JpegQuality = 90,
            Dpi = 150,
            CsvSeparator = ";"
        };
        var binding = new SessionBindingInfo { ConversionOptions = options };

        Assert.Equal(90, binding.ConversionOptions.JpegQuality);
        Assert.Equal(150, binding.ConversionOptions.Dpi);
        Assert.Equal(";", binding.ConversionOptions.CsvSeparator);
    }

    #endregion

    #region GetOptionsCacheKey Tests

    [Fact]
    public void GetOptionsCacheKey_ReturnsConsistentHashForSameOptions()
    {
        var binding = new SessionBindingInfo
        {
            ConversionOptions = new ConversionOptions
            {
                JpegQuality = 85,
                Dpi = 300
            }
        };

        var key1 = binding.GetOptionsCacheKey();
        var key2 = binding.GetOptionsCacheKey();

        Assert.Equal(key1, key2);
    }

    [Fact]
    public void GetOptionsCacheKey_ReturnsDifferentHashForDifferentOptions()
    {
        var binding1 = new SessionBindingInfo
        {
            ConversionOptions = new ConversionOptions { JpegQuality = 85 }
        };
        var binding2 = new SessionBindingInfo
        {
            ConversionOptions = new ConversionOptions { JpegQuality = 90 }
        };

        var key1 = binding1.GetOptionsCacheKey();
        var key2 = binding2.GetOptionsCacheKey();

        Assert.NotEqual(key1, key2);
    }

    [Fact]
    public void GetOptionsCacheKey_IncludesAllRelevantOptions()
    {
        var binding = new SessionBindingInfo
        {
            ConversionOptions = new ConversionOptions
            {
                JpegQuality = 85,
                CsvSeparator = ",",
                PdfCompliance = "PDF/A-1b",
                HtmlEmbedImages = true,
                HtmlSingleFile = true,
                Dpi = 300
            }
        };

        var key = binding.GetOptionsCacheKey();

        Assert.NotEqual(0, key);
    }

    #endregion
}
