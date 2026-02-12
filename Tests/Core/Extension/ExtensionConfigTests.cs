using AsposeMcpServer.Core.Extension;
using AsposeMcpServer.Core.Session;

namespace AsposeMcpServer.Tests.Core.Extension;

/// <summary>
///     Unit tests for ExtensionConfig class.
/// </summary>
public class ExtensionConfigTests
{
    #region Additional Default Properties Tests

    [Fact]
    public void DefaultValues_AdditionalProperties_AreCorrect()
    {
        var config = new ExtensionConfig();

        Assert.Equal(60, config.ErrorRecoveryCooldownSeconds);
        Assert.Equal(5, config.GracefulShutdownTimeoutSeconds);
        Assert.Equal(30, config.HandshakeTimeoutSeconds);
        Assert.Equal(5000, config.StdinWriteTimeoutMs);
        Assert.Equal(10, config.MaxConsecutiveSendFailures);
        Assert.Equal(60, config.RetryLoopTimeoutSeconds);
        Assert.Equal(5, config.ConversionCacheTtlSeconds);
        Assert.Equal(100, config.MaxConversionCacheSize);
        Assert.Equal(5, config.MaxConversionFailures);
        Assert.Equal(60, config.FailureBackoffSeconds);
    }

    #endregion

    #region Default Values Tests

    [Fact]
    public void DefaultValues_AreCorrect()
    {
        var config = new ExtensionConfig();

        Assert.False(config.Enabled);
        Assert.Null(config.ConfigPath);
        Assert.Equal(30, config.SnapshotTtlSeconds.Default);
        Assert.Equal(30, config.IdleTimeoutMinutes.Default);
        Assert.Equal(30, config.HealthCheckIntervalSeconds);
        Assert.Equal(3, config.MaxRestartAttempts);
        Assert.Equal(5, config.RestartCooldownSeconds);
        Assert.Equal(100, config.FrameIntervalMs.Default);
        Assert.Equal("stdin", config.DefaultTransportMode);
        Assert.Equal(100 * 1024 * 1024, config.MaxSnapshotSizeBytes);
        Assert.Equal(500 * 1024 * 1024, config.MinFreeDiskSpaceBytes);
    }

    [Fact]
    public void ConstrainedInt_DefaultFloorCeiling_AreCorrect()
    {
        var config = new ExtensionConfig();

        Assert.Equal(100, config.FrameIntervalMs.Default);
        Assert.Equal(10, config.FrameIntervalMs.Floor);
        Assert.Equal(5000, config.FrameIntervalMs.Ceiling);

        Assert.Equal(30, config.SnapshotTtlSeconds.Default);
        Assert.Equal(5, config.SnapshotTtlSeconds.Floor);
        Assert.Equal(300, config.SnapshotTtlSeconds.Ceiling);

        Assert.Equal(3, config.MaxMissedHeartbeats.Default);
        Assert.Equal(1, config.MaxMissedHeartbeats.Floor);
        Assert.Equal(20, config.MaxMissedHeartbeats.Ceiling);

        Assert.Equal(100, config.DebounceDelayMs);

        Assert.Equal(30, config.IdleTimeoutMinutes.Default);
        Assert.Equal(1, config.IdleTimeoutMinutes.Floor);
        Assert.Equal(1440, config.IdleTimeoutMinutes.Ceiling);
        Assert.Equal(0, config.IdleTimeoutMinutes.SpecialValue);
        Assert.True(config.IdleTimeoutMinutes.SpecialAllowed);
    }

    #endregion

    #region ConstrainedInt Apply Tests

    [Fact]
    public void ConstrainedInt_Apply_UsesDefaultWhenNull()
    {
        var config = new ExtensionConfig();

        var result = config.FrameIntervalMs.Apply(null);

        Assert.Equal(100, result);
    }

    [Fact]
    public void ConstrainedInt_Apply_ClampsToFloor()
    {
        var config = new ExtensionConfig();

        var result = config.FrameIntervalMs.Apply(5);

        Assert.Equal(10, result);
    }

    [Fact]
    public void ConstrainedInt_Apply_ClampsToCeiling()
    {
        var config = new ExtensionConfig();

        var result = config.FrameIntervalMs.Apply(10000);

        Assert.Equal(5000, result);
    }

    [Fact]
    public void ConstrainedInt_Apply_PassesThroughValidValue()
    {
        var config = new ExtensionConfig();

        var result = config.FrameIntervalMs.Apply(200);

        Assert.Equal(200, result);
    }

    [Fact]
    public void ConstrainedIntWithSpecial_Apply_AllowsSpecialValue()
    {
        var config = new ExtensionConfig
        {
            IdleTimeoutMinutes = { SpecialAllowed = true }
        };

        var result = config.IdleTimeoutMinutes.Apply(0);

        Assert.Equal(0, result);
    }

    [Fact]
    public void ConstrainedIntWithSpecial_Apply_RejectsSpecialValueWhenNotAllowed()
    {
        var config = new ExtensionConfig
        {
            IdleTimeoutMinutes = { SpecialAllowed = false }
        };

        var result = config.IdleTimeoutMinutes.Apply(0);

        Assert.Equal(30, result);
    }

    #endregion

    #region ApplyWithWarning Tests

    [Fact]
    public void ApplyWithWarning_NoConstraint_ReturnsNoWarning()
    {
        var config = new ExtensionConfig();

        var result = config.FrameIntervalMs.ApplyWithWarning(200, "FrameIntervalMs");

        Assert.Equal(200, result.Value);
        Assert.False(result.HasWarning);
        Assert.Null(result.Warning);
    }

    [Fact]
    public void ApplyWithWarning_BelowFloor_ReturnsWarning()
    {
        var config = new ExtensionConfig();

        var result = config.FrameIntervalMs.ApplyWithWarning(5, "FrameIntervalMs");

        Assert.Equal(10, result.Value);
        Assert.True(result.HasWarning);
        Assert.Contains("below minimum", result.Warning);
        Assert.Contains("FrameIntervalMs=5", result.Warning);
    }

    [Fact]
    public void ApplyWithWarning_AboveCeiling_ReturnsWarning()
    {
        var config = new ExtensionConfig();

        var result = config.FrameIntervalMs.ApplyWithWarning(10000, "FrameIntervalMs");

        Assert.Equal(5000, result.Value);
        Assert.True(result.HasWarning);
        Assert.Contains("above maximum", result.Warning);
        Assert.Contains("FrameIntervalMs=10000", result.Warning);
    }

    [Fact]
    public void ApplyWithWarning_NullValue_UsesDefault_NoWarning()
    {
        var config = new ExtensionConfig();

        var result = config.FrameIntervalMs.ApplyWithWarning(null, "FrameIntervalMs");

        Assert.Equal(100, result.Value);
        Assert.False(result.HasWarning);
    }

    [Fact]
    public void ApplyWithWarning_SpecialValueAllowed_NoWarning()
    {
        var config = new ExtensionConfig
        {
            IdleTimeoutMinutes = { SpecialAllowed = true }
        };

        var result = config.IdleTimeoutMinutes.ApplyWithWarning(0, "IdleTimeoutMinutes");

        Assert.Equal(0, result.Value);
        Assert.False(result.HasWarning);
    }

    [Fact]
    public void ApplyWithWarning_SpecialValueNotAllowed_ReturnsWarning()
    {
        var config = new ExtensionConfig
        {
            IdleTimeoutMinutes = { SpecialAllowed = false }
        };

        var result = config.IdleTimeoutMinutes.ApplyWithWarning(0, "IdleTimeoutMinutes");

        Assert.Equal(30, result.Value);
        Assert.True(result.HasWarning);
        Assert.Contains("special value", result.Warning);
        Assert.Contains("not allowed", result.Warning);
    }

    #endregion

    #region LoadFromArgs Tests

    [Fact]
    public void LoadFromArgs_NoExtensionArg_ReturnsDisabled()
    {
        var args = Array.Empty<string>();

        var config = ExtensionConfig.LoadFromArgs(args);

        Assert.False(config.Enabled);
    }

    [Fact]
    public void LoadFromArgs_WithExtensionEnabled_ReturnsEnabled()
    {
        var args = new[] { "--extension-enabled" };

        var config = ExtensionConfig.LoadFromArgs(args);

        Assert.True(config.Enabled);
    }

    [Fact]
    public void LoadFromArgs_WithExtensionDisabled_ReturnsDisabled()
    {
        var args = new[] { "--extension-disabled" };

        var config = ExtensionConfig.LoadFromArgs(args);

        Assert.False(config.Enabled);
    }

    [Fact]
    public void LoadFromArgs_WithExtensionConfig_SetsConfigPath()
    {
        var args = new[] { "--extension-config", "custom/path.json" };

        var config = ExtensionConfig.LoadFromArgs(args);

        Assert.Equal("custom/path.json", config.ConfigPath);
    }

    [Fact]
    public void LoadFromArgs_WithSnapshotTtl_SetsSnapshotTtlDefault()
    {
        var args = new[] { "--extension-snapshot-ttl", "60" };

        var config = ExtensionConfig.LoadFromArgs(args);

        Assert.Equal(60, config.SnapshotTtlSeconds.Default);
    }

    [Fact]
    public void LoadFromArgs_WithIdleTimeout_SetsIdleTimeoutDefault()
    {
        var args = new[] { "--extension-idle-timeout", "45" };

        var config = ExtensionConfig.LoadFromArgs(args);

        Assert.Equal(45, config.IdleTimeoutMinutes.Default);
    }

    [Fact]
    public void LoadFromArgs_WithHealthInterval_SetsHealthCheckIntervalSeconds()
    {
        var args = new[] { "--extension-health-interval", "15" };

        var config = ExtensionConfig.LoadFromArgs(args);

        Assert.Equal(15, config.HealthCheckIntervalSeconds);
    }

    [Fact]
    public void LoadFromArgs_WithMaxRestarts_SetsMaxRestartAttempts()
    {
        var args = new[] { "--extension-max-restarts", "5" };

        var config = ExtensionConfig.LoadFromArgs(args);

        Assert.Equal(5, config.MaxRestartAttempts);
    }

    [Fact]
    public void LoadFromArgs_WithTransportMode_SetsDefaultTransportMode()
    {
        var args = new[] { "--extension-transport-mode", "mmap" };

        var config = ExtensionConfig.LoadFromArgs(args);

        Assert.Equal("mmap", config.DefaultTransportMode);
    }

    [Fact]
    public void LoadFromArgs_ColonSeparator_Works()
    {
        var args = new[] { "--extension-snapshot-ttl:45" };

        var config = ExtensionConfig.LoadFromArgs(args);

        Assert.Equal(45, config.SnapshotTtlSeconds.Default);
    }

    [Fact]
    public void LoadFromArgs_EqualsSeparator_Works()
    {
        var args = new[] { "--extension-snapshot-ttl=90" };

        var config = ExtensionConfig.LoadFromArgs(args);

        Assert.Equal(90, config.SnapshotTtlSeconds.Default);
    }

    #endregion

    #region Validate Tests

    [Fact]
    public void Validate_Disabled_DoesNotThrow()
    {
        var config = new ExtensionConfig { Enabled = false };

        var exception = Record.Exception(() => config.Validate());

        Assert.Null(exception);
    }

    [Fact]
    public void Validate_Enabled_WithDefaults_DoesNotThrow()
    {
        var config = new ExtensionConfig { Enabled = true };

        var exception = Record.Exception(() => config.Validate());

        Assert.Null(exception);
    }

    [Fact]
    public void Validate_InvalidHealthCheckInterval_Throws()
    {
        var config = new ExtensionConfig { Enabled = true, HealthCheckIntervalSeconds = 0 };

        var exception = Assert.Throws<InvalidOperationException>(() => config.Validate());

        Assert.Contains("HealthCheckIntervalSeconds", exception.Message);
    }

    [Fact]
    public void Validate_NegativeMaxRestarts_Throws()
    {
        var config = new ExtensionConfig { Enabled = true, MaxRestartAttempts = -1 };

        var exception = Assert.Throws<InvalidOperationException>(() => config.Validate());

        Assert.Contains("MaxRestartAttempts", exception.Message);
    }

    [Fact]
    public void Validate_InvalidTransportMode_Throws()
    {
        var config = new ExtensionConfig { Enabled = true, DefaultTransportMode = "invalid" };

        var exception = Assert.Throws<InvalidOperationException>(() => config.Validate());

        Assert.Contains("DefaultTransportMode", exception.Message);
    }

    [Fact]
    public void Validate_ValidTransportModes_DoNotThrow()
    {
        var validModes = new[] { "mmap", "stdin", "file" };

        foreach (var mode in validModes)
        {
            var config = new ExtensionConfig { Enabled = true, DefaultTransportMode = mode };
            var exception = Record.Exception(() => config.Validate());
            Assert.Null(exception);
        }
    }

    #endregion

    #region Validate With SessionConfig Tests

    [Fact]
    public void Validate_WithSessionConfig_Disabled_DoesNotThrow()
    {
        var config = new ExtensionConfig { Enabled = false };
        var sessionConfig = new SessionConfig { Enabled = false };

        var exception = Record.Exception(() => config.Validate(sessionConfig));

        Assert.Null(exception);
    }

    [Fact]
    public void Validate_WithSessionConfig_EnabledWithSession_DoesNotThrow()
    {
        var config = new ExtensionConfig { Enabled = true };
        var sessionConfig = new SessionConfig { Enabled = true };

        var exception = Record.Exception(() => config.Validate(sessionConfig));

        Assert.Null(exception);
    }

    [Fact]
    public void Validate_WithSessionConfig_EnabledWithoutSession_Throws()
    {
        var config = new ExtensionConfig { Enabled = true };
        var sessionConfig = new SessionConfig { Enabled = false };

        var exception = Assert.Throws<InvalidOperationException>(() => config.Validate(sessionConfig));

        Assert.Contains("Session", exception.Message);
    }

    #endregion

    #region Property Setters Tests

    [Fact]
    public void Enabled_CanBeSet()
    {
        var config = new ExtensionConfig { Enabled = true };

        Assert.True(config.Enabled);
    }

    [Fact]
    public void ConfigPath_CanBeSet()
    {
        var config = new ExtensionConfig { ConfigPath = "custom.json" };

        Assert.Equal("custom.json", config.ConfigPath);
    }

    [Fact]
    public void TempDirectory_HasDefaultValue()
    {
        var config = new ExtensionConfig();

        Assert.False(string.IsNullOrEmpty(config.TempDirectory));
        Assert.Equal(Path.GetTempPath(), config.TempDirectory);
    }

    [Fact]
    public void ConstrainedInt_CanModifyDefaults()
    {
        var config = new ExtensionConfig
        {
            FrameIntervalMs = { Default = 200 },
            SnapshotTtlSeconds = { Default = 60 }
        };

        Assert.Equal(200, config.FrameIntervalMs.Default);
        Assert.Equal(60, config.SnapshotTtlSeconds.Default);
    }

    [Fact]
    public void ConstrainedInt_CanModifyFloorCeiling()
    {
        var config = new ExtensionConfig
        {
            FrameIntervalMs = { Floor = 50, Ceiling = 2000 }
        };

        Assert.Equal(50, config.FrameIntervalMs.Floor);
        Assert.Equal(2000, config.FrameIntervalMs.Ceiling);
    }

    #endregion

    #region Additional Validation Tests

    [Fact]
    public void Validate_HealthCheckIntervalTooHigh_Throws()
    {
        var config = new ExtensionConfig { Enabled = true, HealthCheckIntervalSeconds = 3601 };

        var exception = Assert.Throws<InvalidOperationException>(() => config.Validate());

        Assert.Contains("HealthCheckIntervalSeconds", exception.Message);
        Assert.Contains("3600", exception.Message);
    }

    [Fact]
    public void Validate_MaxRestartsTooHigh_Throws()
    {
        var config = new ExtensionConfig { Enabled = true, MaxRestartAttempts = 101 };

        var exception = Assert.Throws<InvalidOperationException>(() => config.Validate());

        Assert.Contains("MaxRestartAttempts", exception.Message);
        Assert.Contains("100", exception.Message);
    }

    [Fact]
    public void Validate_NegativeRestartCooldown_Throws()
    {
        var config = new ExtensionConfig { Enabled = true, RestartCooldownSeconds = -1 };

        var exception = Assert.Throws<InvalidOperationException>(() => config.Validate());

        Assert.Contains("RestartCooldownSeconds", exception.Message);
    }

    [Fact]
    public void Validate_NegativeErrorRecoveryCooldown_Throws()
    {
        var config = new ExtensionConfig { Enabled = true, ErrorRecoveryCooldownSeconds = -1 };

        var exception = Assert.Throws<InvalidOperationException>(() => config.Validate());

        Assert.Contains("ErrorRecoveryCooldownSeconds", exception.Message);
    }

    [Fact]
    public void Validate_ErrorRecoveryCooldownTooHigh_Throws()
    {
        var config = new ExtensionConfig { Enabled = true, ErrorRecoveryCooldownSeconds = 3601 };

        var exception = Assert.Throws<InvalidOperationException>(() => config.Validate());

        Assert.Contains("ErrorRecoveryCooldownSeconds", exception.Message);
    }

    [Fact]
    public void Validate_GracefulShutdownTimeoutTooLow_Throws()
    {
        var config = new ExtensionConfig { Enabled = true, GracefulShutdownTimeoutSeconds = 0 };

        var exception = Assert.Throws<InvalidOperationException>(() => config.Validate());

        Assert.Contains("GracefulShutdownTimeoutSeconds", exception.Message);
    }

    [Fact]
    public void Validate_StdinWriteTimeoutTooLow_Throws()
    {
        var config = new ExtensionConfig { Enabled = true, StdinWriteTimeoutMs = 999 };

        var exception = Assert.Throws<InvalidOperationException>(() => config.Validate());

        Assert.Contains("StdinWriteTimeoutMs", exception.Message);
    }

    [Fact]
    public void Validate_StdinWriteTimeoutTooHigh_Throws()
    {
        var config = new ExtensionConfig { Enabled = true, StdinWriteTimeoutMs = 60001 };

        var exception = Assert.Throws<InvalidOperationException>(() => config.Validate());

        Assert.Contains("StdinWriteTimeoutMs", exception.Message);
    }

    [Fact]
    public void Validate_NegativeMaxConsecutiveSendFailures_Throws()
    {
        var config = new ExtensionConfig { Enabled = true, MaxConsecutiveSendFailures = -1 };

        var exception = Assert.Throws<InvalidOperationException>(() => config.Validate());

        Assert.Contains("MaxConsecutiveSendFailures", exception.Message);
    }

    [Fact]
    public void Validate_RetryLoopTimeoutTooLow_Throws()
    {
        var config = new ExtensionConfig { Enabled = true, RetryLoopTimeoutSeconds = 4 };

        var exception = Assert.Throws<InvalidOperationException>(() => config.Validate());

        Assert.Contains("RetryLoopTimeoutSeconds", exception.Message);
    }

    [Fact]
    public void Validate_RetryLoopTimeoutTooHigh_Throws()
    {
        var config = new ExtensionConfig { Enabled = true, RetryLoopTimeoutSeconds = 301 };

        var exception = Assert.Throws<InvalidOperationException>(() => config.Validate());

        Assert.Contains("RetryLoopTimeoutSeconds", exception.Message);
    }

    [Fact]
    public void Validate_ConversionCacheTtlTooLow_Throws()
    {
        var config = new ExtensionConfig { Enabled = true, ConversionCacheTtlSeconds = 0 };

        var exception = Assert.Throws<InvalidOperationException>(() => config.Validate());

        Assert.Contains("ConversionCacheTtlSeconds", exception.Message);
    }

    [Fact]
    public void Validate_MaxConversionCacheSizeTooLow_Throws()
    {
        var config = new ExtensionConfig { Enabled = true, MaxConversionCacheSize = 0 };

        var exception = Assert.Throws<InvalidOperationException>(() => config.Validate());

        Assert.Contains("MaxConversionCacheSize", exception.Message);
    }

    [Fact]
    public void Validate_MaxConversionCacheSizeTooHigh_Throws()
    {
        var config = new ExtensionConfig { Enabled = true, MaxConversionCacheSize = 10001 };

        var exception = Assert.Throws<InvalidOperationException>(() => config.Validate());

        Assert.Contains("MaxConversionCacheSize", exception.Message);
    }

    [Fact]
    public void Validate_MaxConversionFailuresTooLow_Throws()
    {
        var config = new ExtensionConfig { Enabled = true, MaxConversionFailures = 0 };

        var exception = Assert.Throws<InvalidOperationException>(() => config.Validate());

        Assert.Contains("MaxConversionFailures", exception.Message);
    }

    [Fact]
    public void Validate_NegativeFailureBackoff_Throws()
    {
        var config = new ExtensionConfig { Enabled = true, FailureBackoffSeconds = -1 };

        var exception = Assert.Throws<InvalidOperationException>(() => config.Validate());

        Assert.Contains("FailureBackoffSeconds", exception.Message);
    }

    [Fact]
    public void Validate_FailureBackoffTooHigh_Throws()
    {
        var config = new ExtensionConfig { Enabled = true, FailureBackoffSeconds = 3601 };

        var exception = Assert.Throws<InvalidOperationException>(() => config.Validate());

        Assert.Contains("FailureBackoffSeconds", exception.Message);
    }

    [Fact]
    public void Validate_MaxSnapshotSizeTooLow_Throws()
    {
        var config = new ExtensionConfig { Enabled = true, MaxSnapshotSizeBytes = 1024 * 1024 - 1 };

        var exception = Assert.Throws<InvalidOperationException>(() => config.Validate());

        Assert.Contains("MaxSnapshotSizeBytes", exception.Message);
    }

    [Fact]
    public void Validate_MaxSnapshotSizeTooHigh_Throws()
    {
        var config = new ExtensionConfig { Enabled = true, MaxSnapshotSizeBytes = 1024L * 1024 * 1024 + 1 };

        var exception = Assert.Throws<InvalidOperationException>(() => config.Validate());

        Assert.Contains("MaxSnapshotSizeBytes", exception.Message);
    }

    [Fact]
    public void Validate_NegativeMinFreeDiskSpace_Throws()
    {
        var config = new ExtensionConfig { Enabled = true, MinFreeDiskSpaceBytes = -1 };

        var exception = Assert.Throws<InvalidOperationException>(() => config.Validate());

        Assert.Contains("MinFreeDiskSpaceBytes", exception.Message);
    }

    [Fact]
    public void Validate_MinFreeDiskSpaceTooHigh_Throws()
    {
        var config = new ExtensionConfig { Enabled = true, MinFreeDiskSpaceBytes = 10L * 1024 * 1024 * 1024 + 1 };

        var exception = Assert.Throws<InvalidOperationException>(() => config.Validate());

        Assert.Contains("MinFreeDiskSpaceBytes", exception.Message);
    }

    [Fact]
    public void Validate_EmptyTempDirectory_Throws()
    {
        var config = new ExtensionConfig { Enabled = true, TempDirectory = "" };

        var exception = Assert.Throws<InvalidOperationException>(() => config.Validate());

        Assert.Contains("TempDirectory", exception.Message);
    }

    [Fact]
    public void Validate_NegativeDebounceDelay_Throws()
    {
        var config = new ExtensionConfig { Enabled = true, DebounceDelayMs = -1 };

        var exception = Assert.Throws<InvalidOperationException>(() => config.Validate());

        Assert.Contains("DebounceDelayMs", exception.Message);
    }

    [Fact]
    public void Validate_DebounceDelayTooHigh_Throws()
    {
        var config = new ExtensionConfig { Enabled = true, DebounceDelayMs = 10001 };

        var exception = Assert.Throws<InvalidOperationException>(() => config.Validate());

        Assert.Contains("DebounceDelayMs", exception.Message);
    }

    [Fact]
    public void Validate_ConstrainedInt_FloorBelowMinimum_Throws()
    {
        var config = new ExtensionConfig
        {
            Enabled = true,
            FrameIntervalMs = { Floor = 0 }
        };

        var exception = Assert.Throws<InvalidOperationException>(() => config.Validate());

        Assert.Contains("FrameIntervalMs.Floor", exception.Message);
    }

    [Fact]
    public void Validate_ConstrainedInt_CeilingAboveMaximum_Throws()
    {
        var config = new ExtensionConfig
        {
            Enabled = true,
            FrameIntervalMs = { Ceiling = 60001 }
        };

        var exception = Assert.Throws<InvalidOperationException>(() => config.Validate());

        Assert.Contains("FrameIntervalMs.Ceiling", exception.Message);
    }

    [Fact]
    public void Validate_ConstrainedInt_FloorGreaterThanCeiling_Throws()
    {
        var config = new ExtensionConfig
        {
            Enabled = true,
            FrameIntervalMs = { Floor = 1000, Ceiling = 500 }
        };

        var exception = Assert.Throws<InvalidOperationException>(() => config.Validate());

        Assert.Contains("FrameIntervalMs.Floor", exception.Message);
    }

    [Fact]
    public void Validate_ConstrainedInt_DefaultOutsideRange_Throws()
    {
        var config = new ExtensionConfig
        {
            Enabled = true,
            FrameIntervalMs = { Default = 50000, Floor = 10, Ceiling = 5000 }
        };

        var exception = Assert.Throws<InvalidOperationException>(() => config.Validate());

        Assert.Contains("FrameIntervalMs.Default", exception.Message);
    }

    #endregion

    #region Additional LoadFromArgs Tests

    [Fact]
    public void LoadFromArgs_WithTempDir_SetsTempDirectory()
    {
        var tempPath = Path.GetTempPath();
        var args = new[] { "--extension-temp-dir", tempPath };

        var config = ExtensionConfig.LoadFromArgs(args);

        Assert.Equal(tempPath, config.TempDirectory);
    }

    [Fact]
    public void LoadFromArgs_WithTempDir_ColonSeparator_Works()
    {
        var tempPath = Path.GetTempPath();
        var args = new[] { $"--extension-temp-dir:{tempPath}" };

        var config = ExtensionConfig.LoadFromArgs(args);

        Assert.Equal(tempPath, config.TempDirectory);
    }

    [Fact]
    public void LoadFromArgs_WithTempDir_EqualsSeparator_Works()
    {
        var tempPath = Path.GetTempPath();
        var args = new[] { $"--extension-temp-dir={tempPath}" };

        var config = ExtensionConfig.LoadFromArgs(args);

        Assert.Equal(tempPath, config.TempDirectory);
    }

    [Fact]
    public void LoadFromArgs_WithMaxSnapshotSize_SetsMaxSnapshotSizeBytes()
    {
        var args = new[] { "--extension-max-snapshot-size", "52428800" };

        var config = ExtensionConfig.LoadFromArgs(args);

        Assert.Equal(52428800L, config.MaxSnapshotSizeBytes);
    }

    [Fact]
    public void LoadFromArgs_WithMaxSnapshotSize_ColonSeparator_Works()
    {
        var args = new[] { "--extension-max-snapshot-size:52428800" };

        var config = ExtensionConfig.LoadFromArgs(args);

        Assert.Equal(52428800L, config.MaxSnapshotSizeBytes);
    }

    [Fact]
    public void LoadFromArgs_WithMinFreeDiskSpace_SetsMinFreeDiskSpaceBytes()
    {
        var args = new[] { "--extension-min-free-disk-space", "268435456" };

        var config = ExtensionConfig.LoadFromArgs(args);

        Assert.Equal(268435456L, config.MinFreeDiskSpaceBytes);
    }

    [Fact]
    public void LoadFromArgs_WithRestartCooldown_SetsRestartCooldownSeconds()
    {
        var args = new[] { "--extension-restart-cooldown", "10" };

        var config = ExtensionConfig.LoadFromArgs(args);

        Assert.Equal(10, config.RestartCooldownSeconds);
    }

    [Fact]
    public void LoadFromArgs_WithGracefulShutdownTimeout_SetsValue()
    {
        var args = new[] { "--extension-graceful-shutdown-timeout", "10" };

        var config = ExtensionConfig.LoadFromArgs(args);

        Assert.Equal(10, config.GracefulShutdownTimeoutSeconds);
    }

    [Fact]
    public void LoadFromArgs_WithMaxConversionFailures_SetsValue()
    {
        var args = new[] { "--extension-max-conversion-failures", "10" };

        var config = ExtensionConfig.LoadFromArgs(args);

        Assert.Equal(10, config.MaxConversionFailures);
    }

    [Fact]
    public void LoadFromArgs_WithFailureBackoff_SetsValue()
    {
        var args = new[] { "--extension-failure-backoff", "120" };

        var config = ExtensionConfig.LoadFromArgs(args);

        Assert.Equal(120, config.FailureBackoffSeconds);
    }

    [Fact]
    public void LoadFromArgs_WithConversionCacheTtl_SetsValue()
    {
        var args = new[] { "--extension-conversion-cache-ttl", "10" };

        var config = ExtensionConfig.LoadFromArgs(args);

        Assert.Equal(10, config.ConversionCacheTtlSeconds);
    }

    [Fact]
    public void LoadFromArgs_WithMaxCacheSize_SetsValue()
    {
        var args = new[] { "--extension-max-cache-size", "200" };

        var config = ExtensionConfig.LoadFromArgs(args);

        Assert.Equal(200, config.MaxConversionCacheSize);
    }

    [Fact]
    public void LoadFromArgs_WithFrameIntervalDefault_SetsValue()
    {
        var args = new[] { "--extension-frame-interval-default", "200" };

        var config = ExtensionConfig.LoadFromArgs(args);

        Assert.Equal(200, config.FrameIntervalMs.Default);
    }

    [Fact]
    public void LoadFromArgs_WithFrameIntervalFloor_SetsValue()
    {
        var args = new[] { "--extension-frame-interval-floor", "20" };

        var config = ExtensionConfig.LoadFromArgs(args);

        Assert.Equal(20, config.FrameIntervalMs.Floor);
    }

    [Fact]
    public void LoadFromArgs_WithFrameIntervalCeiling_SetsValue()
    {
        var args = new[] { "--extension-frame-interval-ceiling", "3000" };

        var config = ExtensionConfig.LoadFromArgs(args);

        Assert.Equal(3000, config.FrameIntervalMs.Ceiling);
    }

    [Fact]
    public void LoadFromArgs_WithSnapshotTtlDefault_SetsValue()
    {
        var args = new[] { "--extension-snapshot-ttl-default", "60" };

        var config = ExtensionConfig.LoadFromArgs(args);

        Assert.Equal(60, config.SnapshotTtlSeconds.Default);
    }

    [Fact]
    public void LoadFromArgs_WithSnapshotTtlFloor_SetsValue()
    {
        var args = new[] { "--extension-snapshot-ttl-floor", "10" };

        var config = ExtensionConfig.LoadFromArgs(args);

        Assert.Equal(10, config.SnapshotTtlSeconds.Floor);
    }

    [Fact]
    public void LoadFromArgs_WithSnapshotTtlCeiling_SetsValue()
    {
        var args = new[] { "--extension-snapshot-ttl-ceiling", "600" };

        var config = ExtensionConfig.LoadFromArgs(args);

        Assert.Equal(600, config.SnapshotTtlSeconds.Ceiling);
    }

    [Fact]
    public void LoadFromArgs_WithMaxMissedHeartbeatsDefault_SetsValue()
    {
        var args = new[] { "--extension-max-missed-heartbeats-default", "5" };

        var config = ExtensionConfig.LoadFromArgs(args);

        Assert.Equal(5, config.MaxMissedHeartbeats.Default);
    }

    [Fact]
    public void LoadFromArgs_WithMaxMissedHeartbeatsFloor_SetsValue()
    {
        var args = new[] { "--extension-max-missed-heartbeats-floor", "2" };

        var config = ExtensionConfig.LoadFromArgs(args);

        Assert.Equal(2, config.MaxMissedHeartbeats.Floor);
    }

    [Fact]
    public void LoadFromArgs_WithMaxMissedHeartbeatsCeiling_SetsValue()
    {
        var args = new[] { "--extension-max-missed-heartbeats-ceiling", "10" };

        var config = ExtensionConfig.LoadFromArgs(args);

        Assert.Equal(10, config.MaxMissedHeartbeats.Ceiling);
    }

    [Fact]
    public void LoadFromArgs_WithMaxMissedHeartbeats_SetsDefaultValue()
    {
        var args = new[] { "--extension-max-missed-heartbeats", "5" };

        var config = ExtensionConfig.LoadFromArgs(args);

        Assert.Equal(5, config.MaxMissedHeartbeats.Default);
    }

    [Fact]
    public void LoadFromArgs_WithDebounceDelay_SetsValue()
    {
        var args = new[] { "--extension-debounce-delay", "200" };

        var config = ExtensionConfig.LoadFromArgs(args);

        Assert.Equal(200, config.DebounceDelayMs);
    }

    [Fact]
    public void LoadFromArgs_WithIdleTimeoutDefault_SetsValue()
    {
        var args = new[] { "--extension-idle-timeout-default", "60" };

        var config = ExtensionConfig.LoadFromArgs(args);

        Assert.Equal(60, config.IdleTimeoutMinutes.Default);
    }

    [Fact]
    public void LoadFromArgs_WithIdleTimeoutFloor_SetsValue()
    {
        var args = new[] { "--extension-idle-timeout-floor", "5" };

        var config = ExtensionConfig.LoadFromArgs(args);

        Assert.Equal(5, config.IdleTimeoutMinutes.Floor);
    }

    [Fact]
    public void LoadFromArgs_WithIdleTimeoutCeiling_SetsValue()
    {
        var args = new[] { "--extension-idle-timeout-ceiling", "720" };

        var config = ExtensionConfig.LoadFromArgs(args);

        Assert.Equal(720, config.IdleTimeoutMinutes.Ceiling);
    }

    [Fact]
    public void LoadFromArgs_WithIdleTimeoutSpecialAllowed_SetsTrue()
    {
        var args = new[] { "--extension-idle-timeout-special-allowed" };

        var config = ExtensionConfig.LoadFromArgs(args);

        Assert.True(config.IdleTimeoutMinutes.SpecialAllowed);
    }

    [Fact]
    public void LoadFromArgs_WithIdleTimeoutSpecialDisallowed_SetsFalse()
    {
        var args = new[] { "--extension-idle-timeout-special-disallowed" };

        var config = ExtensionConfig.LoadFromArgs(args);

        Assert.False(config.IdleTimeoutMinutes.SpecialAllowed);
    }

    [Fact]
    public void LoadFromArgs_WithFrameInterval_SetsDefaultValue()
    {
        var args = new[] { "--extension-frame-interval", "250" };

        var config = ExtensionConfig.LoadFromArgs(args);

        Assert.Equal(250, config.FrameIntervalMs.Default);
    }

    [Fact]
    public void LoadFromArgs_WithConfig_ColonSeparator_Works()
    {
        var args = new[] { "--extension-config:custom.json" };

        var config = ExtensionConfig.LoadFromArgs(args);

        Assert.Equal("custom.json", config.ConfigPath);
    }

    [Fact]
    public void LoadFromArgs_WithConfig_EqualsSeparator_Works()
    {
        var args = new[] { "--extension-config=custom.json" };

        var config = ExtensionConfig.LoadFromArgs(args);

        Assert.Equal("custom.json", config.ConfigPath);
    }

    [Fact]
    public void LoadFromArgs_WithTransportMode_ColonSeparator_Works()
    {
        var args = new[] { "--extension-transport-mode:file" };

        var config = ExtensionConfig.LoadFromArgs(args);

        Assert.Equal("file", config.DefaultTransportMode);
    }

    [Fact]
    public void LoadFromArgs_WithTransportMode_EqualsSeparator_Works()
    {
        var args = new[] { "--extension-transport-mode=mmap" };

        var config = ExtensionConfig.LoadFromArgs(args);

        Assert.Equal("mmap", config.DefaultTransportMode);
    }

    #endregion
}
