using AsposeMcpServer.Core.Extension;
using AsposeMcpServer.Core.Session;

namespace AsposeMcpServer.Tests.Core.Extension;

/// <summary>
///     Unit tests for ExtensionConfig class.
/// </summary>
public class ExtensionConfigTests
{
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
}
