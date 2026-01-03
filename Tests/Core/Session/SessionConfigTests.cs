using AsposeMcpServer.Core.Session;

namespace AsposeMcpServer.Tests.Core.Session;

/// <summary>
///     Unit tests for SessionConfig class
/// </summary>
public class SessionConfigTests
{
    [Fact]
    public void SessionConfig_DefaultValues_ShouldBeCorrect()
    {
        var config = new SessionConfig();

        Assert.False(config.Enabled);
        Assert.Equal(DisconnectBehavior.SaveToTemp, config.OnDisconnect);
        Assert.Equal(30, config.IdleTimeoutMinutes);
        Assert.Equal(10, config.MaxSessions);
        Assert.Equal(100, config.MaxFileSizeMb);
        Assert.Equal(24, config.TempRetentionHours);
    }

    [Fact]
    public void SessionConfig_TempDirectory_ShouldHaveDefault()
    {
        var config = new SessionConfig();

        Assert.NotNull(config.TempDirectory);
        Assert.NotEmpty(config.TempDirectory);
        Assert.Equal(Path.GetTempPath(), config.TempDirectory);
    }

    [Fact]
    public void SessionConfig_Enabled_ShouldBeSettable()
    {
        var config = new SessionConfig { Enabled = true };

        Assert.True(config.Enabled);
    }

    [Fact]
    public void SessionConfig_OnDisconnect_ShouldBeSettable()
    {
        var config = new SessionConfig { OnDisconnect = DisconnectBehavior.AutoSave };

        Assert.Equal(DisconnectBehavior.AutoSave, config.OnDisconnect);
    }

    [Fact]
    public void SessionConfig_IdleTimeoutMinutes_ShouldBeSettable()
    {
        var config = new SessionConfig { IdleTimeoutMinutes = 60 };

        Assert.Equal(60, config.IdleTimeoutMinutes);
    }

    [Fact]
    public void SessionConfig_IdleTimeoutMinutes_ZeroMeansNoTimeout()
    {
        var config = new SessionConfig { IdleTimeoutMinutes = 0 };

        Assert.Equal(0, config.IdleTimeoutMinutes);
    }

    [Fact]
    public void SessionConfig_MaxSessions_ShouldBeSettable()
    {
        var config = new SessionConfig { MaxSessions = 50 };

        Assert.Equal(50, config.MaxSessions);
    }

    [Fact]
    public void SessionConfig_MaxFileSizeMb_ShouldBeSettable()
    {
        var config = new SessionConfig { MaxFileSizeMb = 200 };

        Assert.Equal(200, config.MaxFileSizeMb);
    }

    [Fact]
    public void SessionConfig_TempRetentionHours_ShouldBeSettable()
    {
        var config = new SessionConfig { TempRetentionHours = 48 };

        Assert.Equal(48, config.TempRetentionHours);
    }

    [Fact]
    public void SessionConfig_TempDirectory_ShouldBeSettable()
    {
        var customPath = "C:\\CustomTemp";
        var config = new SessionConfig { TempDirectory = customPath };

        Assert.Equal(customPath, config.TempDirectory);
    }
}