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

    #region Isolation Mode Tests

    [Fact]
    public void SessionConfig_IsolationMode_DefaultShouldBeUser()
    {
        var config = new SessionConfig();

        Assert.Equal(SessionIsolationMode.User, config.IsolationMode);
    }

    [Fact]
    public void SessionConfig_IsolationMode_ShouldBeSettable()
    {
        var config = new SessionConfig { IsolationMode = SessionIsolationMode.Tenant };

        Assert.Equal(SessionIsolationMode.Tenant, config.IsolationMode);
    }

    [Fact]
    public void SessionConfig_LoadFromArgs_IsolationNone()
    {
        var config = SessionConfig.LoadFromArgs(["--session-isolation:none"]);

        Assert.Equal(SessionIsolationMode.None, config.IsolationMode);
    }

    [Fact]
    public void SessionConfig_LoadFromArgs_IsolationTenant()
    {
        var config = SessionConfig.LoadFromArgs(["--session-isolation:tenant"]);

        Assert.Equal(SessionIsolationMode.Tenant, config.IsolationMode);
    }

    [Fact]
    public void SessionConfig_LoadFromArgs_IsolationUser()
    {
        var config = SessionConfig.LoadFromArgs(["--session-isolation:user"]);

        Assert.Equal(SessionIsolationMode.User, config.IsolationMode);
    }

    [Fact]
    public void SessionConfig_LoadFromArgs_IsolationCaseInsensitive()
    {
        var config = SessionConfig.LoadFromArgs(["--session-isolation:TENANT"]);

        Assert.Equal(SessionIsolationMode.Tenant, config.IsolationMode);
    }

    [Fact]
    public void SessionConfig_LoadFromArgs_IsolationWithEqualsSign()
    {
        var config = SessionConfig.LoadFromArgs(["--session-isolation=none"]);

        Assert.Equal(SessionIsolationMode.None, config.IsolationMode);
    }

    [Fact]
    public void SessionConfig_LoadFromEnvironment_Isolation()
    {
        Environment.SetEnvironmentVariable("ASPOSE_SESSION_ISOLATION", "tenant");

        try
        {
            var config = SessionConfig.LoadFromArgs([]);

            Assert.Equal(SessionIsolationMode.Tenant, config.IsolationMode);
        }
        finally
        {
            Environment.SetEnvironmentVariable("ASPOSE_SESSION_ISOLATION", null);
        }
    }

    [Fact]
    public void SessionConfig_LoadFromArgs_IsolationOverridesEnvironment()
    {
        Environment.SetEnvironmentVariable("ASPOSE_SESSION_ISOLATION", "tenant");

        try
        {
            var config = SessionConfig.LoadFromArgs(["--session-isolation:user"]);

            Assert.Equal(SessionIsolationMode.User, config.IsolationMode);
        }
        finally
        {
            Environment.SetEnvironmentVariable("ASPOSE_SESSION_ISOLATION", null);
        }
    }

    #endregion
}