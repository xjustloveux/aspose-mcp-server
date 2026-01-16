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
    public void SessionConfig_IsolationMode_DefaultShouldBeGroup()
    {
        var config = new SessionConfig();

        Assert.Equal(SessionIsolationMode.Group, config.IsolationMode);
    }

    [Fact]
    public void SessionConfig_IsolationMode_ShouldBeSettable()
    {
        var config = new SessionConfig { IsolationMode = SessionIsolationMode.None };

        Assert.Equal(SessionIsolationMode.None, config.IsolationMode);
    }

    [Fact]
    public void SessionConfig_LoadFromArgs_IsolationNone()
    {
        var config = SessionConfig.LoadFromArgs(["--session-isolation:none"]);

        Assert.Equal(SessionIsolationMode.None, config.IsolationMode);
    }

    [Fact]
    public void SessionConfig_LoadFromArgs_IsolationGroup()
    {
        var config = SessionConfig.LoadFromArgs(["--session-isolation:group"]);

        Assert.Equal(SessionIsolationMode.Group, config.IsolationMode);
    }

    [Fact]
    public void SessionConfig_LoadFromArgs_IsolationCaseInsensitive()
    {
        var config = SessionConfig.LoadFromArgs(["--session-isolation:GROUP"]);

        Assert.Equal(SessionIsolationMode.Group, config.IsolationMode);
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
        Environment.SetEnvironmentVariable("ASPOSE_SESSION_ISOLATION", "group");

        try
        {
            var config = SessionConfig.LoadFromArgs([]);

            Assert.Equal(SessionIsolationMode.Group, config.IsolationMode);
        }
        finally
        {
            Environment.SetEnvironmentVariable("ASPOSE_SESSION_ISOLATION", null);
        }
    }

    [Fact]
    public void SessionConfig_LoadFromArgs_IsolationOverridesEnvironment()
    {
        Environment.SetEnvironmentVariable("ASPOSE_SESSION_ISOLATION", "none");

        try
        {
            var config = SessionConfig.LoadFromArgs(["--session-isolation:group"]);

            Assert.Equal(SessionIsolationMode.Group, config.IsolationMode);
        }
        finally
        {
            Environment.SetEnvironmentVariable("ASPOSE_SESSION_ISOLATION", null);
        }
    }

    #endregion

    #region LoadFromEnvironment Tests

    [Fact]
    public void SessionConfig_LoadFromEnvironment_Enabled()
    {
        Environment.SetEnvironmentVariable("ASPOSE_SESSION_ENABLED", "true");

        try
        {
            var config = SessionConfig.LoadFromArgs([]);
            Assert.True(config.Enabled);
        }
        finally
        {
            Environment.SetEnvironmentVariable("ASPOSE_SESSION_ENABLED", null);
        }
    }

    [Fact]
    public void SessionConfig_LoadFromEnvironment_MaxSessions()
    {
        Environment.SetEnvironmentVariable("ASPOSE_SESSION_MAX", "25");

        try
        {
            var config = SessionConfig.LoadFromArgs([]);
            Assert.Equal(25, config.MaxSessions);
        }
        finally
        {
            Environment.SetEnvironmentVariable("ASPOSE_SESSION_MAX", null);
        }
    }

    [Fact]
    public void SessionConfig_LoadFromEnvironment_Timeout()
    {
        Environment.SetEnvironmentVariable("ASPOSE_SESSION_TIMEOUT", "45");

        try
        {
            var config = SessionConfig.LoadFromArgs([]);
            Assert.Equal(45, config.IdleTimeoutMinutes);
        }
        finally
        {
            Environment.SetEnvironmentVariable("ASPOSE_SESSION_TIMEOUT", null);
        }
    }

    [Fact]
    public void SessionConfig_LoadFromEnvironment_MaxFileSizeMb()
    {
        Environment.SetEnvironmentVariable("ASPOSE_SESSION_MAX_FILE_SIZE_MB", "250");

        try
        {
            var config = SessionConfig.LoadFromArgs([]);
            Assert.Equal(250, config.MaxFileSizeMb);
        }
        finally
        {
            Environment.SetEnvironmentVariable("ASPOSE_SESSION_MAX_FILE_SIZE_MB", null);
        }
    }

    [Fact]
    public void SessionConfig_LoadFromEnvironment_TempRetentionHours()
    {
        Environment.SetEnvironmentVariable("ASPOSE_SESSION_TEMP_RETENTION_HOURS", "48");

        try
        {
            var config = SessionConfig.LoadFromArgs([]);
            Assert.Equal(48, config.TempRetentionHours);
        }
        finally
        {
            Environment.SetEnvironmentVariable("ASPOSE_SESSION_TEMP_RETENTION_HOURS", null);
        }
    }

    [Fact]
    public void SessionConfig_LoadFromEnvironment_TempDir()
    {
        Environment.SetEnvironmentVariable("ASPOSE_SESSION_TEMP_DIR", "/custom/temp");

        try
        {
            var config = SessionConfig.LoadFromArgs([]);
            Assert.Equal("/custom/temp", config.TempDirectory);
        }
        finally
        {
            Environment.SetEnvironmentVariable("ASPOSE_SESSION_TEMP_DIR", null);
        }
    }

    [Fact]
    public void SessionConfig_LoadFromEnvironment_OnDisconnect()
    {
        Environment.SetEnvironmentVariable("ASPOSE_SESSION_ON_DISCONNECT", "autosave");

        try
        {
            var config = SessionConfig.LoadFromArgs([]);
            Assert.Equal(DisconnectBehavior.AutoSave, config.OnDisconnect);
        }
        finally
        {
            Environment.SetEnvironmentVariable("ASPOSE_SESSION_ON_DISCONNECT", null);
        }
    }

    #endregion

    #region LoadFromCommandLine Tests

    [Fact]
    public void SessionConfig_LoadFromCommandLine_SessionEnabled()
    {
        var config = SessionConfig.LoadFromArgs(["--session-enabled"]);
        Assert.True(config.Enabled);
    }

    [Fact]
    public void SessionConfig_LoadFromCommandLine_SessionDisabled()
    {
        var config = SessionConfig.LoadFromArgs(["--session-disabled"]);
        Assert.False(config.Enabled);
    }

    [Fact]
    public void SessionConfig_LoadFromCommandLine_MaxSessionsWithColon()
    {
        var config = SessionConfig.LoadFromArgs(["--session-max:15"]);
        Assert.Equal(15, config.MaxSessions);
    }

    [Fact]
    public void SessionConfig_LoadFromCommandLine_MaxSessionsWithEquals()
    {
        var config = SessionConfig.LoadFromArgs(["--session-max=20"]);
        Assert.Equal(20, config.MaxSessions);
    }

    [Fact]
    public void SessionConfig_LoadFromCommandLine_TimeoutWithColon()
    {
        var config = SessionConfig.LoadFromArgs(["--session-timeout:60"]);
        Assert.Equal(60, config.IdleTimeoutMinutes);
    }

    [Fact]
    public void SessionConfig_LoadFromCommandLine_TimeoutWithEquals()
    {
        var config = SessionConfig.LoadFromArgs(["--session-timeout=90"]);
        Assert.Equal(90, config.IdleTimeoutMinutes);
    }

    [Fact]
    public void SessionConfig_LoadFromCommandLine_MaxFileSizeWithColon()
    {
        var config = SessionConfig.LoadFromArgs(["--session-max-file-size:150"]);
        Assert.Equal(150, config.MaxFileSizeMb);
    }

    [Fact]
    public void SessionConfig_LoadFromCommandLine_MaxFileSizeWithEquals()
    {
        var config = SessionConfig.LoadFromArgs(["--session-max-file-size=200"]);
        Assert.Equal(200, config.MaxFileSizeMb);
    }

    [Fact]
    public void SessionConfig_LoadFromCommandLine_TempDirWithColon()
    {
        var config = SessionConfig.LoadFromArgs(["--session-temp-dir:/tmp/sessions"]);
        Assert.Equal("/tmp/sessions", config.TempDirectory);
    }

    [Fact]
    public void SessionConfig_LoadFromCommandLine_TempDirWithEquals()
    {
        var config = SessionConfig.LoadFromArgs(["--session-temp-dir=/var/tmp/sessions"]);
        Assert.Equal("/var/tmp/sessions", config.TempDirectory);
    }

    [Fact]
    public void SessionConfig_LoadFromCommandLine_TempRetentionHoursWithColon()
    {
        var config = SessionConfig.LoadFromArgs(["--session-temp-retention-hours:72"]);
        Assert.Equal(72, config.TempRetentionHours);
    }

    [Fact]
    public void SessionConfig_LoadFromCommandLine_TempRetentionHoursWithEquals()
    {
        var config = SessionConfig.LoadFromArgs(["--session-temp-retention-hours=96"]);
        Assert.Equal(96, config.TempRetentionHours);
    }

    [Fact]
    public void SessionConfig_LoadFromCommandLine_OnDisconnectWithColon()
    {
        var config = SessionConfig.LoadFromArgs(["--session-on-disconnect:discard"]);
        Assert.Equal(DisconnectBehavior.Discard, config.OnDisconnect);
    }

    [Fact]
    public void SessionConfig_LoadFromCommandLine_OnDisconnectWithEquals()
    {
        var config = SessionConfig.LoadFromArgs(["--session-on-disconnect=promptonreconnect"]);
        Assert.Equal(DisconnectBehavior.PromptOnReconnect, config.OnDisconnect);
    }

    [Fact]
    public void SessionConfig_LoadFromCommandLine_MultipleArgs()
    {
        var config = SessionConfig.LoadFromArgs([
            "--session-enabled",
            "--session-max:50",
            "--session-timeout=120",
            "--session-isolation:none"
        ]);

        Assert.True(config.Enabled);
        Assert.Equal(50, config.MaxSessions);
        Assert.Equal(120, config.IdleTimeoutMinutes);
        Assert.Equal(SessionIsolationMode.None, config.IsolationMode);
    }

    #endregion

    #region Validate Tests

    [Fact]
    public void SessionConfig_Validate_DisabledConfig_ShouldNotThrow()
    {
        var config = new SessionConfig
        {
            Enabled = false,
            MaxSessions = 0,
            MaxFileSizeMb = 0
        };

        config.Validate();
    }

    [Fact]
    public void SessionConfig_Validate_MaxSessionsLessThanOne_ShouldThrow()
    {
        var config = new SessionConfig
        {
            Enabled = true,
            MaxSessions = 0
        };

        var ex = Assert.Throws<InvalidOperationException>(() => config.Validate());
        Assert.Contains("MaxSessions must be at least 1", ex.Message);
    }

    [Fact]
    public void SessionConfig_Validate_IdleTimeoutNegative_ShouldThrow()
    {
        var config = new SessionConfig
        {
            Enabled = true,
            IdleTimeoutMinutes = -1
        };

        var ex = Assert.Throws<InvalidOperationException>(() => config.Validate());
        Assert.Contains("IdleTimeoutMinutes cannot be negative", ex.Message);
    }

    [Fact]
    public void SessionConfig_Validate_MaxFileSizeMbLessThanOne_ShouldThrow()
    {
        var config = new SessionConfig
        {
            Enabled = true,
            MaxFileSizeMb = 0
        };

        var ex = Assert.Throws<InvalidOperationException>(() => config.Validate());
        Assert.Contains("MaxFileSizeMb must be at least 1", ex.Message);
    }

    [Fact]
    public void SessionConfig_Validate_TempRetentionHoursLessThanOne_ShouldThrow()
    {
        var config = new SessionConfig
        {
            Enabled = true,
            TempRetentionHours = 0
        };

        var ex = Assert.Throws<InvalidOperationException>(() => config.Validate());
        Assert.Contains("TempRetentionHours must be at least 1", ex.Message);
    }

    [Fact]
    public void SessionConfig_Validate_EmptyTempDirectory_ShouldThrow()
    {
        var config = new SessionConfig
        {
            Enabled = true,
            TempDirectory = ""
        };

        var ex = Assert.Throws<InvalidOperationException>(() => config.Validate());
        Assert.Contains("TempDirectory cannot be empty", ex.Message);
    }

    [Fact]
    public void SessionConfig_Validate_ValidConfig_ShouldNotThrow()
    {
        var config = new SessionConfig
        {
            Enabled = true,
            MaxSessions = 10,
            IdleTimeoutMinutes = 30,
            MaxFileSizeMb = 100,
            TempRetentionHours = 24,
            TempDirectory = "/tmp"
        };

        config.Validate();
    }

    #endregion
}
