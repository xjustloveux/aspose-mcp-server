using AsposeMcpServer.Core.Tasks;

namespace AsposeMcpServer.Tests.Core.Tasks;

public class TaskConfigTests
{
    [Fact]
    public void DefaultConfig_ShouldHaveExpectedDefaults()
    {
        var config = new TaskConfig();

        Assert.True(config.Enabled);
        Assert.Equal(5, config.MaxConcurrentTasks);
        Assert.Equal(300000, config.DefaultTtlMs);
        Assert.Equal(3600000, config.MaxTtlMs);
        Assert.Equal(5000, config.DefaultPollIntervalMs);
        Assert.Equal(60000, config.CleanupIntervalMs);
    }

    [Fact]
    public void LoadFromArgs_WithNoTasks_ShouldDisable()
    {
        var args = new[] { "--no-tasks" };

        var config = TaskConfig.LoadFromArgs(args);

        Assert.False(config.Enabled);
    }

    [Fact]
    public void LoadFromArgs_WithMaxConcurrent_ShouldSetValue()
    {
        var args = new[] { "--tasks-max-concurrent:10" };

        var config = TaskConfig.LoadFromArgs(args);

        Assert.Equal(10, config.MaxConcurrentTasks);
    }

    [Fact]
    public void LoadFromArgs_WithDefaultTtl_ShouldSetValue()
    {
        var args = new[] { "--tasks-default-ttl:600000" };

        var config = TaskConfig.LoadFromArgs(args);

        Assert.Equal(600000, config.DefaultTtlMs);
    }

    [Fact]
    public void LoadFromArgs_WithMaxTtl_ShouldSetValue()
    {
        var args = new[] { "--tasks-max-ttl:7200000" };

        var config = TaskConfig.LoadFromArgs(args);

        Assert.Equal(7200000, config.MaxTtlMs);
    }

    [Fact]
    public void LoadFromArgs_WithEmptyArgs_ShouldReturnDefaults()
    {
        var args = Array.Empty<string>();

        var config = TaskConfig.LoadFromArgs(args);

        Assert.True(config.Enabled);
        Assert.Equal(5, config.MaxConcurrentTasks);
    }

    [Fact]
    public void LoadFromArgs_WithMultipleFlags_ShouldApplyAll()
    {
        var args = new[]
        {
            "--tasks-max-concurrent:20",
            "--tasks-default-ttl:120000",
            "--tasks-max-ttl:600000"
        };

        var config = TaskConfig.LoadFromArgs(args);

        Assert.True(config.Enabled);
        Assert.Equal(20, config.MaxConcurrentTasks);
        Assert.Equal(120000, config.DefaultTtlMs);
        Assert.Equal(600000, config.MaxTtlMs);
    }

    [Fact]
    public void Validate_WithValidConfig_ShouldNotThrow()
    {
        var config = new TaskConfig();

        var exception = Record.Exception(() => config.Validate());

        Assert.Null(exception);
    }

    [Fact]
    public void Validate_WithZeroMaxConcurrent_ShouldThrow()
    {
        var config = new TaskConfig { MaxConcurrentTasks = 0 };

        Assert.Throws<InvalidOperationException>(() => config.Validate());
    }

    [Fact]
    public void Validate_WithTooHighMaxConcurrent_ShouldThrow()
    {
        var config = new TaskConfig { MaxConcurrentTasks = 101 };

        Assert.Throws<InvalidOperationException>(() => config.Validate());
    }

    [Fact]
    public void Validate_WithTooLowDefaultTtl_ShouldThrow()
    {
        var config = new TaskConfig { DefaultTtlMs = 500 };

        Assert.Throws<InvalidOperationException>(() => config.Validate());
    }

    [Fact]
    public void Validate_WithMaxTtlLessThanDefaultTtl_ShouldThrow()
    {
        var config = new TaskConfig
        {
            DefaultTtlMs = 300000,
            MaxTtlMs = 60000
        };

        Assert.Throws<InvalidOperationException>(() => config.Validate());
    }

    [Fact]
    public void Validate_WithTooLowCleanupInterval_ShouldThrow()
    {
        var config = new TaskConfig { CleanupIntervalMs = 500 };

        Assert.Throws<InvalidOperationException>(() => config.Validate());
    }
}
