using AsposeMcpServer.Core.Security;

namespace AsposeMcpServer.Tests.Core.Security;

/// <summary>
///     Unit tests for AuthCache class
/// </summary>
public class AuthCacheTests
{
    #region Clear Tests

    [Fact]
    public async Task Clear_RemovesAllEntries()
    {
        var cache = new AuthCache<TestResult>(300, 100);

        for (var i = 1; i <= 5; i++)
        {
            var index = i;
            await cache.GetOrValidateAsync(
                $"token-{index}",
                () => Task.FromResult(new TestResult { IsValid = true, Value = $"value-{index}" }),
                r => r.IsValid);
        }

        Assert.Equal(5, cache.Count);

        cache.Clear();

        Assert.Equal(0, cache.Count);
    }

    #endregion

    #region Thread Safety Tests

    [Fact]
    public async Task ConcurrentAccess_ThreadSafe()
    {
        var cache = new AuthCache<TestResult>(300, 1000);
        var tasks = new List<Task>();
        var validateCallCount = 0;

        for (var i = 0; i < 100; i++)
        {
            var token = $"token-{i % 10}";
            tasks.Add(Task.Run(async () =>
            {
                await cache.GetOrValidateAsync(
                    token,
                    async () =>
                    {
                        Interlocked.Increment(ref validateCallCount);
                        await Task.Delay(10);
                        return new TestResult { IsValid = true, Value = token };
                    },
                    r => r.IsValid);
            }));
        }

        await Task.WhenAll(tasks);

        Assert.True(cache.Count <= 10);
    }

    #endregion

    #region Test Helper Class

    private class TestResult
    {
        public bool IsValid { get; init; }
        public string? Value { get; init; }
    }

    #endregion

    #region Constructor Tests

    [Fact]
    public void Constructor_InvalidTtl_ShouldThrowArgumentOutOfRangeException()
    {
        Assert.Throws<ArgumentOutOfRangeException>(() => new AuthCache<TestResult>(0, 100));
        Assert.Throws<ArgumentOutOfRangeException>(() => new AuthCache<TestResult>(-1, 100));
    }

    [Fact]
    public void Constructor_InvalidMaxSize_ShouldThrowArgumentOutOfRangeException()
    {
        Assert.Throws<ArgumentOutOfRangeException>(() => new AuthCache<TestResult>(300, 0));
        Assert.Throws<ArgumentOutOfRangeException>(() => new AuthCache<TestResult>(300, -1));
    }

    [Fact]
    public void Constructor_ValidParameters_ShouldCreateCache()
    {
        var cache = new AuthCache<TestResult>(300, 100);
        Assert.Equal(0, cache.Count);
    }

    #endregion

    #region GetOrValidateAsync Tests

    [Fact]
    public async Task GetOrValidateAsync_CacheMiss_CallsValidateFunc()
    {
        var cache = new AuthCache<TestResult>(300, 100);
        var validateCalled = false;

        var result = await cache.GetOrValidateAsync(
            "test-token",
            async () =>
            {
                validateCalled = true;
                await Task.CompletedTask;
                return new TestResult { IsValid = true, Value = "test" };
            },
            r => r.IsValid);

        Assert.True(validateCalled);
        Assert.True(result.IsValid);
        Assert.Equal("test", result.Value);
        Assert.Equal(1, cache.Count);
    }

    [Fact]
    public async Task GetOrValidateAsync_CacheHit_ReturnsCachedResult()
    {
        var cache = new AuthCache<TestResult>(300, 100);
        var validateCallCount = 0;

        await cache.GetOrValidateAsync(
            "test-token",
            async () =>
            {
                validateCallCount++;
                await Task.CompletedTask;
                return new TestResult { IsValid = true, Value = "original" };
            },
            r => r.IsValid);

        var result2 = await cache.GetOrValidateAsync(
            "test-token",
            async () =>
            {
                validateCallCount++;
                await Task.CompletedTask;
                return new TestResult { IsValid = true, Value = "new" };
            },
            r => r.IsValid);

        Assert.Equal(1, validateCallCount);
        Assert.Equal("original", result2.Value);
    }

    [Fact]
    public async Task GetOrValidateAsync_CacheExpired_CallsValidateFuncAgain()
    {
        var cache = new AuthCache<TestResult>(1, 100);
        var validateCallCount = 0;

        await cache.GetOrValidateAsync(
            "test-token",
            async () =>
            {
                validateCallCount++;
                await Task.CompletedTask;
                return new TestResult { IsValid = true, Value = $"call-{validateCallCount}" };
            },
            r => r.IsValid);

        await Task.Delay(1100);

        var result = await cache.GetOrValidateAsync(
            "test-token",
            async () =>
            {
                validateCallCount++;
                await Task.CompletedTask;
                return new TestResult { IsValid = true, Value = $"call-{validateCallCount}" };
            },
            r => r.IsValid);

        Assert.Equal(2, validateCallCount);
        Assert.Equal("call-2", result.Value);
    }

    [Fact]
    public async Task GetOrValidateAsync_ValidationFailed_DoesNotCache()
    {
        var cache = new AuthCache<TestResult>(300, 100);
        var validateCallCount = 0;

        await cache.GetOrValidateAsync(
            "test-token",
            async () =>
            {
                validateCallCount++;
                await Task.CompletedTask;
                return new TestResult { IsValid = false, Value = "failed" };
            },
            r => r.IsValid);

        await cache.GetOrValidateAsync(
            "test-token",
            async () =>
            {
                validateCallCount++;
                await Task.CompletedTask;
                return new TestResult { IsValid = false, Value = "failed-again" };
            },
            r => r.IsValid);

        Assert.Equal(2, validateCallCount);
        Assert.Equal(0, cache.Count);
    }

    [Fact]
    public async Task GetOrValidateAsync_DifferentTokens_CachesSeparately()
    {
        var cache = new AuthCache<TestResult>(300, 100);

        await cache.GetOrValidateAsync(
            "token-1",
            async () =>
            {
                await Task.CompletedTask;
                return new TestResult { IsValid = true, Value = "value-1" };
            },
            r => r.IsValid);

        await cache.GetOrValidateAsync(
            "token-2",
            async () =>
            {
                await Task.CompletedTask;
                return new TestResult { IsValid = true, Value = "value-2" };
            },
            r => r.IsValid);

        Assert.Equal(2, cache.Count);
    }

    [Fact]
    public async Task GetOrValidateAsync_MaxSizeReached_EvictsOldEntries()
    {
        var cache = new AuthCache<TestResult>(300, 3);

        for (var i = 1; i <= 5; i++)
        {
            var index = i;
            await cache.GetOrValidateAsync(
                $"token-{index}",
                async () =>
                {
                    await Task.CompletedTask;
                    return new TestResult { IsValid = true, Value = $"value-{index}" };
                },
                r => r.IsValid);
        }

        Assert.True(cache.Count <= 3);
    }

    [Fact]
    public async Task GetOrValidateAsync_NullToken_ThrowsArgumentNullException()
    {
        var cache = new AuthCache<TestResult>(300, 100);

        await Assert.ThrowsAsync<ArgumentNullException>(() =>
            cache.GetOrValidateAsync(
                null!,
                () => Task.FromResult(new TestResult { IsValid = true }),
                r => r.IsValid));
    }

    [Fact]
    public async Task GetOrValidateAsync_NullValidateFunc_ThrowsArgumentNullException()
    {
        var cache = new AuthCache<TestResult>(300, 100);

        await Assert.ThrowsAsync<ArgumentNullException>(() =>
            cache.GetOrValidateAsync(
                "test-token",
                null!,
                r => r.IsValid));
    }

    [Fact]
    public async Task GetOrValidateAsync_NullShouldCache_ThrowsArgumentNullException()
    {
        var cache = new AuthCache<TestResult>(300, 100);

        await Assert.ThrowsAsync<ArgumentNullException>(() =>
            cache.GetOrValidateAsync(
                "test-token",
                () => Task.FromResult(new TestResult { IsValid = true }),
                null!));
    }

    #endregion

    #region CleanupExpired Tests

    [Fact]
    public async Task CleanupExpired_RemovesExpiredEntries()
    {
        var cache = new AuthCache<TestResult>(1, 100);

        await cache.GetOrValidateAsync(
            "token-1",
            () => Task.FromResult(new TestResult { IsValid = true }),
            r => r.IsValid);

        await cache.GetOrValidateAsync(
            "token-2",
            () => Task.FromResult(new TestResult { IsValid = true }),
            r => r.IsValid);

        Assert.Equal(2, cache.Count);

        await Task.Delay(1100);

        var removed = cache.CleanupExpired();

        Assert.Equal(2, removed);
        Assert.Equal(0, cache.Count);
    }

    [Fact]
    public async Task CleanupExpired_KeepsNonExpiredEntries()
    {
        var cache = new AuthCache<TestResult>(300, 100);

        await cache.GetOrValidateAsync(
            "token-1",
            () => Task.FromResult(new TestResult { IsValid = true }),
            r => r.IsValid);

        var removed = cache.CleanupExpired();

        Assert.Equal(0, removed);
        Assert.Equal(1, cache.Count);
    }

    #endregion
}
