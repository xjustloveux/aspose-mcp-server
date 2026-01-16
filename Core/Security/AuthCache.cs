using System.Collections.Concurrent;
using System.Security.Cryptography;
using System.Text;

namespace AsposeMcpServer.Core.Security;

/// <summary>
///     Thread-safe authentication result cache with TTL and size limit.
///     Uses LRU (Least Recently Used) eviction when max size is reached.
/// </summary>
/// <typeparam name="TResult">The type of authentication result to cache</typeparam>
public class AuthCache<TResult> where TResult : class
{
    /// <summary>
    ///     Internal cache storage using concurrent dictionary for thread safety
    /// </summary>
    private readonly ConcurrentDictionary<string, CacheEntry> _cache = new();

    /// <summary>
    ///     Lock object for cleanup and eviction operations
    /// </summary>
    private readonly object _cleanupLock = new();

    /// <summary>
    ///     Maximum number of entries allowed in the cache
    /// </summary>
    private readonly int _maxSize;

    /// <summary>
    ///     Time-to-live for cache entries in seconds
    /// </summary>
    private readonly int _ttlSeconds;

    /// <summary>
    ///     Initializes a new instance of the AuthCache
    /// </summary>
    /// <param name="ttlSeconds">Time-to-live for cache entries in seconds</param>
    /// <param name="maxSize">Maximum number of entries in the cache</param>
    /// <exception cref="ArgumentOutOfRangeException">Thrown when ttlSeconds or maxSize is less than or equal to 0</exception>
    public AuthCache(int ttlSeconds, int maxSize)
    {
        if (ttlSeconds <= 0)
            throw new ArgumentOutOfRangeException(nameof(ttlSeconds), "TTL must be greater than 0");
        if (maxSize <= 0)
            throw new ArgumentOutOfRangeException(nameof(maxSize), "Max size must be greater than 0");

        _ttlSeconds = ttlSeconds;
        _maxSize = maxSize;
    }

    /// <summary>
    ///     Gets the current number of entries in the cache
    /// </summary>
    public int Count => _cache.Count;

    /// <summary>
    ///     Gets a cached result or validates and caches a new one
    /// </summary>
    /// <param name="token">The token to use as cache key (will be hashed)</param>
    /// <param name="validateFunc">Function to call when cache misses</param>
    /// <param name="shouldCache">Predicate to determine if result should be cached (typically only cache successful results)</param>
    /// <returns>The authentication result</returns>
    /// <exception cref="ArgumentNullException">Thrown when token, validateFunc, or shouldCache is null</exception>
    public async Task<TResult> GetOrValidateAsync(
        string token,
        Func<Task<TResult>> validateFunc,
        Func<TResult, bool> shouldCache)
    {
        ArgumentNullException.ThrowIfNull(token);
        ArgumentNullException.ThrowIfNull(validateFunc);
        ArgumentNullException.ThrowIfNull(shouldCache);

        var cacheKey = ComputeCacheKey(token);

        if (_cache.TryGetValue(cacheKey, out var entry))
        {
            if (entry.Expiry > DateTime.UtcNow)
            {
                entry.LastAccess = DateTime.UtcNow;
                return entry.Result;
            }

            _cache.TryRemove(cacheKey, out _);
        }

        var result = await validateFunc();

        if (shouldCache(result))
        {
            EnsureCapacity();

            var newEntry = new CacheEntry
            {
                Result = result,
                Expiry = DateTime.UtcNow.AddSeconds(_ttlSeconds),
                LastAccess = DateTime.UtcNow
            };

            _cache[cacheKey] = newEntry;
        }

        return result;
    }

    /// <summary>
    ///     Removes expired entries from the cache
    /// </summary>
    /// <returns>The number of entries removed</returns>
    public int CleanupExpired()
    {
        var now = DateTime.UtcNow;
        var removed = 0;

        foreach (var kvp in _cache)
            if (kvp.Value.Expiry <= now && _cache.TryRemove(kvp.Key, out _))
                removed++;

        return removed;
    }

    /// <summary>
    ///     Clears all entries from the cache
    /// </summary>
    public void Clear()
    {
        _cache.Clear();
    }

    /// <summary>
    ///     Computes a secure cache key from the token using SHA256 hash.
    ///     This avoids storing the original token in memory.
    /// </summary>
    /// <param name="token">The token to hash</param>
    /// <returns>Base64 encoded hash of the token</returns>
    private static string ComputeCacheKey(string token)
    {
        var bytes = SHA256.HashData(Encoding.UTF8.GetBytes(token));
        return Convert.ToBase64String(bytes);
    }

    /// <summary>
    ///     Ensures the cache has capacity for new entries.
    ///     Uses LRU eviction when max size is reached.
    /// </summary>
    private void EnsureCapacity()
    {
        if (_cache.Count < _maxSize)
            return;

        lock (_cleanupLock)
        {
            if (_cache.Count < _maxSize)
                return;

            CleanupExpired();

            if (_cache.Count < _maxSize)
                return;

            var entriesToRemove = _cache
                .OrderBy(kvp => kvp.Value.LastAccess)
                .Take(_cache.Count - _maxSize + 1)
                .Select(kvp => kvp.Key)
                .ToList();

            foreach (var key in entriesToRemove)
                _cache.TryRemove(key, out _);
        }
    }

    /// <summary>
    ///     Internal cache entry structure
    /// </summary>
    private sealed class CacheEntry
    {
        /// <summary>
        ///     The cached authentication result
        /// </summary>
        public required TResult Result { get; init; }

        /// <summary>
        ///     When this entry expires
        /// </summary>
        public required DateTime Expiry { get; init; }

        /// <summary>
        ///     When this entry was last accessed (for LRU eviction)
        /// </summary>
        public DateTime LastAccess { get; set; }
    }
}
