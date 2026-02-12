using AsposeMcpServer.Core.Conversion;
using AsposeMcpServer.Core.Session;

namespace AsposeMcpServer.Core.Extension;

/// <summary>
///     Information about a session-extension binding.
///     Thread-safe for concurrent access.
/// </summary>
public class SessionBindingInfo
{
    /// <summary>
    ///     Backoff duration after max failures reached.
    /// </summary>
    private readonly TimeSpan _failureBackoffDuration;

    /// <summary>
    ///     Lock object for thread-safe property updates.
    /// </summary>
    private readonly object _lock = new();

    /// <summary>
    ///     Maximum number of consecutive conversion failures before backing off.
    /// </summary>
    private readonly int _maxConversionFailures;

    /// <summary>
    ///     Time when backoff expires (null if not in backoff).
    /// </summary>
    private DateTime? _backoffUntil;

    /// <summary>
    ///     Count of consecutive conversion failures.
    /// </summary>
    private int _conversionFailures;

    /// <summary>
    ///     Backing field for ConversionOptions.
    /// </summary>
    private ConversionOptions _conversionOptions = new();

    /// <summary>
    ///     Backing field for LastSentAt.
    /// </summary>
    private DateTime? _lastSentAt;

    /// <summary>
    ///     Backing field for NeedsSend.
    /// </summary>
    private bool _needsSend = true;

    /// <summary>
    ///     Backing field for OutputFormat.
    /// </summary>
    private string _outputFormat = string.Empty;

    /// <summary>
    ///     Initializes a new instance of the <see cref="SessionBindingInfo" /> class.
    /// </summary>
    /// <param name="maxConversionFailures">Maximum conversion failures before backoff.</param>
    /// <param name="failureBackoffSeconds">Backoff duration in seconds.</param>
    public SessionBindingInfo(int maxConversionFailures = 5, int failureBackoffSeconds = 60)
    {
        _maxConversionFailures = maxConversionFailures;
        _failureBackoffDuration = TimeSpan.FromSeconds(failureBackoffSeconds);
    }

    /// <summary>
    ///     Gets or sets the session ID.
    /// </summary>
    public string SessionId { get; set; } = string.Empty;

    /// <summary>
    ///     Gets or sets the extension ID.
    /// </summary>
    public string ExtensionId { get; set; } = string.Empty;

    /// <summary>
    ///     Gets or sets the session owner identity for accessing isolated sessions.
    /// </summary>
    public SessionIdentity Owner { get; set; } = SessionIdentity.GetAnonymous();

    /// <summary>
    ///     Gets or sets the output format.
    ///     Thread-safe.
    /// </summary>
    public string OutputFormat
    {
        get
        {
            lock (_lock)
            {
                return _outputFormat;
            }
        }
        set
        {
            lock (_lock)
            {
                _outputFormat = value;
            }
        }
    }

    /// <summary>
    ///     Gets or sets the conversion options.
    ///     Thread-safe.
    /// </summary>
    public ConversionOptions ConversionOptions
    {
        get
        {
            lock (_lock)
            {
                return _conversionOptions;
            }
        }
        set
        {
            lock (_lock)
            {
                _conversionOptions = value;
            }
        }
    }

    /// <summary>
    ///     Gets or sets when the binding was created.
    /// </summary>
    public DateTime CreatedAt { get; set; }

    /// <summary>
    ///     Gets or sets when the last snapshot was sent.
    ///     Thread-safe.
    /// </summary>
    public DateTime? LastSentAt
    {
        get
        {
            lock (_lock)
            {
                return _lastSentAt;
            }
        }
        set
        {
            lock (_lock)
            {
                _lastSentAt = value;
            }
        }
    }

    /// <summary>
    ///     Gets or sets whether a snapshot needs to be sent.
    ///     Thread-safe.
    /// </summary>
    public bool NeedsSend
    {
        get
        {
            lock (_lock)
            {
                return _needsSend;
            }
        }
        set
        {
            lock (_lock)
            {
                _needsSend = value;
            }
        }
    }

    /// <summary>
    ///     Gets the count of consecutive conversion failures.
    ///     Thread-safe.
    /// </summary>
    public int ConversionFailures
    {
        get
        {
            lock (_lock)
            {
                return _conversionFailures;
            }
        }
    }

    /// <summary>
    ///     Checks if the binding is currently in backoff due to conversion failures.
    ///     Thread-safe.
    /// </summary>
    /// <returns>True if in backoff period.</returns>
    public bool IsInBackoff()
    {
        lock (_lock)
        {
            if (_backoffUntil == null)
                return false;

            if (DateTime.UtcNow >= _backoffUntil.Value)
            {
                _backoffUntil = null;
                _conversionFailures = 0;
                return false;
            }

            return true;
        }
    }

    /// <summary>
    ///     Records a conversion failure and enters backoff if max failures reached.
    /// </summary>
    /// <returns>True if now in backoff.</returns>
    public bool RecordConversionFailure()
    {
        lock (_lock)
        {
            _conversionFailures++;

            if (_conversionFailures >= _maxConversionFailures)
            {
                _backoffUntil = DateTime.UtcNow.Add(_failureBackoffDuration);
                return true;
            }

            return false;
        }
    }

    /// <summary>
    ///     Resets the conversion failure counter on successful conversion.
    /// </summary>
    public void ResetConversionFailures()
    {
        lock (_lock)
        {
            _conversionFailures = 0;
            _backoffUntil = null;
        }
    }

    /// <summary>
    ///     Atomically updates the last sent time and clears the needs send flag.
    /// </summary>
    /// <param name="sentAt">The time the snapshot was sent.</param>
    public void UpdateLastSent(DateTime sentAt)
    {
        lock (_lock)
        {
            _lastSentAt = sentAt;
            _needsSend = false;
            _conversionFailures = 0;
            _backoffUntil = null;
        }
    }

    /// <summary>
    ///     Atomically updates the output format and sets the needs send flag.
    /// </summary>
    /// <param name="newFormat">The new output format.</param>
    public void UpdateFormat(string newFormat)
    {
        lock (_lock)
        {
            _outputFormat = newFormat;
            _needsSend = true;
            _conversionFailures = 0;
            _backoffUntil = null;
        }
    }

    /// <summary>
    ///     Atomically updates the output format and conversion options, and sets the needs send flag.
    /// </summary>
    /// <param name="newFormat">The new output format.</param>
    /// <param name="newOptions">The new conversion options.</param>
    public void UpdateFormatAndOptions(string newFormat, ConversionOptions newOptions)
    {
        lock (_lock)
        {
            _outputFormat = newFormat;
            _conversionOptions = newOptions;
            _needsSend = true;
            _conversionFailures = 0;
            _backoffUntil = null;
        }
    }

    /// <summary>
    ///     Gets a cache key suffix based on the conversion options.
    ///     Used for cache invalidation when options change.
    /// </summary>
    /// <returns>A hash code representing the current options.</returns>
    public int GetOptionsCacheKey()
    {
        lock (_lock)
        {
            return HashCode.Combine(
                _conversionOptions.JpegQuality,
                _conversionOptions.CsvSeparator,
                _conversionOptions.PdfCompliance,
                _conversionOptions.HtmlEmbedImages,
                _conversionOptions.HtmlSingleFile,
                _conversionOptions.Dpi);
        }
    }
}
