using System.Text.Json;

namespace AsposeMcpServer.Core.Transport;

/// <summary>
///     Bounds the <em>shape</em> of an incoming message, not just its size.
///     <para>
///         A byte cap bounds what a peer may send. It does not bound what binding that message
///         costs: measured per thread on this repository's own benchmark, ten megabytes of the
///         smallest possible JSON elements bind to about <c>155 MB</c> — fifteen times the bytes
///         the transport accepted (R19-RES02). That is a bounded constant, not a finding, and this
///         guard is defence in depth against it rather than a fix for it; an earlier reading of
///         sixty-six times was the process-wide allocation counter measuring other tests. Every
///         per-array and per-string limit this server has arrives after the binder has already
///         built the objects, which is why the bound is worth having in front of it.
///     </para>
///     <para>
///         Counting is done with <see cref="Utf8JsonReader" />, which walks the bytes without
///         materialising anything, so the guard costs a pass over a message the transport has
///         already agreed to hold and nothing per element.
///     </para>
/// </summary>
public static class PayloadShapeGuard
{
    /// <summary>The most values one message may contain, at any depth.</summary>
    /// <remarks>
    ///     Sized against what this server's own limits allow a request to ask for —
    ///     <c>TableBudget.MaxCells</c> is 200,000 — with room for the structure around it. A
    ///     request above this is refused by a handler anyway; the point of refusing here is that
    ///     it is refused before the objects exist.
    /// </remarks>
    public const int MaxValues = 400_000;

    /// <summary>The deepest nesting one message may contain.</summary>
    /// <remarks>
    ///     Nothing this server accepts is nested deeply. A depth limit is what stops a message
    ///     whose cost is in its structure rather than its element count.
    /// </remarks>
    public const int MaxDepth = 64;

    /// <summary>Whether a message is a shape this server will bind.</summary>
    /// <param name="utf8">The message bytes.</param>
    /// <param name="reason">Why it was refused, when it was.</param>
    /// <returns><c>true</c> when the message is within both bounds.</returns>
    /// <remarks>
    ///     Malformed JSON is not this guard's business and is allowed through: the reader below
    ///     would throw on it, and turning a parse error into a transport-level refusal would
    ///     change which component reports it and what the peer is told.
    /// </remarks>
    public static bool IsWithinBounds(ReadOnlySpan<byte> utf8, out string reason)
    {
        reason = string.Empty;

        var reader = new Utf8JsonReader(utf8, new JsonReaderOptions
        {
            CommentHandling = JsonCommentHandling.Skip,
            MaxDepth = MaxDepth
        });

        var values = 0;

        try
        {
            while (reader.Read())
                switch (reader.TokenType)
                {
                    case JsonTokenType.String:
                    case JsonTokenType.Number:
                    case JsonTokenType.True:
                    case JsonTokenType.False:
                    case JsonTokenType.Null:
                    case JsonTokenType.StartArray:
                    case JsonTokenType.StartObject:
                        if (++values > MaxValues)
                        {
                            reason = $"the message holds more than {MaxValues:N0} values, which is "
                                     + "more than this server binds";
                            return false;
                        }

                        break;
                }
        }
        catch (JsonException exception)
        {
            // Either malformed, or nested past MaxDepth. The first is for the JSON-RPC layer to
            // report; the second is this guard's, and is what the message says.
            if (exception.Message.Contains("depth", StringComparison.OrdinalIgnoreCase))
            {
                reason = $"the message is nested deeper than {MaxDepth}, which is deeper than this "
                         + "server binds";
                return false;
            }
        }

        return true;
    }
}
