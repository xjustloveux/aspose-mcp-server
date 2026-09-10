namespace AsposeMcpServer.Core.Transport;

/// <summary>
///     Bounds how long one logical WebSocket message may take to arrive, and how many frames it
///     may be split across.
///     <para>
///         The connection's idle deadline is pushed out by every frame that moves, and a frame
///         carrying one byte moves. A client sending a byte at a time therefore held a child
///         process, a connection slot and a half-assembled message for as long as it liked while
///         never being idle (R3-R02). This budget starts when a message's first frame arrives and
///         does not move until that message is complete.
///     </para>
/// </summary>
/// <param name="maxAssembly">Longest one message may take to arrive.</param>
/// <param name="maxFragments">Largest number of frames one message may be split across.</param>
internal sealed class MessageAssemblyBudget(TimeSpan maxAssembly, int maxFragments)
{
    private DateTimeOffset _started;

    /// <summary>Frames received for the message currently being assembled.</summary>
    public int Fragments { get; private set; }

    /// <summary>
    ///     Records one frame of the message being assembled.
    /// </summary>
    /// <param name="now">The current time, passed in so the budget can be exercised directly.</param>
    /// <returns><c>false</c> when this frame takes the message past either bound.</returns>
    public bool TryAddFragment(DateTimeOffset now)
    {
        if (Fragments == 0) _started = now;
        Fragments++;

        return now - _started <= maxAssembly && Fragments <= maxFragments;
    }

    /// <summary>
    ///     How long the message being assembled still has.
    /// </summary>
    /// <param name="now">The current time.</param>
    /// <param name="whenIdle">What to return when no message is part-way through.</param>
    /// <returns>
    ///     The time left before this message passes its deadline, never negative, or
    ///     <paramref name="whenIdle" /> when nothing is being assembled.
    /// </returns>
    public TimeSpan Remaining(DateTimeOffset now, TimeSpan whenIdle)
    {
        if (Fragments == 0) return whenIdle;

        var left = maxAssembly - (now - _started);
        return left > TimeSpan.Zero ? left : TimeSpan.Zero;
    }

    /// <summary>Marks the message complete, so the next frame starts a new one.</summary>
    public void Complete()
    {
        Fragments = 0;
    }
}
