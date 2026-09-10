using AsposeMcpServer.Core.Transport;

namespace AsposeMcpServer.Tests.Core.Transport;

/// <summary>
///     A WebSocket connection's idle deadline is pushed out by every frame that moves, and a frame
///     carrying one byte moves. Sending a byte at a time held a child process and a connection slot
///     indefinitely while never being idle (R3-R02).
/// </summary>
public class MessageAssemblyBudgetTests
{
    private static readonly DateTimeOffset Start = new(2026, 9, 5, 12, 0, 0, TimeSpan.Zero);

    [Fact]
    public void TryAddFragment_WithAMessageThatArrivesPromptly_ShouldBeAccepted()
    {
        var budget = new MessageAssemblyBudget(TimeSpan.FromSeconds(60), 10);

        Assert.True(budget.TryAddFragment(Start));
        Assert.True(budget.TryAddFragment(Start.AddSeconds(1)));
        Assert.True(budget.TryAddFragment(Start.AddSeconds(59)));
    }

    [Fact]
    public void TryAddFragment_PastTheAssemblyDeadline_ShouldBeRefused()
    {
        var budget = new MessageAssemblyBudget(TimeSpan.FromSeconds(60), 10_000);

        Assert.True(budget.TryAddFragment(Start));
        Assert.False(budget.TryAddFragment(Start.AddSeconds(61)));
    }

    [Fact]
    public void TryAddFragment_AtExactlyTheDeadline_ShouldStillBeAccepted()
    {
        var budget = new MessageAssemblyBudget(TimeSpan.FromSeconds(60), 10_000);

        Assert.True(budget.TryAddFragment(Start));
        Assert.True(budget.TryAddFragment(Start.AddSeconds(60)));
    }

    [Fact]
    public void TryAddFragment_PastTheFragmentCap_ShouldBeRefused()
    {
        var budget = new MessageAssemblyBudget(TimeSpan.FromSeconds(60), 3);

        Assert.True(budget.TryAddFragment(Start));
        Assert.True(budget.TryAddFragment(Start));
        Assert.True(budget.TryAddFragment(Start));
        Assert.False(budget.TryAddFragment(Start));
    }

    /// <summary>
    ///     The deadline belongs to one message, not to the connection: a client that keeps sending
    ///     complete messages must not be cut off.
    /// </summary>
    [Fact]
    public void Complete_ShouldStartTheNextMessageAfresh()
    {
        var budget = new MessageAssemblyBudget(TimeSpan.FromSeconds(60), 3);

        Assert.True(budget.TryAddFragment(Start));
        Assert.True(budget.TryAddFragment(Start.AddSeconds(50)));
        budget.Complete();

        Assert.Equal(0, budget.Fragments);
        Assert.True(budget.TryAddFragment(Start.AddSeconds(200)));
        Assert.True(budget.TryAddFragment(Start.AddSeconds(250)));
    }

    /// <summary>
    ///     The budget has to be readable while a message is part-way through, because waiting for
    ///     the next frame is part of the deadline: checking only when a frame arrived meant a
    ///     client that sent one non-final frame and stopped was never measured again (R4-S08).
    /// </summary>
    [Fact]
    public void Remaining_WithNothingBeingAssembled_ShouldReturnTheIdleValue()
    {
        var budget = new MessageAssemblyBudget(TimeSpan.FromSeconds(60), 10);

        Assert.Equal(TimeSpan.FromSeconds(60), budget.Remaining(Start, TimeSpan.FromSeconds(60)));
    }

    [Fact]
    public void Remaining_PartWayThroughAMessage_ShouldCountDown()
    {
        var budget = new MessageAssemblyBudget(TimeSpan.FromSeconds(60), 10);
        budget.TryAddFragment(Start);

        Assert.Equal(TimeSpan.FromSeconds(45), budget.Remaining(Start.AddSeconds(15), TimeSpan.FromSeconds(60)));
    }

    [Fact]
    public void Remaining_PastTheDeadline_ShouldBeZeroRatherThanNegative()
    {
        var budget = new MessageAssemblyBudget(TimeSpan.FromSeconds(60), 10);
        budget.TryAddFragment(Start);

        Assert.Equal(TimeSpan.Zero, budget.Remaining(Start.AddSeconds(90), TimeSpan.FromSeconds(60)));
    }

    [Fact]
    public void Remaining_AfterTheMessageCompletes_ShouldReturnTheIdleValueAgain()
    {
        var budget = new MessageAssemblyBudget(TimeSpan.FromSeconds(60), 10);
        budget.TryAddFragment(Start);
        budget.Complete();

        Assert.Equal(TimeSpan.FromSeconds(60),
            budget.Remaining(Start.AddSeconds(30), TimeSpan.FromSeconds(60)));
    }
}
