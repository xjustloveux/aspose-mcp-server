using System.Diagnostics;
using System.Text;
using System.Text.Json;
using AsposeMcpServer.Core.Transport;
using AsposeMcpServer.Helpers;

namespace AsposeMcpServer.Tests.Infrastructure;

/// <summary>
///     R19-RES02, measured rather than judged on shape.
///     <para>
///         Several tools take nested arrays and strings — table data, shape lists — and not every
///         path bounds them per array or per string before the model binder materialises them. The
///         static reading is "unbounded nesting reaches the binder". Whether that is a finding
///         depends on what a payload at the transport's own 10 MiB ceiling actually costs to bind
///         and traverse, and on whether <see cref="TableBudget" /> gets there first.
///     </para>
///     <para>
///         The measurement §32.5 asks for, kept as a test so the answer stays true rather than
///         being a number in a document. It bounds the ratio between what a caller may send and
///         what binding it costs; a failure here means the amplification has become real and the
///         bound has to move in front of the binder.
///     </para>
/// </summary>
public class NestedPayloadAmplificationTests
{
    /// <summary>The transport's own ceiling on one message.</summary>
    private const int TransportCap = 10 * 1024 * 1024;

    [Fact]
    public void APayloadAtTheTransportCap_ShouldNotAmplifyWhenItIsBound()
    {
        // The worst shape a caller can spend those bytes on: the largest number of the smallest
        // possible elements, which is where a per-element object cost turns into amplification.
        var json = ManySmallCells();
        Assert.True(json.Length <= TransportCap,
            $"the fixture payload is {json.Length:N0} bytes, above the transport cap");

        // Per thread, not per process. `GC.GetTotalAllocatedBytes` counts the whole runtime, and
        // xUnit runs collections in parallel — so the first two runs of this measured everyone
        // else's work as well and reported 66x and then 126x for identical code. A benchmark whose
        // number moves with what happens to be running beside it measures nothing.
        var before = GC.GetAllocatedBytesForCurrentThread();
        var clock = Stopwatch.StartNew();

        var bound = JsonSerializer.Deserialize<List<List<string>>>(json);

        var cells = 0;
        foreach (var row in bound!) cells += row.Count;

        clock.Stop();
        var allocated = GC.GetAllocatedBytesForCurrentThread() - before;
        var factor = (double)allocated / json.Length;

        // Measured per thread: ten megabytes bind to roughly one hundred and fifty-five, about
        // fifteen times the bytes the transport accepted. That is a bounded constant, not runaway
        // amplification, so R19-RES02 is answered "no" rather than turned into a finding.
        //
        // Worth recording how nearly it went the other way. The first two runs of this reported
        // 66x and then 126x for identical code, because they used the process-wide allocation
        // counter while xUnit ran other collections alongside — the number was other tests' work.
        // A benchmark that moves with what happens to be running beside it is not a measurement,
        // and acting on one of those readings would have produced a "finding" out of noise.
        Assert.True(factor < 30,
            $"{cells:N0} cells from {json.Length:N0} bytes allocated {allocated:N0} bytes "
            + $"({factor:F1}x) in {clock.ElapsedMilliseconds:N0} ms. Binding costs a bounded "
            + "multiple of what the transport already accepted, which is why the per-array limits "
            + "arriving after the binder are survivable. A failure here means that stopped being "
            + "true and the bound has to move in front of the binder.");

        Assert.True(clock.ElapsedMilliseconds < 10_000,
            $"binding {json.Length:N0} bytes took {clock.ElapsedMilliseconds:N0} ms");
    }

    [Fact]
    public void ThatSamePayload_ShouldBeRefusedByTheTableBudget()
    {
        // The control. Whatever binding costs, the work the caller asked for is refused before any
        // document is touched — which is the guard that makes the ordering above survivable.
        var rows = TableBudget.MaxRows + 1;

        var failure = Assert.ThrowsAny<ArgumentException>(() =>
            TableBudget.EnsureWithinBudget(rows, 1));

        Assert.Contains("row", failure.Message, StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>A payload at the transport ceiling made of the smallest useful cells.</summary>
    /// <returns>The JSON text.</returns>
    private static string ManySmallCells()
    {
        var builder = new StringBuilder(TransportCap);
        builder.Append('[');

        var first = true;
        while (builder.Length < TransportCap - 4096)
        {
            if (!first) builder.Append(',');
            first = false;

            builder.Append("[\"a\",\"b\",\"c\",\"d\",\"e\",\"f\",\"g\",\"h\"]");
        }

        builder.Append(']');
        return builder.ToString();
    }

    [Fact]
    public void ThePayloadThatAmplifies_ShouldBeRefusedBeforeItIsBound()
    {
        // Defence in depth rather than the fix for a finding: the measurement above says fifteen
        // times is survivable, and `PayloadShapeGuard` was written while a contaminated reading
        // said sixty-six. It is kept because it costs one non-allocating pass over a message the
        // transport has already agreed to hold, and it bounds the constant — but the honest
        // verdict on R19-RES02 is "not a finding", not "a finding that was fixed".
        var json = Encoding.UTF8.GetBytes(ManySmallCells());

        Assert.False(PayloadShapeGuard.IsWithinBounds(json, out var reason));
        Assert.Contains("values", reason, StringComparison.Ordinal);
    }

    [Fact]
    public void AnOrdinaryRequest_ShouldPassTheShapeGuard()
    {
        // The control for the control. A bound that refuses real requests is not a bound, it is
        // an outage.
        var json = Encoding.UTF8.GetBytes(
            "{\"jsonrpc\":\"2.0\",\"id\":1,\"method\":\"tools/call\",\"params\":"
            + "{\"name\":\"word_text\",\"arguments\":{\"path\":\"a.docx\","
            + "\"operation\":\"get\",\"paragraphIndex\":3}}}");

        Assert.True(PayloadShapeGuard.IsWithinBounds(json, out _));
    }

    [Fact]
    public void APayloadNestedTooDeeply_ShouldBeRefused()
    {
        var json = Encoding.UTF8.GetBytes(
            new string('[', PayloadShapeGuard.MaxDepth + 8)
            + new string(']', PayloadShapeGuard.MaxDepth + 8));

        Assert.False(PayloadShapeGuard.IsWithinBounds(json, out var reason));
        Assert.Contains("nested", reason, StringComparison.Ordinal);
    }
}
