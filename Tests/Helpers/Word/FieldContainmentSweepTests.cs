using System.Diagnostics;
using AsposeMcpServer.Helpers.Word;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Helpers.Word;

/// <summary>
///     R21-RES03: the nested-field check is a sweep, not every pair.
/// </summary>
[Collection(PerformanceTestCollection.Name)]
public class FieldContainmentSweepTests
{
    [Fact]
    public void AnInnerStrictlyInsideAnOuter_IsFound()
    {
        Assert.Equal(0, FieldBoundaryHelper.FirstContained([(0, 10)], [(2, 5)]));
    }

    [Fact]
    public void AnInnerOnlyOverlappingAnOuter_IsNotContained()
    {
        Assert.Equal(-1, FieldBoundaryHelper.FirstContained([(0, 10)], [(5, 15)]));
    }

    [Fact]
    public void SharingAStartOrEnd_IsNotContainment()
    {
        // Strict on both sides, exactly as the pairwise check was.
        Assert.Equal(-1, FieldBoundaryHelper.FirstContained([(0, 10)], [(0, 5)]));
        Assert.Equal(-1, FieldBoundaryHelper.FirstContained([(0, 10)], [(5, 10)]));
    }

    [Fact]
    public void AnInnerInsideALaterOuter_IsFoundThroughTheOpenSet()
    {
        // Two outers; the first closes before the inner, the second contains it.
        Assert.Equal(1, FieldBoundaryHelper.FirstContained([(0, 3), (4, 20)], [(1, 5), (8, 9)]));
    }

    [Fact]
    public void TheSweep_ShouldAgreeWithThePairwiseCheckOnRandomIntervals()
    {
        var random = new Random(20260908);
        for (var round = 0; round < 200; round++)
        {
            var outers = Enumerable.Range(0, 30).Select(_ => Span(random)).ToList();
            var inners = Enumerable.Range(0, 30).Select(_ => Span(random)).ToList();

            var pairwise = inners.Select((inner, i) => (inner, i))
                .Where(p => outers.Any(o => p.inner.Start > o.Start && p.inner.End < o.End))
                .Select(p => p.i).DefaultIfEmpty(-1).First();
            var sweep = FieldBoundaryHelper.FirstContained(outers, inners);

            // Either both found something or neither did; the sweep reports in start order, the
            // pairwise in list order, so the index may differ when several are contained.
            Assert.Equal(pairwise >= 0, sweep >= 0);
        }
    }

    [Fact]
    public void TenThousandByTenThousand_ShouldFinishInABoundedTime()
    {
        // The pairwise check was a hundred million comparisons here. Balanced, non-nested
        // intervals so nothing short-circuits: the sweep has to visit everything. This collection
        // does not run beside the document-heavy suite, and a warm-up plus median rejects an
        // algorithmic regression without turning one scheduler or JIT pause into a release failure.
        var outers = Enumerable.Range(0, 10_000).Select(i => (i * 4, i * 4 + 1)).ToList();
        var inners = Enumerable.Range(0, 10_000).Select(i => (i * 4 + 2, i * 4 + 3)).ToList();

        Assert.Equal(-1, FieldBoundaryHelper.FirstContained(outers, inners));

        var elapsed = new long[5];
        for (var trial = 0; trial < elapsed.Length; trial++)
        {
            var clock = Stopwatch.StartNew();
            var result = FieldBoundaryHelper.FirstContained(outers, inners);
            clock.Stop();

            Assert.Equal(-1, result);
            elapsed[trial] = clock.ElapsedMilliseconds;
        }

        Array.Sort(elapsed);
        var median = elapsed[elapsed.Length / 2];

        Assert.True(median < 500,
            $"the sweep median was {median} ms for 10,000 x 10,000; trials: "
            + string.Join(", ", elapsed));
    }

    private static (int Start, int End) Span(Random random)
    {
        var a = random.Next(0, 200);
        var b = random.Next(0, 200);
        return a < b ? (a, b) : (b, a + 1);
    }
}
