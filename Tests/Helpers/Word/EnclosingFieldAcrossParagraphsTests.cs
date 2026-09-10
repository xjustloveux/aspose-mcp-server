using Aspose.Words;
using Aspose.Words.Fields;
using AsposeMcpServer.Helpers.Word;
using AsposeMcpServer.Tests.Infrastructure;
using WordParagraph = Aspose.Words.Paragraph;

namespace AsposeMcpServer.Tests.Helpers.Word;

/// <summary>
///     R9-W02: <c>GetEnclosingField</c> has to answer for a field whose markers sit in different
///     paragraphs.
///     <para>
///         It started from the fields of the node's own paragraph and walked <c>NextSibling</c>
///         from the field's start to its end. Both halves fail on a field that straddles a
///         paragraph break: the run after the break has a different parent, so the walk can never
///         reach it, and the field may not be listed on that paragraph's range at all. Text
///         operations use this to avoid inserting into, deleting from or replacing inside a
///         field's code or result, so a node it wrongly calls "not in a field" is one those
///         operations will edit.
///     </para>
///     <para>
///         Every position is asserted against the single-paragraph form of the same document as
///         well, because the two must agree — a helper that is right only when a field happens not
///         to wrap is the defect, not the fix.
///     </para>
/// </summary>
public class EnclosingFieldAcrossParagraphsTests : WordTestBase
{
    /// <summary>
    ///     Builds a document with one REF field, optionally broken across two paragraphs.
    /// </summary>
    /// <param name="acrossParagraphs">Whether the field's result spans a paragraph break.</param>
    /// <returns>The document and the runs to ask about.</returns>
    private static Positions Build(bool acrossParagraphs)
    {
        var document = new Document();
        var builder = new DocumentBuilder(document);

        builder.Writeln("Before the field");
        var field = builder.InsertField("REF bookmark");

        var separator = field.Separator;
        Assert.NotNull(separator);
        builder.MoveTo(separator);
        builder.Write("result-one");
        if (acrossParagraphs)
        {
            builder.InsertParagraph();
            builder.Write("result-two");
        }

        FieldBoundaryHelper.MoveToAfterField(builder, field);
        builder.Write("after-the-field");

        var runs = document.GetChildNodes(NodeType.Run, true).Cast<Run>().ToList();

        Run? WithText(string text)
        {
            return runs.FirstOrDefault(run => run.GetText().Contains(text, StringComparison.Ordinal));
        }

        // The code region's text is written by InsertField itself, so it is found by its content.
        var inCode = runs.FirstOrDefault(run =>
            run.GetText().Contains("REF", StringComparison.Ordinal));

        return new Positions(document, field, inCode, WithText("result-one"),
            acrossParagraphs ? WithText("result-two") : null, WithText("after-the-field"));
    }

    [Fact]
    public void EveryPositionInASingleParagraphField_ShouldBeAnsweredCorrectly()
    {
        var p = Build(false);

        Assert.NotNull(p.InCode);
        Assert.NotNull(p.InResult);
        Assert.NotNull(p.Outside);

        Assert.Same(p.Field.Start, FieldBoundaryHelper.GetEnclosingField(p.InCode!)?.Start);
        Assert.Same(p.Field.Start, FieldBoundaryHelper.GetEnclosingField(p.InResult!)?.Start);
        Assert.Null(FieldBoundaryHelper.GetEnclosingField(p.Outside!));
    }

    [Fact]
    public void ARunInTheCodeOfAFieldThatWraps_ShouldStillBeInsideIt()
    {
        var p = Build(true);

        Assert.NotNull(p.InCode);
        Assert.Same(p.Field.Start, FieldBoundaryHelper.GetEnclosingField(p.InCode!)?.Start);
    }

    [Fact]
    public void ARunInTheResultBeforeTheBreak_ShouldBeInsideTheField()
    {
        var p = Build(true);

        Assert.NotNull(p.InResult);
        Assert.Same(p.Field.Start, FieldBoundaryHelper.GetEnclosingField(p.InResult!)?.Start);
    }

    [Fact]
    public void ARunInTheResultAfterTheBreak_ShouldBeInsideTheField()
    {
        // The position the sibling walk can never reach: a different parent from the field's
        // start, so the walk ends before it.
        var p = Build(true);

        Assert.NotNull(p.InLaterParagraph);
        Assert.Same(p.Field.Start, FieldBoundaryHelper.GetEnclosingField(p.InLaterParagraph!)?.Start);
    }

    [Fact]
    public void TheFieldsOwnEnd_ShouldNotCountAsBeingInsideIt()
    {
        // The boundary this helper has always drawn: the end marker closes the field rather than
        // living in it, and text operations treat it as the first position after the field.
        var p = Build(true);

        Assert.NotNull(p.Field.End);
        Assert.Null(FieldBoundaryHelper.GetEnclosingField(p.Field.End));
    }

    [Fact]
    public void ARunAfterAFieldThatWraps_ShouldNotBeInsideIt()
    {
        var p = Build(true);

        Assert.NotNull(p.Outside);
        Assert.Null(FieldBoundaryHelper.GetEnclosingField(p.Outside!));
    }

    [Fact]
    public void TheAnswersForBothShapes_ShouldAgree()
    {
        var single = Build(false);
        var wrapped = Build(true);

        string Describe(Run? run)
        {
            return run == null ? "absent"
                : FieldBoundaryHelper.GetEnclosingField(run) == null ? "outside" : "inside";
        }

        Assert.Equal(Describe(single.InCode), Describe(wrapped.InCode));
        Assert.Equal(Describe(single.InResult), Describe(wrapped.InResult));
        Assert.Equal(Describe(single.Outside), Describe(wrapped.Outside));

        // And the position that exists only in the wrapped shape has to be inside it.
        Assert.Equal("inside", Describe(wrapped.InLaterParagraph));
    }

    [Fact]
    public void ANestedFieldsRun_ShouldReportTheInnermostField()
    {
        // Fields nest, so "the enclosing field" is the innermost one; reporting the outer one
        // would let an operation edit inside the inner field's code.
        var document = new Document();
        var builder = new DocumentBuilder(document);
        builder.Writeln("Body");

        var outer = builder.InsertField("IF 1 = 1");
        var separator = outer.Separator;
        Assert.NotNull(separator);
        builder.MoveTo(separator);
        var inner = builder.InsertField("PAGE");

        var innerSeparator = inner.Separator;
        Assert.NotNull(innerSeparator);
        builder.MoveTo(innerSeparator);
        builder.Write("inner-result");

        var run = document.GetChildNodes(NodeType.Run, true).Cast<Run>()
            .First(r => r.GetText().Contains("inner-result", StringComparison.Ordinal));

        Assert.Same(inner.Start, FieldBoundaryHelper.GetEnclosingField(run)?.Start);
    }

    [Fact]
    public void ANodeFromAnotherDocument_ShouldNotBeClaimedByAnyField()
    {
        var other = new Document();
        new DocumentBuilder(other).Write("elsewhere");
        var run = other.GetChildNodes(NodeType.Run, true).Cast<Run>().First();

        Assert.Null(FieldBoundaryHelper.GetEnclosingField(run));
    }

    [Fact]
    public void ARunInAParagraphWithNoFields_ShouldNotBeClaimed()
    {
        var p = Build(true);
        var paragraphs = p.Document.GetChildNodes(NodeType.Paragraph, true).Cast<WordParagraph>();

        Assert.NotEmpty(paragraphs);
        Assert.Null(FieldBoundaryHelper.GetEnclosingField(p.Outside!));
    }

    /// <summary>The five positions this helper has to be right about.</summary>
    /// <param name="Document">The document holding them.</param>
    /// <param name="Field">The field itself.</param>
    /// <param name="InCode">A run inside the field's code region.</param>
    /// <param name="InResult">A run inside the field's result, before any break.</param>
    /// <param name="InLaterParagraph">A run in the field's result after a paragraph break.</param>
    /// <param name="Outside">A run outside the field entirely.</param>
    private record Positions(
        Document Document,
        Field Field,
        Run? InCode,
        Run? InResult,
        Run? InLaterParagraph,
        Run? Outside);
}
