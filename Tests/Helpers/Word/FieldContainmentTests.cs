using Aspose.Words;
using Aspose.Words.Fields;
using AsposeMcpServer.Helpers.Word;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Helpers.Word;

/// <summary>
///     R8-W03: both containment entry points have to give the same answer.
///     <para>
///         The two-parameter overload decided containment by walking siblings from the outer
///         field's start to its end. A field whose markers sit in different paragraphs has no such
///         path, so the walk simply ran out and reported "not contained" — the exact case the
///         document-order version was written to fix. Nothing called it any more, which is why the
///         disagreement went unnoticed rather than why it was harmless.
///     </para>
/// </summary>
public class FieldContainmentTests : WordTestBase
{
    /// <summary>Both answers for one pair of fields, so they can be compared.</summary>
    /// <param name="document">The document holding them.</param>
    /// <param name="inner">The field that might be nested.</param>
    /// <param name="outer">The field that might contain it.</param>
    /// <returns>The two-parameter answer and the document-order answer.</returns>
    private static (bool Simple, bool Ordered) BothAnswers(
        Document document, Field inner, Field outer)
    {
        return (FieldBoundaryHelper.IsFieldWithinField(inner, outer),
            FieldBoundaryHelper.IsFieldWithinField(inner, outer,
                FieldBoundaryHelper.DocumentOrder(document)));
    }

    /// <summary>Builds a document with an inner field nested inside an outer one.</summary>
    /// <param name="acrossParagraphs">Whether the outer field's result spans a paragraph break.</param>
    /// <returns>The document and its two fields, outer first.</returns>
    private static (Document Document, Field Outer, Field Inner) Nested(bool acrossParagraphs)
    {
        var document = new Document();
        var builder = new DocumentBuilder(document);

        var outer = builder.InsertField("REF bookmark");
        var separator = outer.Separator;
        Assert.NotNull(separator);
        builder.MoveTo(separator);

        if (acrossParagraphs)
            builder.Writeln("result line one");

        var inner = builder.InsertField("PAGE");
        return (document, outer, inner);
    }

    [Fact]
    public void ANestedFieldInTheSameParagraph_ShouldBeContainedByBothEntryPoints()
    {
        var (document, outer, inner) = Nested(false);

        var answers = BothAnswers(document, inner, outer);

        Assert.True(answers.Ordered);
        Assert.Equal(answers.Ordered, answers.Simple);
    }

    [Fact]
    public void ANestedFieldAcrossAParagraphBreak_ShouldBeContainedByBothEntryPoints()
    {
        // This is the case the sibling walk could not see: the markers end up in different
        // paragraphs, so there is no sibling path from one to the other.
        var (document, outer, inner) = Nested(true);

        var answers = BothAnswers(document, inner, outer);

        Assert.True(answers.Ordered);
        Assert.Equal(answers.Ordered, answers.Simple);
    }

    [Fact]
    public void TwoFieldsSideBySide_ShouldNotBeContainedByEitherEntryPoint()
    {
        var document = new Document();
        var builder = new DocumentBuilder(document);
        var first = builder.InsertField("PAGE");
        FieldBoundaryHelper.MoveToAfterField(builder, first);
        var second = builder.InsertField("NUMPAGES");

        var answers = BothAnswers(document, second, first);

        Assert.False(answers.Ordered);
        Assert.Equal(answers.Ordered, answers.Simple);
    }

    [Fact]
    public void AFieldComparedWithItself_ShouldNotBeContainedByEitherEntryPoint()
    {
        var document = new Document();
        var field = new DocumentBuilder(document).InsertField("PAGE");

        var answers = BothAnswers(document, field, field);

        Assert.False(answers.Ordered);
        Assert.Equal(answers.Ordered, answers.Simple);
    }

    [Fact]
    public void AFieldWhoseMarkersAreGone_ShouldNotBeContainedByEitherEntryPoint()
    {
        var (document, outer, inner) = Nested(false);
        inner.Start?.Remove();

        var answers = BothAnswers(document, inner, outer);

        Assert.False(answers.Ordered);
        Assert.Equal(answers.Ordered, answers.Simple);
    }
}
