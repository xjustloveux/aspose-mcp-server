using System.Text;
using AsposeMcpServer.Helpers.Excel;

namespace AsposeMcpServer.Tests.Handlers.Excel.DataImportExport;

/// <summary>
///     Guards RB-41: CSV export used to emit cell text verbatim, so a value beginning with
///     <c>=</c>, <c>+</c>, <c>-</c>, <c>@</c>, a tab or a carriage return was evaluated by the
///     spreadsheet application that opened the exported file.
/// </summary>
public class CsvFormulaSanitizerTests
{
    /// <summary>
    ///     Runs the streaming sanitiser over the same text the string one is given.
    /// </summary>
    /// <param name="csv">The CSV to sanitise.</param>
    /// <param name="separator">The field separator in use.</param>
    /// <returns>What the streaming sanitiser wrote.</returns>
    private static string SanitizeAsAStream(string csv, char separator)
    {
        using var reader = new StringReader(csv);
        var writer = new StringWriter();
        CsvFormulaSanitizer.Sanitize(reader, writer, separator);
        return writer.ToString();
    }

    /// <summary>
    ///     R4-R03: the sanitiser reads and writes as it goes rather than materialising the whole
    ///     export several times over. These are the cases where a character-at-a-time state machine
    ///     could disagree with one that can look ahead freely: a quoted field spanning a newline, a
    ///     doubled quote inside one, a separator inside one, a formula prefix, and multi-byte text
    ///     whose UTF-16 length is not its UTF-8 length.
    /// </summary>
    [Theory]
    [InlineData("=cmd|' /c calc'!A1,plain")]
    [InlineData("\"a,b\",\"line one\nline two\"")]
    [InlineData("\"he said \"\"hello\"\"\",=SUM(A1:A2)")]
    [InlineData("姓名,金額\n王小明,=1+1")]
    [InlineData("\"=formula inside quotes\",@at,+plus,-minus")]
    [InlineData("trailing,\r\nafter crlf,\"quoted\r\ninside\"")]
    [InlineData("")]
    [InlineData(",,,")]
    public void TheStreamingSanitizer_ShouldAgreeWithTheStringOne(string csv)
    {
        Assert.Equal(CsvFormulaSanitizer.Sanitize(csv, ','), SanitizeAsAStream(csv, ','));
    }

    /// <summary>
    ///     R4-R03: what the streaming form is for. The export used to be held four times at once —
    ///     save buffer, byte array, decoded string, sanitised string — so a file inside the output
    ///     limit could need several times that limit in memory to produce. Reading and writing as it
    ///     goes, nothing allocated grows with the file: the state machine carries one field, and a
    ///     field is one cell.
    /// </summary>
    [Fact]
    public void TheStreamingSanitizer_ShouldNotAllocateAlongsideTheWholeFile()
    {
        const int rows = 40_000;
        var csv = new StringBuilder(rows * 24);
        for (var row = 0; row < rows; row++)
            csv.Append("=formula,姓名").Append(row).Append(",plain text\r\n");

        var text = csv.ToString();
        var sizeInChars = (long)text.Length;

        // Both forms do the same per-field work, so what separates them is exactly the copies of
        // the whole file: the string form builds the result and the buffer behind it, the
        // streaming form builds neither.
        using var reader = new StringReader(text);
        var sink = new CountingWriter();
        var beforeStreaming = GC.GetAllocatedBytesForCurrentThread();
        CsvFormulaSanitizer.Sanitize(reader, sink, ',');
        var streaming = GC.GetAllocatedBytesForCurrentThread() - beforeStreaming;

        var beforeWholeFile = GC.GetAllocatedBytesForCurrentThread();
        var materialised = CsvFormulaSanitizer.Sanitize(text, ',');
        var wholeFile = GC.GetAllocatedBytesForCurrentThread() - beforeWholeFile;

        Assert.True(sink.Written > sizeInChars, "the fixture did not actually sanitise the text");
        Assert.Equal(sink.Written, materialised.Length);

        Assert.True(streaming < wholeFile - sizeInChars,
            $"streaming allocated {streaming:N0} bytes against {wholeFile:N0} for the same " +
            $"{sizeInChars:N0} characters, so it is still holding the file rather than one field");
    }

    [Theory]
    [InlineData("=SUM(A1:A9)", "'=SUM(A1:A9)")]
    [InlineData("+1+1", "'+1+1")]
    [InlineData("-2+3", "'-2+3")]
    [InlineData("@SUM(1)", "'@SUM(1)")]
    public void DangerousLeadingCharacter_ShouldBePrefixed(string input, string expected)
    {
        var result = CsvFormulaSanitizer.Sanitize(input, ',');

        Assert.Equal(expected, result);
    }

    [Fact]
    public void PlainValues_ShouldBeUnchanged()
    {
        const string csv = "name,amount\r\nWidget,12\r\n";

        var result = CsvFormulaSanitizer.Sanitize(csv, ',');

        Assert.Equal(csv, result);
    }

    [Fact]
    public void QuotedFieldWithSeparator_ShouldKeepQuotingAndContent()
    {
        const string csv = "\"Last, First\",2\r\n";

        var result = CsvFormulaSanitizer.Sanitize(csv, ',');

        Assert.Equal(csv, result);
    }

    [Fact]
    public void QuotedFormulaField_ShouldBePrefixedAndStayQuoted()
    {
        const string csv = "\"=cmd|'/c calc'!A1\",ok\r\n";

        var result = CsvFormulaSanitizer.Sanitize(csv, ',');

        Assert.Equal("\"'=cmd|'/c calc'!A1\",ok\r\n", result);
    }

    [Fact]
    public void EmbeddedQuotes_ShouldRemainEscaped()
    {
        const string csv = "\"say \"\"hi\"\"\",1\r\n";

        var result = CsvFormulaSanitizer.Sanitize(csv, ',');

        Assert.Equal(csv, result);
    }

    [Fact]
    public void TabAndCarriageReturnLeadIn_ShouldBePrefixed()
    {
        var result = CsvFormulaSanitizer.Sanitize("\"\t=1+1\",x\r\n", ',');

        Assert.Equal("\"'\t=1+1\",x\r\n", result);
    }

    [Fact]
    public void CustomSeparator_ShouldBeHonoured()
    {
        var result = CsvFormulaSanitizer.Sanitize("a;=1+1\r\n", ';');

        Assert.Equal("a;'=1+1\r\n", result);
    }

    [Fact]
    public void EmptyInput_ShouldBeReturnedUnchanged()
    {
        Assert.Equal(string.Empty, CsvFormulaSanitizer.Sanitize(string.Empty, ','));
    }

    [Fact]
    public void FormulaAfterNewline_ShouldBePrefixed()
    {
        var result = CsvFormulaSanitizer.Sanitize("a,b\r\n=1,2\r\n", ',');

        Assert.Equal("a,b\r\n'=1,2\r\n", result);
    }

    /// <summary>
    ///     A writer that keeps nothing, so the measurement above is of the sanitiser rather than of
    ///     the buffer its output happens to land in.
    /// </summary>
    private sealed class CountingWriter : TextWriter
    {
        /// <summary>How many characters were written.</summary>
        public long Written { get; private set; }

        /// <inheritdoc />
        public override Encoding Encoding => Encoding.UTF8;

        /// <inheritdoc />
        public override void Write(char value)
        {
            Written++;
        }

        /// <inheritdoc />
        public override void Write(string? value)
        {
            Written += value?.Length ?? 0;
        }
    }
}
