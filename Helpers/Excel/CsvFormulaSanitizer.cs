using System.Text;

namespace AsposeMcpServer.Helpers.Excel;

/// <summary>
///     Neutralises CSV formula injection (also known as CSV injection) in exported data.
///     A spreadsheet application evaluates any field whose first character is one of
///     <c>=</c>, <c>+</c>, <c>-</c>, <c>@</c>, or a leading tab / carriage return, so a cell whose
///     text came from an untrusted source can execute when the exported file is opened. Prefixing
///     such a field with an apostrophe makes the application treat it as literal text while keeping
///     the original characters visible.
/// </summary>
public static class CsvFormulaSanitizer
{
    /// <summary>
    ///     Characters that cause a spreadsheet application to evaluate a field rather than display it.
    /// </summary>
    private static readonly char[] DangerousLeadingChars = ['=', '+', '-', '@', '\t', '\r'];

    /// <summary>
    ///     Rewrites <paramref name="csv" /> so that no field can be interpreted as a formula.
    ///     Quoting and embedded separators, quotes and newlines are preserved; only fields whose
    ///     first character is dangerous are changed, and only by gaining a leading apostrophe.
    /// </summary>
    /// <param name="csv">The CSV text produced by the exporter.</param>
    /// <param name="separator">The field separator used when the text was produced.</param>
    /// <returns>The sanitised CSV text.</returns>
    public static string Sanitize(string csv, char separator)
    {
        if (string.IsNullOrEmpty(csv)) return csv;

        using var reader = new StringReader(csv);
        var writer = new StringWriter(new StringBuilder(csv.Length + 16));
        Sanitize(reader, writer, separator);
        return writer.ToString();
    }

    /// <summary>
    ///     Rewrites CSV from a reader to a writer, one character at a time.
    ///     <para>
    ///         The export used to be held four times over at once — the save buffer, the byte array
    ///         it was copied into, the decoded string, and the sanitised string — so a file within
    ///         the output limit could still need several times that limit in memory to produce
    ///         (R4-R03). Nothing here grows with the file: the state machine carries one field at a
    ///         time, and a field is one cell.
    ///     </para>
    /// </summary>
    /// <param name="csv">The CSV text produced by the exporter.</param>
    /// <param name="output">Where to write the sanitised text.</param>
    /// <param name="separator">The field separator used when the text was produced.</param>
    public static void Sanitize(TextReader csv, TextWriter output, char separator)
    {
        var field = new StringBuilder();
        var quoted = false;
        var inQuotes = false;

        int next;
        while ((next = csv.Read()) >= 0)
        {
            var c = (char)next;

            if (inQuotes)
            {
                if (c == '"')
                {
                    // A doubled quote is one literal quote inside the field; a single one ends it.
                    if (csv.Peek() == '"')
                    {
                        csv.Read();
                        field.Append('"');
                        continue;
                    }

                    inQuotes = false;
                    continue;
                }

                field.Append(c);
                continue;
            }

            if (c == '"' && field.Length == 0)
            {
                inQuotes = true;
                quoted = true;
                continue;
            }

            if (c == separator)
            {
                AppendField(output, field, quoted, separator);
                output.Write(separator);
                field.Clear();
                quoted = false;
                continue;
            }

            if (c is '\r' or '\n')
            {
                AppendField(output, field, quoted, separator);
                field.Clear();
                quoted = false;
                if (c == '\r' && csv.Peek() == '\n')
                {
                    csv.Read();
                    output.Write("\r\n");
                }
                else
                {
                    output.Write(c);
                }

                continue;
            }

            field.Append(c);
        }

        AppendField(output, field, quoted, separator);
    }

    /// <summary>
    ///     Writes one field, prefixing an apostrophe when the value would otherwise be evaluated and
    ///     quoting the result whenever the emitted text needs it.
    /// </summary>
    /// <param name="output">Destination writer.</param>
    /// <param name="field">The decoded field value.</param>
    /// <param name="wasQuoted">Whether the source field was quoted.</param>
    /// <param name="separator">The field separator in use.</param>
    private static void AppendField(TextWriter output, StringBuilder field, bool wasQuoted, char separator)
    {
        var value = field.ToString();
        if (value.Length > 0 && Array.IndexOf(DangerousLeadingChars, value[0]) >= 0)
            value = "'" + value;

        var needsQuotes = wasQuoted
                          || value.Contains(separator)
                          || value.Contains('"')
                          || value.Contains('\n')
                          || value.Contains('\r');

        if (!needsQuotes)
        {
            output.Write(value);
            return;
        }

        output.Write('"');
        output.Write(value.Replace("\"", "\"\""));
        output.Write('"');
    }
}
