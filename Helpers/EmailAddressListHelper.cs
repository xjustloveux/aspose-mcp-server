using System.Text;

namespace AsposeMcpServer.Helpers;

/// <summary>
///     Splits an address list following RFC 5322 quoting rules, and rejects header values
///     carrying line breaks.
///     <para>
///         Splitting on every comma (LOW-07) breaks a display name that legitimately contains
///         one: <c>"Last, First" &lt;user@example.com&gt;</c> became the two fragments
///         <c>"Last</c> and <c>First" &lt;user@example.com&gt;</c>, so the message was addressed
///         to something nobody asked for. A separator only counts when it sits outside a quoted
///         string, an angle-addr and a parenthesised comment. Besides the comma RFC 5322 uses, a
///         semicolon is accepted as a separator too, because callers routinely paste lists in the
///         Outlook style; that is a deliberate leniency beyond the RFC, not part of it.
///     </para>
///     <para>
///         A bare CR or LF in a header value (LOW-08) lets the caller append arbitrary headers,
///         so header inputs are refused at the entry point rather than trusted to whatever the
///         writer happens to do with them.
///     </para>
/// </summary>
public static class EmailAddressListHelper
{
    /// <summary>
    ///     The largest number of recipients one message may address across To, CC and Bcc.
    ///     <para>
    ///         The per-field limit is the same number, and it was the only one applied, so a
    ///         message could be addressed to three times as many recipients as any single field
    ///         allowed simply by spreading them across the three (R3-C08).
    ///     </para>
    /// </summary>
    public const int MaxRecipientsPerMessage = 1000;

    /// <summary>
    ///     Splits a comma- or semicolon-separated address list, honouring quoted display names,
    ///     angle-addrs and comments.
    /// </summary>
    /// <param name="addresses">The raw address list, e.g. <c>"Last, First" &lt;a@b.com&gt;, c@d.com</c>.</param>
    /// <returns>The individual address entries, trimmed, with empty entries removed.</returns>
    /// <exception cref="ArgumentException">
    ///     Thrown when the list contains a line break, is longer than the shared string limit, or
    ///     yields more entries than the shared array limit allows.
    /// </exception>
    public static IReadOnlyList<string> Split(string? addresses)
    {
        if (string.IsNullOrWhiteSpace(addresses)) return [];

        // The shared string and array limits apply here too. This method parses caller-supplied
        // text character by character and materialises one entry per address, so without them a
        // single header value sets both the parsing work and the recipient count (R2-S08).
        SecurityHelper.ValidateStringLength(addresses, nameof(addresses));
        EnsureNoHeaderInjection(addresses, nameof(addresses));

        var result = new List<string>();
        var current = new StringBuilder();
        var inQuotes = false;
        var inAngle = false;
        var commentDepth = 0;

        for (var i = 0; i < addresses.Length; i++)
        {
            var c = addresses[i];

            if (inQuotes)
            {
                current.Append(c);
                if (c == '\\' && i + 1 < addresses.Length)
                {
                    current.Append(addresses[++i]);
                    continue;
                }

                if (c == '"') inQuotes = false;
                continue;
            }

            switch (c)
            {
                case '"':
                    inQuotes = true;
                    current.Append(c);
                    continue;
                case '(' when !inAngle:
                    commentDepth++;
                    current.Append(c);
                    continue;
                case ')' when commentDepth > 0:
                    commentDepth--;
                    current.Append(c);
                    continue;
                case '<':
                    inAngle = true;
                    current.Append(c);
                    continue;
                case '>':
                    inAngle = false;
                    current.Append(c);
                    continue;
                case ',' or ';' when !inAngle && commentDepth == 0:
                    AppendEntry(result, current);
                    continue;
                default:
                    current.Append(c);
                    continue;
            }
        }

        AppendEntry(result, current);
        SecurityHelper.ValidateArraySize(result, nameof(addresses));
        return result;
    }

    /// <summary>
    ///     Rejects a message that would address more recipients than one message may.
    /// </summary>
    /// <param name="total">Recipients the message would carry once the change is applied.</param>
    /// <exception cref="ArgumentException">Thrown when the total exceeds the limit.</exception>
    public static void EnsureRecipientTotal(int total)
    {
        if (total > MaxRecipientsPerMessage)
            throw new ArgumentException(
                $"A message may address at most {MaxRecipientsPerMessage} recipients across To, CC "
                + $"and Bcc, but this change would address {total}.");
    }

    /// <summary>
    ///     Rejects a header value containing CR or LF, which would let the caller inject
    ///     additional headers or a message body.
    /// </summary>
    /// <param name="value">The header value to check. A null or empty value is accepted.</param>
    /// <param name="paramName">Parameter name reported in the exception.</param>
    /// <exception cref="ArgumentException">Thrown when the value contains CR or LF.</exception>
    public static void EnsureNoHeaderInjection(string? value, string paramName)
    {
        if (string.IsNullOrEmpty(value)) return;

        if (value.Contains('\r') || value.Contains('\n'))
            throw new ArgumentException(
                $"{paramName} must not contain line breaks; they would be interpreted as the start of another header.",
                paramName);
    }

    /// <summary>Adds the accumulated entry to the result when it is not blank, then clears it.</summary>
    /// <param name="result">Collected entries.</param>
    /// <param name="current">Buffer holding the entry being read.</param>
    private static void AppendEntry(List<string> result, StringBuilder current)
    {
        var entry = current.ToString().Trim();
        current.Clear();
        if (entry.Length > 0) result.Add(entry);
    }
}
