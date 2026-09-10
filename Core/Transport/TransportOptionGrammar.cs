namespace AsposeMcpServer.Core.Transport;

/// <summary>
///     The one definition of how the transport options are written on a command line.
///     <para>
///         Two places read these options: <see cref="TransportConfig" />, which acts on them, and
///         <see cref="ChildProcessArguments" />, which strips them before launching a stdio child.
///         Each carried its own idea of what a value looked like, and both were wrong in the same
///         way — they consumed the token after <c>--host</c> whatever it was. Written as
///         <c>--host --allowed-path C:\safe</c> the parent took <c>--allowed-path</c> as its host
///         name and the child never saw the path allowlist at all, so a malformed transport option
///         silently widened what the server would open (R3-C01).
///     </para>
/// </summary>
public static class TransportOptionGrammar
{
    /// <summary>Transport-selecting switches, which take no value.</summary>
    public static readonly string[] Flags = ["--stdio", "--http", "--ws", "--websocket"];

    /// <summary>Transport options that carry a value.</summary>
    public static readonly string[] ValueOptions = ["--port", "--host"];

    /// <summary>The separators accepted between an option and an inline value.</summary>
    private static readonly char[] InlineSeparators = [':', '='];

    /// <summary>Whether a token is another option rather than a value.</summary>
    /// <param name="token">The token to classify.</param>
    /// <returns>
    ///     <c>true</c> for <c>--anything</c> and for a single dash followed by a letter. A leading
    ///     dash on its own is not enough: a negative number is a value, not an option.
    /// </returns>
    public static bool IsOption(string? token)
    {
        if (string.IsNullOrEmpty(token) || token[0] != '-') return false;
        return token.StartsWith("--", StringComparison.Ordinal) ||
               (token.Length > 1 && char.IsLetter(token[1]));
    }

    /// <summary>Whether a token names a value-carrying option in its separated form.</summary>
    /// <param name="token">The token to classify.</param>
    /// <returns><c>true</c> for exactly <c>--port</c> or <c>--host</c>.</returns>
    public static bool IsSeparatedValueOption(string token)
    {
        return ValueOptions.Any(option => token.Equals(option, StringComparison.OrdinalIgnoreCase));
    }

    /// <summary>Whether a token carries its value inline, as <c>--port:3000</c> or <c>--host=api</c>.</summary>
    /// <param name="token">The token to classify.</param>
    /// <returns><c>true</c> when the token is a value option followed by a separator.</returns>
    public static bool IsInlineValueOption(string token)
    {
        return ValueOptions.Any(option => InlineSeparators.Any(separator =>
            token.StartsWith(option + separator, StringComparison.OrdinalIgnoreCase)));
    }

    /// <summary>Whether the token following an option is that option's value.</summary>
    /// <param name="option">The value-carrying option, <c>--port</c> or <c>--host</c>.</param>
    /// <param name="next">The token that follows it, or <c>null</c> at the end of the line.</param>
    /// <returns>
    ///     <c>true</c> only when the token can be read as a value for that option: a number for
    ///     <c>--port</c>, and any non-option, non-blank token for <c>--host</c>.
    /// </returns>
    public static bool IsValueFor(string option, string? next)
    {
        if (next is null) return false;

        if (option.Equals("--port", StringComparison.OrdinalIgnoreCase))
            return int.TryParse(next, out _);

        return !string.IsNullOrWhiteSpace(next) && !IsOption(next);
    }
}
