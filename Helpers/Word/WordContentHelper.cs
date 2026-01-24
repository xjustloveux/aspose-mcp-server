using System.Text;

namespace AsposeMcpServer.Helpers.Word;

/// <summary>
///     Helper class for Word content operations.
/// </summary>
public static class WordContentHelper
{
    /// <summary>
    ///     Cleans text by removing control characters and normalizing whitespace.
    /// </summary>
    /// <param name="text">The text to clean.</param>
    /// <returns>The cleaned text with control characters removed and whitespace normalized.</returns>
    public static string CleanText(string text)
    {
        if (string.IsNullOrEmpty(text))
            return text;

        var sb = new StringBuilder();
        var state = new CleanTextState();

        foreach (var c in text)
            ProcessCharacter(c, sb, state);

        return sb.ToString().Trim();
    }

    /// <summary>
    ///     Processes a single character for text cleaning.
    /// </summary>
    /// <param name="c">The character to process.</param>
    /// <param name="sb">The string builder to append to.</param>
    /// <param name="state">The current text cleaning state.</param>
    private static void ProcessCharacter(char c, StringBuilder sb, CleanTextState state)
    {
        if (ShouldSkipCharacter(c)) return;

        if (c == '\n')
        {
            ProcessNewline(sb, state);
            return;
        }

        if (c == ' ' || c == '\t')
        {
            ProcessWhitespace(sb, state);
            return;
        }

        sb.Append(c);
        state.LastWasNewline = false;
        state.LastWasSpace = false;
    }

    /// <summary>
    ///     Determines if a character should be skipped during cleaning.
    /// </summary>
    /// <param name="c">The character to check.</param>
    /// <returns>True if the character should be skipped; otherwise, false.</returns>
    private static bool ShouldSkipCharacter(char c)
    {
        return c == '\r' || (char.IsControl(c) && c != '\n' && c != '\r' && c != '\t');
    }

    /// <summary>
    ///     Processes a newline character for text cleaning.
    /// </summary>
    /// <param name="sb">The string builder to append to.</param>
    /// <param name="state">The current text cleaning state.</param>
    private static void ProcessNewline(StringBuilder sb, CleanTextState state)
    {
        if (!state.LastWasNewline)
        {
            sb.Append('\n');
            state.LastWasNewline = true;
        }
        else if (sb is [.., '\n'] and not [.., '\n', '\n'])
        {
            sb.Append('\n');
        }

        state.LastWasSpace = false;
    }

    /// <summary>
    ///     Processes a whitespace character for text cleaning.
    /// </summary>
    /// <param name="sb">The string builder to append to.</param>
    /// <param name="state">The current text cleaning state.</param>
    private static void ProcessWhitespace(StringBuilder sb, CleanTextState state)
    {
        if (state is { LastWasSpace: false, LastWasNewline: false })
        {
            sb.Append(' ');
            state.LastWasSpace = true;
        }
    }

    /// <summary>
    ///     State class to track text cleaning context.
    /// </summary>
    private sealed class CleanTextState
    {
        /// <summary>
        ///     Gets or sets whether the last processed character was a newline.
        /// </summary>
        public bool LastWasNewline { get; set; }

        /// <summary>
        ///     Gets or sets whether the last processed character was a space.
        /// </summary>
        public bool LastWasSpace { get; set; }
    }
}
