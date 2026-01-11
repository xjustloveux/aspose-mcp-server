using System.Text;

namespace AsposeMcpServer.Handlers.Word.Content;

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
        var lastWasNewline = false;
        var lastWasSpace = false;

        foreach (var c in text)
        {
            if (char.IsControl(c) && c != '\n' && c != '\r' && c != '\t')
                continue;

            if (c == '\r')
                continue;

            if (c == '\n')
            {
                if (!lastWasNewline)
                {
                    sb.Append('\n');
                    lastWasNewline = true;
                }
                else
                {
                    if (sb is [.., '\n'] and not [.., '\n', '\n'])
                        sb.Append('\n');
                }

                lastWasSpace = false;
                continue;
            }

            if (c == ' ' || c == '\t')
            {
                if (!lastWasSpace && !lastWasNewline)
                {
                    sb.Append(' ');
                    lastWasSpace = true;
                }

                continue;
            }

            sb.Append(c);
            lastWasNewline = false;
            lastWasSpace = false;
        }

        return sb.ToString().Trim();
    }
}
