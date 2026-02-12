namespace AsposeMcpServer.Helpers.Word;

/// <summary>
///     Represents a parsed list item with text content and indentation level.
/// </summary>
/// <param name="Text">The text content of the list item.</param>
/// <param name="Level">The indentation level (0-8).</param>
public record ListItemInfo(string Text, int Level);
