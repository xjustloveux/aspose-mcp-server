namespace AsposeMcpServer.Errors.Excel;

/// <summary>
///     A cell range this server refuses to work with, and the reason, written here.
///     <para>
///         The statistics handler used to forward the message of any <see cref="ArgumentException" />
///         it caught, on the reasoning that such an exception was authored in this repository. It
///         is not: an <see cref="ArgumentException" /> raised inside the spreadsheet library
///         arrives at the same catch and can name a path (R3-C09). A distinct type is what tells
///         the two apart — a message is only shown to the caller when this server wrote it.
///     </para>
/// </summary>
public class RangeNotUsableException : ArgumentException
{
    /// <summary>
    ///     Creates the exception with a message written in this repository.
    /// </summary>
    /// <param name="message">
    ///     Text that will be shown to the caller. It must be composed here and must not embed the
    ///     message of another exception.
    /// </param>
    public RangeNotUsableException(string message) : base(message)
    {
    }

    /// <summary>
    ///     Creates the exception with the default message.
    /// </summary>
    public RangeNotUsableException()
    {
    }

    /// <summary>
    ///     Creates the exception with a message and an inner exception.
    /// </summary>
    /// <param name="message">Text that will be shown to the caller.</param>
    /// <param name="innerException">The exception that caused this one.</param>
    public RangeNotUsableException(string message, Exception innerException)
        : base(message, innerException)
    {
    }
}
