using Aspose.Words;
using AsposeMcpServer.Core.Handlers;
using WordParagraph = Aspose.Words.Paragraph;

namespace AsposeMcpServer.Handlers.Word.Field;

/// <summary>
///     Handler for inserting fields in Word documents.
/// </summary>
public class InsertFieldWordHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "insert_field";

    /// <summary>
    ///     Inserts a field at the specified paragraph position.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: fieldType (DATE, TIME, PAGE, NUMPAGES, AUTHOR, etc.)
    ///     Optional: fieldArgument, paragraphIndex, insertAtStart
    /// </param>
    /// <returns>Success message with field details.</returns>
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var fieldType = parameters.GetOptional<string?>("fieldType");
        var fieldArgument = parameters.GetOptional<string?>("fieldArgument");
        var paragraphIndex = parameters.GetOptional<int?>("paragraphIndex");
        var insertAtStart = parameters.GetOptional("insertAtStart", false);

        if (string.IsNullOrEmpty(fieldType))
            throw new ArgumentException("fieldType is required for insert_field operation");

        var document = context.Document;
        var builder = new DocumentBuilder(document);

        if (paragraphIndex.HasValue)
        {
            var paragraphs = document.GetChildNodes(NodeType.Paragraph, true);
            if (paragraphIndex.Value == -1)
            {
                builder.MoveToDocumentEnd();
            }
            else if (paragraphIndex.Value >= 0 && paragraphIndex.Value < paragraphs.Count)
            {
                if (paragraphs[paragraphIndex.Value] is WordParagraph targetPara)
                {
                    if (insertAtStart)
                    {
                        builder.MoveTo(targetPara);
                        if (targetPara.Runs.Count > 0)
                            builder.MoveTo(targetPara.Runs[0]);
                    }
                    else
                    {
                        builder.MoveTo(targetPara);
                        if (targetPara.Runs.Count > 0)
                            builder.MoveTo(targetPara.Runs[^1]);
                    }
                }
                else
                {
                    throw new InvalidOperationException($"Unable to find paragraph at index {paragraphIndex.Value}");
                }
            }
            else
            {
                throw new ArgumentException(
                    $"Paragraph index {paragraphIndex.Value} is out of range (document has {paragraphs.Count} paragraphs)");
            }
        }
        else
        {
            builder.MoveToDocumentEnd();
        }

        var code = fieldType.ToUpper();
        if (!string.IsNullOrEmpty(fieldArgument))
            code += " " + fieldArgument;

        var field = builder.InsertField(code);
        field.Update();

        MarkModified(context);

        var result = $"Field inserted successfully\nField type: {fieldType}\n";
        if (!string.IsNullOrEmpty(fieldArgument))
            result += $"Field argument: {fieldArgument}\n";
        result += $"Field code: {code}\n";

        try
        {
            var fieldResult = field.Result;
            if (!string.IsNullOrEmpty(fieldResult))
                result += $"Field result: {fieldResult}";
        }
        catch
        {
            // Ignore errors reading field result (some fields may not have results)
        }

        return result;
    }
}
