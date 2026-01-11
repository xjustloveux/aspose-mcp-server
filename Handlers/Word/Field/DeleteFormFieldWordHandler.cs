using Aspose.Words;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Word.Field;

/// <summary>
///     Handler for deleting form fields from Word documents.
/// </summary>
public class DeleteFormFieldWordHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "delete_form_field";

    /// <summary>
    ///     Deletes one or more form fields from the document.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Optional: fieldName (single field), fieldNames (array of fields)
    ///     If neither provided, deletes all form fields.
    /// </param>
    /// <returns>Success message with deletion count.</returns>
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var fieldName = parameters.GetOptional<string?>("fieldName");
        var fieldNames = parameters.GetOptional<string[]?>("fieldNames");

        var document = context.Document;
        var formFields = document.Range.FormFields;

        List<string> fieldsToDelete;
        if (fieldNames is { Length: > 0 })
            fieldsToDelete = fieldNames.Where(f => !string.IsNullOrEmpty(f)).ToList();
        else if (!string.IsNullOrEmpty(fieldName))
            fieldsToDelete = [fieldName];
        else
            fieldsToDelete = formFields.Select(f => f.Name).ToList();

        var deletedCount = 0;
        foreach (var name in fieldsToDelete)
        {
            var field = formFields[name];
            if (field != null)
            {
                field.Remove();
                deletedCount++;
            }
        }

        MarkModified(context);
        return Success($"Deleted {deletedCount} form field(s)");
    }
}
