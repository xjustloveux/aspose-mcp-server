using System.Text.Json;
using System.Text.Json.Nodes;
using Aspose.Words;
using Aspose.Words.Fields;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools.Word;

/// <summary>
///     Unified tool for field operations in Word documents
///     Merges: WordInsertFieldTool, WordEditFieldTool, WordDeleteFieldTool, WordUpdateFieldTool,
///     WordGetFieldsTool, WordGetFieldDetailTool, WordAddFormFieldTool, WordEditFormFieldTool,
///     WordDeleteFormFieldTool, WordGetFormFieldsTool
/// </summary>
public class WordFieldTool : IAsposeTool
{
    /// <summary>
    ///     Gets the description of the tool and its usage examples
    /// </summary>
    public string Description =>
        @"Manage fields and form fields in Word documents. Supports 11 operations: insert_field, edit_field, delete_field, update_field, update_all, get_fields, get_field_detail, add_form_field, edit_form_field, delete_form_field, get_form_fields.

Usage examples:
- Insert field: word_field(operation='insert_field', path='doc.docx', fieldType='DATE', paragraphIndex=0)
- Edit field: word_field(operation='edit_field', path='doc.docx', fieldIndex=0, fieldArgument='yyyy-MM-dd')
- Delete field: word_field(operation='delete_field', path='doc.docx', fieldIndex=0)
- Update field: word_field(operation='update_field', path='doc.docx', fieldIndex=0)
- Update all fields: word_field(operation='update_all', path='doc.docx') or word_field(operation='update_field', path='doc.docx', updateAll=true)
- Get fields: word_field(operation='get_fields', path='doc.docx')
- Add form field: word_field(operation='add_form_field', path='doc.docx', formFieldType='TextInput', fieldName='name', paragraphIndex=0)";

    /// <summary>
    ///     Gets the JSON schema defining the input parameters for the tool
    /// </summary>
    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            operation = new
            {
                type = "string",
                description = @"Operation to perform.
- 'insert_field': Insert a field (required params: path, fieldType)
- 'edit_field': Edit a field (required params: path, fieldIndex)
- 'delete_field': Delete a field (required params: path, fieldIndex)
- 'update_field': Update a field (required params: path, fieldIndex optional if updateAll=true)
- 'update_all': Alias for update_field with updateAll=true (required params: path)
- 'get_fields': Get all fields (required params: path)
- 'get_field_detail': Get field details (required params: path, fieldIndex)
- 'add_form_field': Add a form field (required params: path, formFieldType, fieldName, options for DropDown type)
- 'edit_form_field': Edit a form field (required params: path, fieldName)
- 'delete_form_field': Delete a form field (required params: path, fieldName)
- 'get_form_fields': Get all form fields (required params: path)",
                @enum = new[]
                {
                    "insert_field", "edit_field", "delete_field", "update_field", "update_all", "get_fields",
                    "get_field_detail", "add_form_field", "edit_form_field", "delete_form_field", "get_form_fields"
                }
            },
            path = new
            {
                type = "string",
                description = "Document file path (required for all operations)"
            },
            outputPath = new
            {
                type = "string",
                description = "Output file path (if not provided, overwrites input, for write operations)"
            },
            // Field operations
            fieldType = new
            {
                type = "string",
                description =
                    "Field type: DATE, TIME, PAGE, NUMPAGES, AUTHOR, FILENAME, etc. (required for insert_field operation)",
                @enum = new[]
                {
                    "DATE", "TIME", "PAGE", "NUMPAGES", "AUTHOR", "FILENAME", "TITLE", "SUBJECT",
                    "COMPANY", "CREATEDATE", "SAVEDATE", "PRINTDATE", "USERNAME", "DOCPROPERTY",
                    "HYPERLINK", "REF", "MERGEFIELD", "IF", "SEQ"
                }
            },
            fieldArgument = new
            {
                type = "string",
                description = "Field argument (optional, for insert_field operation)"
            },
            paragraphIndex = new
            {
                type = "number",
                description =
                    "Paragraph index to insert into (0-based, optional, for insert_field operation). Valid range: 0 to (total paragraphs - 1), or -1 for document end (last paragraph). When not specified, inserts at document end."
            },
            insertAtStart = new
            {
                type = "boolean",
                description =
                    "Insert at the start of the paragraph (optional, default: false, for insert_field operation)"
            },
            fieldIndex = new
            {
                type = "number",
                description = "Field index (0-based, required for edit_field/delete_field/get_field_detail operations)"
            },
            fieldCode = new
            {
                type = "string",
                description = "New field code (optional, for edit_field operation)"
            },
            lockField = new
            {
                type = "boolean",
                description = "Lock the field (optional, for edit_field operation)"
            },
            unlockField = new
            {
                type = "boolean",
                description = "Unlock the field (optional, for edit_field operation)"
            },
            updateField = new
            {
                type = "boolean",
                description = "Update the field after editing (optional, default: true, for edit_field operation)"
            },
            keepResult = new
            {
                type = "boolean",
                description =
                    "Keep the field result text after deletion (optional, default: false, for delete_field operation)"
            },
            updateAll = new
            {
                type = "boolean",
                description =
                    "Update all fields (optional, default: false if fieldIndex provided, for update_field operation)"
            },
            includeCode = new
            {
                type = "boolean",
                description = "Include field code in results (optional, default: true, for get_fields operation)"
            },
            includeResult = new
            {
                type = "boolean",
                description = "Include field result in results (optional, default: true, for get_fields operation)"
            },
            // Form field operations
            formFieldType = new
            {
                type = "string",
                description = "Form field type: TextInput, CheckBox, DropDown (required for add_form_field operation)",
                @enum = new[] { "TextInput", "CheckBox", "DropDown" }
            },
            fieldName = new
            {
                type = "string",
                description = "Field name (required for form field operations)"
            },
            defaultValue = new
            {
                type = "string",
                description = "Default value (optional, for add_form_field/edit_form_field operations)"
            },
            options = new
            {
                type = "array",
                items = new { type = "string" },
                description = "Options for dropdown (required for DropDown type, for add_form_field operation)"
            },
            checkedValue = new
            {
                type = "boolean",
                description =
                    "Checked state (optional, for CheckBox type, for add_form_field/edit_form_field operations)"
            },
            value = new
            {
                type = "string",
                description = "New value (optional, for TextInput type, for edit_form_field operation)"
            },
            selectedIndex = new
            {
                type = "number",
                description = "Selected option index (optional, for DropDown type, for edit_form_field operation)"
            },
            fieldNames = new
            {
                type = "array",
                items = new { type = "string" },
                description = "Array of form field names to delete (optional, for delete_form_field operation)"
            }
        },
        required = new[] { "operation", "path" }
    };

    /// <summary>
    ///     Executes the tool operation with the provided JSON arguments
    /// </summary>
    /// <param name="arguments">JSON arguments object containing operation parameters</param>
    /// <returns>Result message as a string</returns>
    /// <exception cref="ArgumentException">Thrown when operation is unknown or required parameters are missing.</exception>
    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var operation = ArgumentHelper.GetString(arguments, "operation");
        var path = ArgumentHelper.GetAndValidatePath(arguments);
        SecurityHelper.ValidateFilePath(path, allowAbsolutePaths: true);
        var outputPath = ArgumentHelper.GetStringNullable(arguments, "outputPath") ?? path;
        SecurityHelper.ValidateFilePath(outputPath, "outputPath", true);

        // Ensure output directory exists for write operations
        var outputDir = Path.GetDirectoryName(outputPath);
        if (!string.IsNullOrEmpty(outputDir))
            Directory.CreateDirectory(outputDir);

        // Normalize operation name: update_all is an alias for update_field with updateAll=true
        if (operation == "update_all")
        {
            operation = "update_field";
            // Ensure updateAll is set to true if not already specified
            if (arguments != null && !arguments.ContainsKey("updateAll")) arguments["updateAll"] = true;
        }

        return operation switch
        {
            "insert_field" => await InsertFieldAsync(path, outputPath, arguments),
            "edit_field" => await EditFieldAsync(path, outputPath, arguments),
            "delete_field" => await DeleteFieldAsync(path, outputPath, arguments),
            "update_field" => await UpdateFieldAsync(path, outputPath, arguments),
            "get_fields" => await GetFieldsAsync(path, arguments),
            "get_field_detail" => await GetFieldDetailAsync(path, arguments),
            "add_form_field" => await AddFormFieldAsync(path, outputPath, arguments),
            "edit_form_field" => await EditFormFieldAsync(path, outputPath, arguments),
            "delete_form_field" => await DeleteFormFieldAsync(path, outputPath, arguments),
            "get_form_fields" => await GetFormFieldsAsync(path),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Inserts a field into the document
    /// </summary>
    /// <param name="path">Word document file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="arguments">JSON arguments containing fieldType, optional fieldArgument, paragraphIndex, runIndex</param>
    /// <returns>Success message</returns>
    private Task<string> InsertFieldAsync(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var fieldType = ArgumentHelper.GetString(arguments, "fieldType").ToUpper();
            var fieldArgument = ArgumentHelper.GetString(arguments, "fieldArgument", "");
            var paragraphIndex = ArgumentHelper.GetIntNullable(arguments, "paragraphIndex");
            var insertAtStart = ArgumentHelper.GetBool(arguments, "insertAtStart", false);

            var doc = new Document(path);
            var builder = new DocumentBuilder(doc);

            if (paragraphIndex.HasValue)
            {
                var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
                if (paragraphIndex.Value == -1)
                {
                    builder.MoveToDocumentEnd();
                }
                else if (paragraphIndex.Value >= 0 && paragraphIndex.Value < paragraphs.Count)
                {
                    if (paragraphs[paragraphIndex.Value] is Paragraph targetPara)
                    {
                        if (insertAtStart)
                        {
                            builder.MoveTo(targetPara);
                            if (targetPara.Runs.Count > 0)
                                builder.MoveTo(targetPara.Runs[0]);
                            else
                                // Empty paragraph - just move to paragraph
                                builder.MoveTo(targetPara);
                        }
                        else
                        {
                            builder.MoveTo(targetPara);

                            if (targetPara.Runs.Count > 0)
                            {
                                var lastRun = targetPara.Runs[^1];
                                builder.MoveTo(lastRun);
                                // Move to end of the last run
                                try
                                {
                                    builder.MoveToParagraph(paragraphIndex.Value, lastRun.Text.Length);
                                }
                                catch (Exception ex)
                                {
                                    // If that fails, just stay at the run position
                                    // The field will be inserted after the run
                                    Console.Error.WriteLine(
                                        $"[WARN] Failed to move to paragraph {paragraphIndex.Value}, staying at run position: {ex.Message}");
                                }
                            }
                            else
                            {
                                // Empty paragraph - just move to paragraph
                                builder.MoveTo(targetPara);
                            }
                        }
                    }
                    else
                    {
                        throw new InvalidOperationException(
                            $"Unable to find paragraph at index {paragraphIndex.Value}");
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

            var fieldCode = fieldType;
            if (!string.IsNullOrEmpty(fieldArgument)) fieldCode += " " + fieldArgument;

            Field field;
            try
            {
                field = builder.InsertField(fieldCode);
                field.Update();
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Unable to insert field '{fieldCode}': {ex.Message}", ex);
            }

            doc.Save(outputPath);

            var result = "Field inserted successfully\n";
            result += $"Field type: {fieldType}\n";
            if (!string.IsNullOrEmpty(fieldArgument))
                result += $"Field argument: {fieldArgument}\n";
            result += $"Field code: {fieldCode}\n";

            try
            {
                var fieldResult = field.Result;
                if (!string.IsNullOrEmpty(fieldResult))
                    result += $"Field result: {fieldResult}\n";
            }
            catch (Exception ex)
            {
                // Field.Result may fail for some field types, but this is not critical
                Console.Error.WriteLine(
                    $"[WARN] Failed to get field result (may be normal for some field types): {ex.Message}");
                // Continue without the result information
            }

            if (paragraphIndex.HasValue)
            {
                if (paragraphIndex.Value == -1)
                    result += "Insert position: end of document (last paragraph)\n";
                else
                    result += $"Insert position: paragraph #{paragraphIndex.Value}\n";
            }
            else
            {
                result += "Insert position: end of document\n";
            }

            result += $"Output: {outputPath}";

            return result;
        });
    }

    /// <summary>
    ///     Edits an existing field
    /// </summary>
    /// <param name="path">Word document file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="arguments">JSON arguments containing fieldIndex, optional fieldType, fieldArgument</param>
    /// <returns>Success message</returns>
    private Task<string> EditFieldAsync(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var fieldIndex = ArgumentHelper.GetInt(arguments, "fieldIndex");
            var fieldCode = ArgumentHelper.GetStringNullable(arguments, "fieldCode");
            var lockField = ArgumentHelper.GetBoolNullable(arguments, "lockField");
            var unlockField = ArgumentHelper.GetBoolNullable(arguments, "unlockField");
            var updateField = ArgumentHelper.GetBool(arguments, "updateField", true);

            var doc = new Document(path);
            var fields = doc.Range.Fields.ToList();

            if (fieldIndex < 0 || fieldIndex >= fields.Count)
                throw new ArgumentException(
                    $"Field index {fieldIndex} is out of range (document has {fields.Count} fields)");

            var field = fields[fieldIndex];
            var oldFieldCode = field.GetFieldCode();
            var oldResult = field.Result ?? "";
            var oldLocked = field.IsLocked;

            var changes = new List<string>();

            if (!string.IsNullOrEmpty(fieldCode))
                try
                {
                    var fieldStart = field.Start;
                    var fieldEnd = field.End;

                    if (fieldStart != null && fieldEnd != null)
                    {
                        var builder = new DocumentBuilder(doc);
                        builder.MoveTo(fieldStart);

                        var currentNode = fieldStart.NextSibling;
                        while (currentNode != null && currentNode != fieldEnd)
                        {
                            var nextNode = currentNode.NextSibling;
                            if (currentNode.NodeType != NodeType.FieldSeparator &&
                                currentNode.NodeType != NodeType.FieldEnd) currentNode.Remove();
                            currentNode = nextNode;
                        }

                        builder.MoveTo(fieldStart);
                        builder.Write(fieldCode);

                        changes.Add($"Field code updated: {oldFieldCode} -> {fieldCode}");
                    }
                }
                catch (Exception ex)
                {
                    throw new InvalidOperationException($"Unable to update field code: {ex.Message}", ex);
                }

            if (lockField.HasValue && lockField.Value)
            {
                field.IsLocked = true;
                changes.Add("Field locked");
            }
            else if (unlockField.HasValue && unlockField.Value)
            {
                field.IsLocked = false;
                changes.Add("Field unlocked");
            }

            if (updateField)
            {
                field.Update();
                doc.UpdateFields();
            }

            doc.Save(outputPath);

            var result = $"Field #{fieldIndex} edited successfully\n";
            result += $"Original field code: {oldFieldCode}\n";
            if (!string.IsNullOrEmpty(fieldCode)) result += $"New field code: {fieldCode}\n";
            result += $"Original result: {oldResult}\n";
            result += $"Original lock status: {(oldLocked ? "Locked" : "Unlocked")}\n";
            if (changes.Count > 0) result += $"Changes: {string.Join(", ", changes)}\n";
            result += $"Output: {outputPath}";

            return result;
        });
    }

    /// <summary>
    ///     Deletes a field from the document
    /// </summary>
    /// <param name="path">Word document file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="arguments">JSON arguments containing fieldIndex</param>
    /// <returns>Success message</returns>
    private Task<string> DeleteFieldAsync(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var fieldIndex = ArgumentHelper.GetInt(arguments, "fieldIndex");
            var keepResult = ArgumentHelper.GetBool(arguments, "keepResult", false);

            var doc = new Document(path);
            var fields = doc.Range.Fields.ToList();

            if (fieldIndex < 0 || fieldIndex >= fields.Count)
                throw new ArgumentException(
                    $"Field index {fieldIndex} is out of range (document has {fields.Count} fields)");

            var field = fields[fieldIndex];
            var fieldType = field.Type.ToString();
            var fieldCode = field.GetFieldCode();
            var fieldResult = field.Result ?? "";

            try
            {
                if (keepResult)
                    // Unlink converts the field to plain text (keeps Result, removes field markers)
                    field.Unlink();
                else
                    // Remove deletes the entire field including its content
                    field.Remove();

                doc.Save(outputPath);

                var remainingFields = doc.Range.Fields.Count;

                var result = $"Field #{fieldIndex} deleted successfully\n";
                result += $"Type: {fieldType}\n";
                result += $"Code: {fieldCode}\n";
                if (!string.IsNullOrEmpty(fieldResult)) result += $"Result: {fieldResult}\n";
                result += $"Keep result text: {(keepResult ? "Yes (converted to plain text)" : "No")}\n";
                result += $"Remaining fields in document: {remainingFields}\n";
                result += $"Output: {outputPath}";

                return result;
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Unable to delete field: {ex.Message}", ex);
            }
        });
    }

    /// <summary>
    ///     Updates field values
    /// </summary>
    /// <param name="path">Word document file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="arguments">JSON arguments containing optional fieldIndex (if null, updates all)</param>
    /// <returns>Success message with update count</returns>
    private Task<string> UpdateFieldAsync(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var fieldIndex = ArgumentHelper.GetIntNullable(arguments, "fieldIndex");
            var fieldTypeFilter = ArgumentHelper.GetStringNullable(arguments, "fieldType");
            var updateAll = ArgumentHelper.GetBool(arguments, "updateAll", !fieldIndex.HasValue);

            var doc = new Document(path);
            var fields = doc.Range.Fields.ToList();
            var lockedFields = new List<string>();

            if (fieldIndex.HasValue && !updateAll)
            {
                if (fieldIndex.Value < 0 || fieldIndex.Value >= fields.Count)
                    throw new ArgumentException(
                        $"Field index {fieldIndex.Value} is out of range (document has {fields.Count} fields)");

                var field = fields[fieldIndex.Value];

                if (field.IsLocked)
                    return $"Warning: Field #{fieldIndex.Value} is locked and cannot be updated.\n" +
                           $"Type: {field.Type}\n" +
                           $"Code: {field.GetFieldCode()}\n" +
                           "Use edit_field with unlockField=true to unlock it first.";

                try
                {
                    var oldResult = field.Result ?? "";
                    field.Update();
                    var newResult = field.Result ?? "";

                    doc.Save(outputPath);

                    var singleResult = $"Field #{fieldIndex.Value} updated successfully\n";
                    singleResult += $"Type: {field.Type}\n";
                    singleResult += $"Code: {field.GetFieldCode()}\n";
                    singleResult += $"Old result: {oldResult}\n";
                    singleResult += $"New result: {newResult}\n";
                    singleResult += $"Output: {outputPath}";

                    return singleResult;
                }
                catch (Exception ex)
                {
                    throw new InvalidOperationException($"Unable to update field #{fieldIndex.Value}: {ex.Message}",
                        ex);
                }
            }

            // First, collect locked fields for warning
            foreach (var field in fields)
            {
                if (!string.IsNullOrEmpty(fieldTypeFilter))
                {
                    var filterType = fieldTypeFilter.ToUpper();
                    var fieldTypeName = field.Type.ToString().ToUpper();
                    if (!fieldTypeName.Contains(filterType)) continue;
                }

                if (field.IsLocked)
                    lockedFields.Add($"Field {field.Type} (index {fields.IndexOf(field)})");
            }

            // Use doc.UpdateFields() for bulk update - more efficient than individual updates
            doc.UpdateFields();
            var updatedCount = fields.Count - lockedFields.Count;

            doc.Save(outputPath);

            var result = $"Successfully updated {updatedCount} field(s)\n";
            if (!string.IsNullOrEmpty(fieldTypeFilter)) result += $"Filter type: {fieldTypeFilter}\n";

            if (lockedFields.Count > 0)
            {
                result += $"\nLocked fields (skipped, {lockedFields.Count}):\n";
                foreach (var locked in lockedFields) result += $"  - {locked}\n";
            }

            result += $"Output: {outputPath}";

            return result;
        });
    }

    /// <summary>
    ///     Gets all fields from the document
    /// </summary>
    /// <param name="path">Word document file path</param>
    /// <param name="arguments">JSON arguments (no specific parameters required)</param>
    /// <returns>JSON formatted string with all fields for better LLM processing</returns>
    private Task<string> GetFieldsAsync(string path, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var fieldTypeFilter = ArgumentHelper.GetStringNullable(arguments, "fieldType");
            var includeCode = ArgumentHelper.GetBool(arguments, "includeCode", true);
            var includeResult = ArgumentHelper.GetBool(arguments, "includeResult", true);

            var doc = new Document(path);
            var fieldsList = new List<object>();
            var fieldIndex = 0;

            foreach (var field in doc.Range.Fields)
            {
                if (!string.IsNullOrEmpty(fieldTypeFilter))
                {
                    var filterType = fieldTypeFilter.ToUpper();
                    var fieldTypeName = field.Type.ToString().ToUpper();
                    if (!fieldTypeName.Contains(filterType) && filterType != "ALL") continue;
                }

                string? extraInfo = null;
                if (field is FieldHyperlink hyperlinkField)
                    extraInfo = $"Address: {hyperlinkField.Address ?? ""}, ScreenTip: {hyperlinkField.ScreenTip ?? ""}";
                else if (field is FieldRef refField)
                    extraInfo = $"Bookmark: {refField.BookmarkName ?? ""}";

                fieldsList.Add(new
                {
                    index = fieldIndex++,
                    type = field.Type.ToString(),
                    code = includeCode ? field.GetFieldCode() : null,
                    result = includeResult ? field.Result ?? "" : null,
                    isLocked = field.IsLocked,
                    isDirty = field.IsDirty,
                    extraInfo
                });
            }

            // Build statistics by type
            var statistics = fieldsList
                .GroupBy(f => ((dynamic)f).type as string)
                .OrderBy(g => g.Key)
                .Select(g => new { type = g.Key, count = g.Count() })
                .ToList();

            var result = new
            {
                count = fieldsList.Count,
                fields = fieldsList,
                statisticsByType = statistics
            };

            return JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
        });
    }

    /// <summary>
    ///     Gets detailed information about a specific field
    /// </summary>
    /// <param name="path">Word document file path</param>
    /// <param name="arguments">JSON arguments containing fieldIndex</param>
    /// <returns>JSON formatted string with field details for better LLM processing</returns>
    private Task<string> GetFieldDetailAsync(string path, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var fieldIndex = ArgumentHelper.GetInt(arguments, "fieldIndex");

            var doc = new Document(path);
            var fields = doc.Range.Fields.ToList();

            if (fieldIndex < 0 || fieldIndex >= fields.Count)
                throw new ArgumentException(
                    $"Field index {fieldIndex} is out of range (document has {fields.Count} fields)");

            var field = fields[fieldIndex];

            // Build extra info based on field type
            string? address = null;
            string? screenTip = null;
            string? bookmarkName = null;

            if (field is FieldHyperlink hyperlinkField)
            {
                address = hyperlinkField.Address;
                screenTip = hyperlinkField.ScreenTip;
            }
            else if (field is FieldRef refField)
            {
                bookmarkName = refField.BookmarkName;
            }

            var result = new
            {
                index = fieldIndex,
                type = field.Type.ToString(),
                typeCode = (int)field.Type,
                code = field.GetFieldCode(),
                result = field.Result,
                isLocked = field.IsLocked,
                isDirty = field.IsDirty,
                hyperlinkAddress = address,
                hyperlinkScreenTip = screenTip,
                bookmarkName
            };

            return JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
        });
    }

    /// <summary>
    ///     Adds a form field to the document
    /// </summary>
    /// <param name="path">Word document file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="arguments">JSON arguments containing fieldType, name, optional defaultValue, paragraphIndex</param>
    /// <returns>Success message</returns>
    private Task<string> AddFormFieldAsync(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var fieldType = ArgumentHelper.GetString(arguments, "formFieldType");
            var fieldName = ArgumentHelper.GetString(arguments, "fieldName");
            var defaultValue = ArgumentHelper.GetStringNullable(arguments, "defaultValue");
            var checkedValue = ArgumentHelper.GetBoolNullable(arguments, "checkedValue");

            var doc = new Document(path);
            var builder = new DocumentBuilder(doc);
            builder.MoveToDocumentEnd();

            switch (fieldType.ToLower())
            {
                case "textinput":
                    builder.InsertTextInput(fieldName, TextFormFieldType.Regular, "", defaultValue ?? "", 0);
                    break;

                case "checkbox":
                    builder.InsertCheckBox(fieldName, checkedValue ?? false, 0);
                    break;

                case "dropdown":
                    var optionsArray = ArgumentHelper.GetArray(arguments, "options");
                    if (optionsArray.Count == 0)
                        throw new ArgumentException("options array cannot be empty for DropDown type");
                    var options = optionsArray.Select(o => o?.GetValue<string>()).Where(o => !string.IsNullOrEmpty(o))
                        .ToArray();
                    builder.InsertComboBox(fieldName, options, 0);
                    break;

                default:
                    throw new ArgumentException(
                        $"Invalid fieldType: {fieldType}. Must be 'TextInput', 'CheckBox', or 'DropDown'");
            }

            doc.Save(outputPath);
            return $"{fieldType} field '{fieldName}' added: {outputPath}";
        });
    }

    /// <summary>
    ///     Edits an existing form field
    /// </summary>
    /// <param name="path">Word document file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="arguments">JSON arguments containing fieldIndex, optional name, defaultValue</param>
    /// <returns>Success message</returns>
    private Task<string> EditFormFieldAsync(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var fieldName = ArgumentHelper.GetString(arguments, "fieldName");
            var value = ArgumentHelper.GetStringNullable(arguments, "value");
            var checkedValue = ArgumentHelper.GetBoolNullable(arguments, "checkedValue");
            var selectedIndex = ArgumentHelper.GetIntNullable(arguments, "selectedIndex");

            var doc = new Document(path);
            var field = doc.Range.FormFields[fieldName];

            if (field == null) throw new ArgumentException($"Form field '{fieldName}' not found");

            if (field.Type == FieldType.FieldFormTextInput && value != null)
                field.Result = value;
            else if (field.Type == FieldType.FieldFormCheckBox && checkedValue.HasValue)
                field.Checked = checkedValue.Value;
            else if (field.Type == FieldType.FieldFormDropDown && selectedIndex.HasValue)
                if (selectedIndex.Value >= 0 && selectedIndex.Value < field.DropDownItems.Count)
                    field.DropDownSelectedIndex = selectedIndex.Value;

            doc.Save(outputPath);
            return $"Form field '{fieldName}' updated: {outputPath}";
        });
    }

    /// <summary>
    ///     Deletes a form field from the document
    /// </summary>
    /// <param name="path">Word document file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="arguments">JSON arguments containing fieldIndex</param>
    /// <returns>Success message</returns>
    private Task<string> DeleteFormFieldAsync(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var fieldName = ArgumentHelper.GetStringNullable(arguments, "fieldName");
            var fieldNamesArray = ArgumentHelper.GetArray(arguments, "fieldNames", false);

            var doc = new Document(path);
            var formFields = doc.Range.FormFields;

            List<string> fieldsToDelete;
            if (fieldNamesArray is { Count: > 0 })
                fieldsToDelete = fieldNamesArray.Select(f => f?.GetValue<string>()).Where(f => !string.IsNullOrEmpty(f))
                    .Select(f => f!).ToList();
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

            doc.Save(outputPath);
            return $"Deleted {deletedCount} form field(s): {outputPath}";
        });
    }

    /// <summary>
    ///     Gets all form fields from the document
    /// </summary>
    /// <param name="path">Word document file path</param>
    /// <returns>JSON formatted string with all form fields for better LLM processing</returns>
    private Task<string> GetFormFieldsAsync(string path)
    {
        return Task.Run(() =>
        {
            var doc = new Document(path);
            var formFields = doc.Range.FormFields.ToList();
            var formFieldsList = new List<object>();

            for (var i = 0; i < formFields.Count; i++)
            {
                var field = formFields[i];
                object fieldData = field.Type switch
                {
                    FieldType.FieldFormTextInput => new
                    {
                        index = i,
                        name = field.Name,
                        type = field.Type.ToString(),
                        value = field.Result
                    },
                    FieldType.FieldFormCheckBox => new
                    {
                        index = i,
                        name = field.Name,
                        type = field.Type.ToString(),
                        isChecked = field.Checked
                    },
                    FieldType.FieldFormDropDown => new
                    {
                        index = i,
                        name = field.Name,
                        type = field.Type.ToString(),
                        selectedIndex = field.DropDownSelectedIndex,
                        options = field.DropDownItems.ToList()
                    },
                    _ => new
                    {
                        index = i,
                        name = field.Name,
                        type = field.Type.ToString()
                    }
                };

                formFieldsList.Add(fieldData);
            }

            var result = new
            {
                count = formFields.Count,
                formFields = formFieldsList
            };

            return JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
        });
    }
}