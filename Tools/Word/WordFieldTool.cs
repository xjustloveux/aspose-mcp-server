using System.Text;
using System.Text.Json.Nodes;
using Aspose.Words;
using Aspose.Words.Fields;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools;

/// <summary>
/// Unified tool for field operations in Word documents
/// Merges: WordInsertFieldTool, WordEditFieldTool, WordDeleteFieldTool, WordUpdateFieldTool,
/// WordGetFieldsTool, WordGetFieldDetailTool, WordAddFormFieldTool, WordEditFormFieldTool,
/// WordDeleteFormFieldTool, WordGetFormFieldsTool
/// </summary>
public class WordFieldTool : IAsposeTool
{
    public string Description => @"Manage fields and form fields in Word documents. Supports 10 operations: insert_field, edit_field, delete_field, update_field (or update_all), get_fields, get_field_detail, add_form_field, edit_form_field, delete_form_field, get_form_fields.

Usage examples:
- Insert field: word_field(operation='insert_field', path='doc.docx', fieldType='DATE', paragraphIndex=0)
- Edit field: word_field(operation='edit_field', path='doc.docx', fieldIndex=0, fieldArgument='yyyy-MM-dd')
- Delete field: word_field(operation='delete_field', path='doc.docx', fieldIndex=0)
- Update field: word_field(operation='update_field', path='doc.docx', fieldIndex=0)
- Update all fields: word_field(operation='update_all', path='doc.docx') or word_field(operation='update_field', path='doc.docx', updateAll=true)
- Get fields: word_field(operation='get_fields', path='doc.docx')
- Add form field: word_field(operation='add_form_field', path='doc.docx', formFieldType='TextInput', fieldName='name', paragraphIndex=0)";

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
- 'add_form_field': Add a form field (required params: path, formFieldType, fieldName)
- 'edit_form_field': Edit a form field (required params: path, fieldName)
- 'delete_form_field': Delete a form field (required params: path, fieldName)
- 'get_form_fields': Get all form fields (required params: path)",
                @enum = new[] { "insert_field", "edit_field", "delete_field", "update_field", "get_fields", "get_field_detail", "add_form_field", "edit_form_field", "delete_form_field", "get_form_fields" }
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
                description = "Field type: DATE, TIME, PAGE, NUMPAGES, AUTHOR, FILENAME, etc. (required for insert_field operation)",
                @enum = new[] { "DATE", "TIME", "PAGE", "NUMPAGES", "AUTHOR", "FILENAME", "TITLE", "SUBJECT", 
                               "COMPANY", "CREATEDATE", "SAVEDATE", "PRINTDATE", "USERNAME", "DOCPROPERTY", 
                               "HYPERLINK", "REF", "MERGEFIELD", "IF", "SEQ" }
            },
            fieldArgument = new
            {
                type = "string",
                description = "Field argument (optional, for insert_field operation)"
            },
            paragraphIndex = new
            {
                type = "number",
                description = "Paragraph index to insert into (0-based, optional, for insert_field operation). Valid range: 0 to (total paragraphs - 1), or -1 for document start. When not specified, inserts at document end."
            },
            insertAtStart = new
            {
                type = "boolean",
                description = "Insert at the start of the paragraph (optional, default: false, for insert_field operation)"
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
                description = "Keep the field result text after deletion (optional, default: false, for delete_field operation)"
            },
            updateAll = new
            {
                type = "boolean",
                description = "Update all fields (optional, default: false if fieldIndex provided, for update_field operation)"
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
                description = "Checked state (optional, for CheckBox type, for add_form_field/edit_form_field operations)"
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

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var operation = arguments?["operation"]?.GetValue<string>() ?? throw new ArgumentException("operation is required");
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        SecurityHelper.ValidateFilePath(path, "path");
        var outputPath = arguments?["outputPath"]?.GetValue<string>() ?? path;
        SecurityHelper.ValidateFilePath(outputPath, "outputPath");

        // Normalize operation name: update_all is an alias for update_field with updateAll=true
        if (operation == "update_all")
        {
            operation = "update_field";
            // Ensure updateAll is set to true if not already specified
            if (arguments != null && !arguments.ContainsKey("updateAll"))
            {
                arguments["updateAll"] = true;
            }
        }
        
        return operation switch
        {
            "insert_field" => await InsertFieldAsync(arguments, path, outputPath),
            "edit_field" => await EditFieldAsync(arguments, path, outputPath),
            "delete_field" => await DeleteFieldAsync(arguments, path, outputPath),
            "update_field" => await UpdateFieldAsync(arguments, path, outputPath),
            "get_fields" => await GetFieldsAsync(arguments, path),
            "get_field_detail" => await GetFieldDetailAsync(arguments, path),
            "add_form_field" => await AddFormFieldAsync(arguments, path, outputPath),
            "edit_form_field" => await EditFormFieldAsync(arguments, path, outputPath),
            "delete_form_field" => await DeleteFormFieldAsync(arguments, path, outputPath),
            "get_form_fields" => await GetFormFieldsAsync(arguments, path),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    private async Task<string> InsertFieldAsync(JsonObject? arguments, string path, string outputPath)
    {
        var fieldType = arguments?["fieldType"]?.GetValue<string>()?.ToUpper() ?? throw new ArgumentException("fieldType is required");
        var fieldArgument = arguments?["fieldArgument"]?.GetValue<string>() ?? "";
        var paragraphIndex = arguments?["paragraphIndex"]?.GetValue<int?>();
        var insertAtStart = arguments?["insertAtStart"]?.GetValue<bool>() ?? false;

        var doc = new Document(path);
        var builder = new DocumentBuilder(doc);

        if (paragraphIndex.HasValue)
        {
            var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
            if (paragraphIndex.Value == -1)
            {
                // paragraphIndex=-1 means document end, not document start
                // Move to the last paragraph in the document body
                var lastSection = doc.LastSection;
                var bodyParagraphs = lastSection.Body.GetChildNodes(NodeType.Paragraph, false);
                if (bodyParagraphs.Count > 0)
                {
                    var lastPara = bodyParagraphs[bodyParagraphs.Count - 1] as Paragraph;
                    if (lastPara != null)
                    {
                        builder.MoveTo(lastPara);
                        if (insertAtStart)
                        {
                            // Move to start of last paragraph
                            if (lastPara.Runs.Count > 0)
                            {
                                builder.MoveTo(lastPara.Runs[0]);
                            }
                        }
                        else
                        {
                            // Move to end of last paragraph
                            if (lastPara.Runs.Count > 0)
                            {
                                var lastRun = lastPara.Runs[lastPara.Runs.Count - 1];
                                builder.MoveTo(lastRun);
                            }
                        }
                    }
                    else
                    {
                        builder.MoveToDocumentEnd();
                    }
                }
                else
                {
                    builder.MoveToDocumentEnd();
                }
            }
            else if (paragraphIndex.Value >= 0 && paragraphIndex.Value < paragraphs.Count)
            {
                var targetPara = paragraphs[paragraphIndex.Value] as Paragraph;
                if (targetPara != null)
                {
                    if (insertAtStart)
                    {
                        builder.MoveTo(targetPara);
                        if (targetPara.Runs.Count > 0)
                        {
                            builder.MoveTo(targetPara.Runs[0]);
                        }
                        else
                        {
                            // Empty paragraph - just move to paragraph
                            builder.MoveTo(targetPara);
                        }
                    }
                    else
                    {
                        builder.MoveTo(targetPara);
                        
                        if (targetPara.Runs.Count > 0)
                        {
                            var lastRun = targetPara.Runs[targetPara.Runs.Count - 1];
                            builder.MoveTo(lastRun);
                            // Move to end of the last run
                            try
                            {
                                builder.MoveToParagraph(paragraphIndex.Value, lastRun.Text.Length);
                            }
                            catch
                            {
                                // If that fails, just stay at the run position
                                // The field will be inserted after the run
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
                    throw new InvalidOperationException($"無法找到索引 {paragraphIndex.Value} 的段落");
                }
            }
            else
            {
                throw new ArgumentException($"段落索引 {paragraphIndex.Value} 超出範圍 (文檔共有 {paragraphs.Count} 個段落)");
            }
        }
        else
        {
            builder.MoveToDocumentEnd();
        }

        string fieldCode = fieldType;
        if (!string.IsNullOrEmpty(fieldArgument))
        {
            fieldCode += " " + fieldArgument;
        }

        Field field;
        try
        {
            field = builder.InsertField(fieldCode);
            field.Update();
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"無法插入欄位 '{fieldCode}': {ex.Message}", ex);
        }

        doc.Save(outputPath);

        var result = $"成功插入欄位\n";
        result += $"欄位類型: {fieldType}\n";
        if (!string.IsNullOrEmpty(fieldArgument))
            result += $"欄位參數: {fieldArgument}\n";
        result += $"欄位代碼: {fieldCode}\n";
        
        try
        {
            var fieldResult = field.Result;
            if (!string.IsNullOrEmpty(fieldResult))
                result += $"欄位結果: {fieldResult}\n";
        }
        catch { }
        
        if (paragraphIndex.HasValue)
        {
            if (paragraphIndex.Value == -1)
                result += "插入位置: 文檔開頭\n";
            else
                result += $"插入位置: 段落 #{paragraphIndex.Value}\n";
        }
        else
        {
            result += "插入位置: 文檔末尾\n";
        }
        
        result += $"輸出: {outputPath}";

        return await Task.FromResult(result);
    }

    private async Task<string> EditFieldAsync(JsonObject? arguments, string path, string outputPath)
    {
        var fieldIndex = arguments?["fieldIndex"]?.GetValue<int>() ?? throw new ArgumentException("fieldIndex is required");
        var fieldCode = arguments?["fieldCode"]?.GetValue<string>();
        var lockField = arguments?["lockField"]?.GetValue<bool?>();
        var unlockField = arguments?["unlockField"]?.GetValue<bool?>();
        var updateField = arguments?["updateField"]?.GetValue<bool>() ?? true;

        var doc = new Document(path);
        var fields = doc.Range.Fields.ToList();
        
        if (fieldIndex < 0 || fieldIndex >= fields.Count)
        {
            throw new ArgumentException($"欄位索引 {fieldIndex} 超出範圍 (文檔共有 {fields.Count} 個欄位)");
        }

        var field = fields[fieldIndex];
        var oldFieldCode = field.GetFieldCode();
        var oldResult = field.Result ?? "";
        var oldLocked = field.IsLocked;
        
        var changes = new List<string>();
        
        if (!string.IsNullOrEmpty(fieldCode))
        {
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
                        if (currentNode.NodeType != NodeType.FieldSeparator && currentNode.NodeType != NodeType.FieldEnd)
                        {
                            currentNode.Remove();
                        }
                        currentNode = nextNode;
                    }
                    
                    builder.MoveTo(fieldStart);
                    builder.Write(fieldCode);
                    
                    changes.Add($"欄位代碼已更新: {oldFieldCode} -> {fieldCode}");
                }
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"無法更新欄位代碼: {ex.Message}", ex);
            }
        }
        
        if (lockField.HasValue && lockField.Value)
        {
            field.IsLocked = true;
            changes.Add("欄位已鎖定");
        }
        else if (unlockField.HasValue && unlockField.Value)
        {
            field.IsLocked = false;
            changes.Add("欄位已解鎖");
        }
        
        if (updateField)
        {
            field.Update();
            doc.UpdateFields();
        }
        
        doc.Save(outputPath);
        
        var result = $"成功編輯欄位 #{fieldIndex}\n";
        result += $"原欄位代碼: {oldFieldCode}\n";
        if (!string.IsNullOrEmpty(fieldCode))
        {
            result += $"新欄位代碼: {fieldCode}\n";
        }
        result += $"原結果: {oldResult}\n";
        result += $"原鎖定狀態: {(oldLocked ? "已鎖定" : "未鎖定")}\n";
        if (changes.Count > 0)
        {
            result += $"變更: {string.Join(", ", changes)}\n";
        }
        result += $"輸出: {outputPath}";
        
        return await Task.FromResult(result);
    }

    private async Task<string> DeleteFieldAsync(JsonObject? arguments, string path, string outputPath)
    {
        var fieldIndex = arguments?["fieldIndex"]?.GetValue<int>() ?? throw new ArgumentException("fieldIndex is required");
        var keepResult = arguments?["keepResult"]?.GetValue<bool>() ?? false;

        var doc = new Document(path);
        var fields = doc.Range.Fields.ToList();
        
        if (fieldIndex < 0 || fieldIndex >= fields.Count)
        {
            throw new ArgumentException($"功能變數索引 {fieldIndex} 超出範圍 (文檔共有 {fields.Count} 個功能變數)");
        }
        
        var field = fields[fieldIndex];
        var fieldType = field.Type.ToString();
        var fieldCode = field.GetFieldCode();
        var fieldResult = field.Result ?? "";
        
        try
        {
            if (keepResult)
            {
                var fieldStart = field.Start;
                var fieldEnd = field.End;
                
                if (fieldStart != null && fieldEnd != null)
                {
                    var nodesToRemove = new List<Node>();
                    var currentNode = (Node)fieldStart;
                    var endNode = (Node)fieldEnd;
                    
                    while (currentNode != null && currentNode != endNode.NextSibling)
                    {
                        if (currentNode.NodeType == NodeType.FieldStart || 
                            currentNode.NodeType == NodeType.FieldSeparator)
                        {
                            nodesToRemove.Add(currentNode);
                        }
                        currentNode = currentNode.NextSibling;
                        
                        if (currentNode == endNode)
                        {
                            break;
                        }
                    }
                    
                    foreach (var node in nodesToRemove)
                    {
                        node.Remove();
                    }
                    
                    endNode.Remove();
                }
            }
            else
            {
                var fieldStart = field.Start;
                var fieldEnd = field.End;
                
                if (fieldStart != null && fieldEnd != null)
                {
                    var nodesToRemove = new List<Node>();
                    var currentNode = (Node)fieldStart;
                    var endNode = (Node)fieldEnd;
                    
                    while (currentNode != null)
                    {
                        var nextNode = currentNode.NextSibling;
                        nodesToRemove.Add(currentNode);
                        
                        if (currentNode == endNode)
                        {
                            break;
                        }
                        
                        currentNode = nextNode;
                    }
                    
                    foreach (var node in nodesToRemove)
                    {
                        node.Remove();
                    }
                }
            }
            
            doc.Save(outputPath);
            
            var remainingFields = doc.Range.Fields.Count;
            
            var result = $"成功刪除功能變數 #{fieldIndex}\n";
            result += $"類型: {fieldType}\n";
            result += $"代碼: {fieldCode}\n";
            if (!string.IsNullOrEmpty(fieldResult))
            {
                result += $"結果: {fieldResult}\n";
            }
            result += $"保留結果文字: {(keepResult ? "是" : "否")}\n";
            result += $"文檔剩餘功能變數數: {remainingFields}\n";
            result += $"輸出: {outputPath}";
            
            return await Task.FromResult(result);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"無法刪除功能變數: {ex.Message}", ex);
        }
    }

    private async Task<string> UpdateFieldAsync(JsonObject? arguments, string path, string outputPath)
    {
        var fieldIndex = arguments?["fieldIndex"]?.GetValue<int?>();
        var fieldTypeFilter = arguments?["fieldType"]?.GetValue<string>();
        var updateAll = arguments?["updateAll"]?.GetValue<bool>() ?? (!fieldIndex.HasValue);

        var doc = new Document(path);
        var fields = doc.Range.Fields.ToList();
        var updatedCount = 0;
        var errors = new List<string>();
        
        if (fieldIndex.HasValue)
        {
            if (fieldIndex.Value < 0 || fieldIndex.Value >= fields.Count)
            {
                throw new ArgumentException($"功能變數索引 {fieldIndex.Value} 超出範圍 (文檔共有 {fields.Count} 個功能變數)");
            }
            
            var field = fields[fieldIndex.Value];
            try
            {
                var oldResult = field.Result ?? "";
                field.Update();
                var newResult = field.Result ?? "";
                updatedCount = 1;
                
                doc.Save(outputPath);
                
                var result = $"成功更新功能變數 #{fieldIndex.Value}\n";
                result += $"類型: {field.Type}\n";
                result += $"代碼: {field.GetFieldCode()}\n";
                result += $"舊結果: {oldResult}\n";
                result += $"新結果: {newResult}\n";
                result += $"輸出: {outputPath}";
                
                return await Task.FromResult(result);
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"無法更新功能變數 #{fieldIndex.Value}: {ex.Message}", ex);
            }
        }
        else
        {
            foreach (var field in fields)
            {
                if (!string.IsNullOrEmpty(fieldTypeFilter))
                {
                    var filterType = fieldTypeFilter.ToUpper();
                    var fieldTypeName = field.Type.ToString().ToUpper();
                    if (!fieldTypeName.Contains(filterType))
                    {
                        continue;
                    }
                }
                
                try
                {
                    field.Update();
                    updatedCount++;
                }
                catch (Exception ex)
                {
                    errors.Add($"功能變數 {field.Type} (索引 {fields.IndexOf(field)}): {ex.Message}");
                }
            }
            
            doc.Save(outputPath);
            
            var result = $"成功更新 {updatedCount} 個功能變數\n";
            if (!string.IsNullOrEmpty(fieldTypeFilter))
            {
                result += $"過濾類型: {fieldTypeFilter}\n";
            }
            
            if (errors.Count > 0)
            {
                result += $"\n更新失敗的功能變數 ({errors.Count} 個):\n";
                foreach (var error in errors)
                {
                    result += $"  - {error}\n";
                }
            }
            
            result += $"輸出: {outputPath}";
            
            return await Task.FromResult(result);
        }
    }

    private async Task<string> GetFieldsAsync(JsonObject? arguments, string path)
    {
        var fieldTypeFilter = arguments?["fieldType"]?.GetValue<string>();
        var includeCode = arguments?["includeCode"]?.GetValue<bool>() ?? true;
        var includeResult = arguments?["includeResult"]?.GetValue<bool>() ?? true;

        var doc = new Document(path);
        var fields = new List<FieldInfo>();
        var fieldIndex = 0;
        
        foreach (Field field in doc.Range.Fields)
        {
            if (!string.IsNullOrEmpty(fieldTypeFilter))
            {
                var filterType = fieldTypeFilter.ToUpper();
                var fieldTypeName = field.Type.ToString().ToUpper();
                if (!fieldTypeName.Contains(filterType) && filterType != "ALL")
                {
                    continue;
                }
            }
            
            var fieldInfo = new FieldInfo
            {
                Index = fieldIndex++,
                Type = field.Type.ToString(),
                Code = field.GetFieldCode(),
                Result = includeResult ? (field.Result ?? "") : null,
                IsLocked = field.IsLocked,
                IsDirty = field.IsDirty
            };
            
            if (field is FieldHyperlink hyperlinkField)
            {
                fieldInfo.ExtraInfo = $"Address: {hyperlinkField.Address ?? ""}, ScreenTip: {hyperlinkField.ScreenTip ?? ""}";
            }
            else if (field is FieldRef refField)
            {
                fieldInfo.ExtraInfo = $"Bookmark: {refField.BookmarkName ?? ""}";
            }
            
            fields.Add(fieldInfo);
        }
        
        var result = new StringBuilder();
        result.AppendLine($"文檔中共有 {fields.Count} 個功能變數\n");
        
        if (fields.Count == 0)
        {
            result.AppendLine("未找到功能變數");
            return await Task.FromResult(result.ToString());
        }
        
        result.AppendLine("【功能變數列表】");
        result.AppendLine(new string('-', 80));
        
        foreach (var fieldInfo in fields)
        {
            result.AppendLine($"索引: {fieldInfo.Index}");
            result.AppendLine($"類型: {fieldInfo.Type}");
            
            if (includeCode)
            {
                result.AppendLine($"代碼: {fieldInfo.Code}");
            }
            
            if (includeResult && fieldInfo.Result != null)
            {
                result.AppendLine($"結果: {fieldInfo.Result}");
            }
            
            if (fieldInfo.ExtraInfo != null)
            {
                result.AppendLine($"額外資訊: {fieldInfo.ExtraInfo}");
            }
            
            result.AppendLine($"鎖定: {(fieldInfo.IsLocked ? "是" : "否")}");
            result.AppendLine($"需要更新: {(fieldInfo.IsDirty ? "是" : "否")}");
            result.AppendLine(new string('-', 80));
        }
        
        var typeGroups = fields.GroupBy(f => f.Type).OrderBy(g => g.Key);
        result.AppendLine("\n【按類型統計】");
        foreach (var group in typeGroups)
        {
            result.AppendLine($"{group.Key}: {group.Count()} 個");
        }
        
        return await Task.FromResult(result.ToString());
    }

    private async Task<string> GetFieldDetailAsync(JsonObject? arguments, string path)
    {
        var fieldIndex = arguments?["fieldIndex"]?.GetValue<int>() ?? throw new ArgumentException("fieldIndex is required");

        var doc = new Document(path);
        var fields = doc.Range.Fields.ToList();
        
        if (fieldIndex < 0 || fieldIndex >= fields.Count)
        {
            throw new ArgumentException($"功能變數索引 {fieldIndex} 超出範圍 (文檔共有 {fields.Count} 個功能變數)");
        }
        
        var field = fields[fieldIndex];
        var result = new StringBuilder();
        
        result.AppendLine("【功能變數詳細資訊】");
        result.AppendLine(new string('=', 80));
        result.AppendLine($"索引: {fieldIndex}");
        result.AppendLine($"類型: {field.Type}");
        result.AppendLine($"類型代碼: {(int)field.Type}");
        result.AppendLine($"代碼: {field.GetFieldCode()}");
        result.AppendLine($"結果: {field.Result ?? "(無結果)"}");
        result.AppendLine($"鎖定: {(field.IsLocked ? "是" : "否")}");
        result.AppendLine($"需要更新: {(field.IsDirty ? "是" : "否")}");
        
        if (field is FieldHyperlink hyperlinkField)
        {
            result.AppendLine($"地址: {hyperlinkField.Address ?? "(無)"}");
            result.AppendLine($"提示文字: {hyperlinkField.ScreenTip ?? "(無)"}");
        }
        else if (field is FieldRef refField)
        {
            result.AppendLine($"書籤名稱: {refField.BookmarkName ?? "(無)"}");
        }
        
        result.AppendLine(new string('=', 80));
        
        return await Task.FromResult(result.ToString());
    }

    private async Task<string> AddFormFieldAsync(JsonObject? arguments, string path, string outputPath)
    {
        var fieldType = arguments?["formFieldType"]?.GetValue<string>() ?? throw new ArgumentException("formFieldType is required");
        var fieldName = arguments?["fieldName"]?.GetValue<string>() ?? throw new ArgumentException("fieldName is required");
        var defaultValue = arguments?["defaultValue"]?.GetValue<string>();
        var optionsArray = arguments?["options"]?.AsArray();
        var checkedValue = arguments?["checkedValue"]?.GetValue<bool?>();

        var doc = new Document(path);
        var builder = new DocumentBuilder(doc);
        builder.MoveToDocumentEnd();

        FormField field;
        switch (fieldType.ToLower())
        {
            case "textinput":
                field = builder.InsertTextInput(fieldName, TextFormFieldType.Regular, "", defaultValue ?? "", 0);
                break;

            case "checkbox":
                field = builder.InsertCheckBox(fieldName, checkedValue ?? false, 0);
                break;

            case "dropdown":
                if (optionsArray == null || optionsArray.Count == 0)
                {
                    throw new ArgumentException("options array is required for DropDown type");
                }
                var options = optionsArray.Select(o => o?.GetValue<string>()).Where(o => !string.IsNullOrEmpty(o)).ToArray();
                field = builder.InsertComboBox(fieldName, options, 0);
                break;

            default:
                throw new ArgumentException($"Invalid fieldType: {fieldType}. Must be 'TextInput', 'CheckBox', or 'DropDown'");
        }

        doc.Save(outputPath);
        return await Task.FromResult($"{fieldType} field '{fieldName}' added: {outputPath}");
    }

    private async Task<string> EditFormFieldAsync(JsonObject? arguments, string path, string outputPath)
    {
        var fieldName = arguments?["fieldName"]?.GetValue<string>() ?? throw new ArgumentException("fieldName is required");
        var value = arguments?["value"]?.GetValue<string>();
        var checkedValue = arguments?["checkedValue"]?.GetValue<bool?>();
        var selectedIndex = arguments?["selectedIndex"]?.GetValue<int?>();

        var doc = new Document(path);
        var field = doc.Range.FormFields[fieldName];

        if (field == null)
        {
            throw new ArgumentException($"Form field '{fieldName}' not found");
        }

        if (field.Type == FieldType.FieldFormTextInput && value != null)
        {
            field.Result = value;
        }
        else if (field.Type == FieldType.FieldFormCheckBox && checkedValue.HasValue)
        {
            field.Checked = checkedValue.Value;
        }
        else if (field.Type == FieldType.FieldFormDropDown && selectedIndex.HasValue)
        {
            if (selectedIndex.Value >= 0 && selectedIndex.Value < field.DropDownItems.Count)
            {
                field.DropDownSelectedIndex = selectedIndex.Value;
            }
        }

        doc.Save(outputPath);
        return await Task.FromResult($"Form field '{fieldName}' updated: {outputPath}");
    }

    private async Task<string> DeleteFormFieldAsync(JsonObject? arguments, string path, string outputPath)
    {
        var fieldName = arguments?["fieldName"]?.GetValue<string>();
        var fieldNamesArray = arguments?["fieldNames"]?.AsArray();

        var doc = new Document(path);
        var formFields = doc.Range.FormFields;

        List<string> fieldsToDelete;
        if (fieldNamesArray != null && fieldNamesArray.Count > 0)
        {
            fieldsToDelete = fieldNamesArray.Select(f => f?.GetValue<string>()).Where(f => !string.IsNullOrEmpty(f)).Select(f => f!).ToList();
        }
        else if (!string.IsNullOrEmpty(fieldName))
        {
            fieldsToDelete = new List<string> { fieldName };
        }
        else
        {
            fieldsToDelete = formFields.Cast<FormField>().Select(f => f.Name).ToList();
        }

        int deletedCount = 0;
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
        return await Task.FromResult($"Deleted {deletedCount} form field(s): {outputPath}");
    }

    private async Task<string> GetFormFieldsAsync(JsonObject? arguments, string path)
    {
        var doc = new Document(path);
        var sb = new StringBuilder();

        sb.AppendLine("=== Form Fields ===");
        sb.AppendLine();

        var formFields = doc.Range.FormFields.Cast<FormField>().ToList();
        for (int i = 0; i < formFields.Count; i++)
        {
            var field = formFields[i];
            sb.AppendLine($"[{i + 1}] Name: {field.Name}");
            sb.AppendLine($"    Type: {field.Type}");

            switch (field.Type)
            {
                case FieldType.FieldFormTextInput:
                    sb.AppendLine($"    Value: {field.Result ?? "(empty)"}");
                    break;
                case FieldType.FieldFormCheckBox:
                    sb.AppendLine($"    Checked: {field.Checked}");
                    break;
                case FieldType.FieldFormDropDown:
                    sb.AppendLine($"    Selected Index: {field.DropDownSelectedIndex}");
                    sb.AppendLine($"    Options: {string.Join(", ", field.DropDownItems.Cast<string>())}");
                    break;
            }
            sb.AppendLine();
        }

        sb.AppendLine($"Total Form Fields: {formFields.Count}");

        return await Task.FromResult(sb.ToString());
    }

    private class FieldInfo
    {
        public int Index { get; set; }
        public string Type { get; set; } = "";
        public string Code { get; set; } = "";
        public string? Result { get; set; }
        public bool IsLocked { get; set; }
        public bool IsDirty { get; set; }
        public string? ExtraInfo { get; set; }
    }
}

