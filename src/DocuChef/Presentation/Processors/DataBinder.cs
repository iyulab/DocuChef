using DollarSignEngine;
using System.Linq;
using System.Reflection;

namespace DocuChef.Presentation.Processors;

/// <summary>
/// Handles data binding by converting template expressions to DollarSign format and evaluating them.
/// Supports context operators (>) for nested property access.
/// </summary>
public class DataBinder
{
    private static readonly Regex ContextOperatorRegex = new(@"(\w+)>([^}\s]+)", RegexOptions.Compiled);
    private static readonly Regex DollarSignExpressionRegex = new(@"\$\{([^}]+)\}", RegexOptions.Compiled);

    /// <summary>
    /// 단일 평가 함수 - 모든 DollarSign.EvalAsync 호출을 여기서 처리
    /// 현재는 보간식 확인을 위해 주석처리되어 있음
    /// </summary>
    private string EvaluateTemplate(string dollarSignTemplate, Dictionary<string, object> variables)
    {
        try
        {
            // Pre-process array indexing expressions for better handling
            //dollarSignTemplate = PreprocessArrayIndexingExpressions(dollarSignTemplate, variables);

            // Setup DollarSignEngine options
            var options = new DollarSignOptions
            {
                SupportDollarSignSyntax = true,
                ThrowOnError = true
            };

            var result = DollarSign.Eval(dollarSignTemplate, variables, options);
            return result ?? string.Empty;
        }
        catch (Exception ex)
        {
            Logger.Debug($"DataBinder.EvaluateTemplate: 평가 중 오류 발생 - {ex.Message}");
            return string.Empty;
        }
    }

    /// <summary>
    /// Recursively creates context variables for nested objects and arrays, but only for required variables.
    /// </summary>
    private void CreateContextVariablesRecursiveFiltered(object obj, string prefix, Dictionary<string, object> variables, int depth, HashSet<string> requiredVariables)
    {
        if (obj == null || depth > 5) // Prevent infinite recursion
            return;

        // Handle Dictionary<string, object> specially
        if (obj is Dictionary<string, object> dictionary)
        {
            foreach (var kvp in dictionary)
            {
                if (kvp.Value == null)
                    continue;

                var propertyName = string.IsNullOrEmpty(prefix)
                    ? kvp.Key
                    : $"{prefix}___{kvp.Key}";

                // Only process if this variable is required
                if (IsVariableRequired(propertyName, kvp.Key, requiredVariables))
                {
                    // Add the property value
                    variables[propertyName] = kvp.Value;

                    // Handle arrays/collections - need to extract nested properties from array elements
                    if (kvp.Value is System.Collections.IEnumerable enumerable && !(kvp.Value is string))
                    {
                        CreateContextVariablesForArrayFiltered(enumerable, propertyName, variables, depth + 1, requiredVariables);

                        // For context operators, also extract common properties from array elements
                        var list = enumerable.Cast<object>().ToList();
                        if (list.Count > 0 && list[0] != null)
                        {
                            // Extract common properties from the first element to create context operator paths
                            ExtractArrayElementPropertiesFiltered(list, propertyName, variables, depth + 1, requiredVariables);
                        }
                    }
                    // Handle complex objects
                    else if (!IsSimpleType(kvp.Value.GetType()))
                    {
                        CreateContextVariablesRecursiveFiltered(kvp.Value, propertyName, variables, depth + 1, requiredVariables);
                    }
                }
            }
            return;
        }

        // Handle regular objects
        var properties = obj.GetType().GetProperties();
        foreach (var property in properties)
        {
            if (!property.CanRead)
                continue;

            try
            {
                var propertyName = string.IsNullOrEmpty(prefix)
                    ? property.Name
                    : $"{prefix}___{property.Name}";

                // Only process if this variable is required
                if (IsVariableRequired(propertyName, property.Name, requiredVariables))
                {
                    var value = property.GetValue(obj);
                    if (value == null)
                        continue;

                    variables[propertyName] = value;

                    // Handle arrays/collections - need to extract nested properties from array elements
                    if (value is System.Collections.IEnumerable enumerable && !(value is string))
                    {
                        CreateContextVariablesForArrayFiltered(enumerable, propertyName, variables, depth + 1, requiredVariables);

                        // For context operators, also extract common properties from array elements
                        var list = enumerable.Cast<object>().ToList();
                        if (list.Count > 0 && list[0] != null)
                        {
                            // Extract common properties from the first element to create context operator paths
                            ExtractArrayElementPropertiesFiltered(list, propertyName, variables, depth + 1, requiredVariables);
                        }
                    }
                    // Handle complex objects
                    else if (!IsSimpleType(value.GetType()))
                    {
                        CreateContextVariablesRecursiveFiltered(value, propertyName, variables, depth + 1, requiredVariables);
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Warning($"DataBinder: Error processing property '{property.Name}': {ex.Message}");
            }
        }
    }

    /// <summary>
    /// Checks if a variable is required based on the required variables list.
    /// </summary>
    private bool IsVariableRequired(string fullPropertyName, string propertyName, HashSet<string> requiredVariables)
    {
        return requiredVariables.Contains(propertyName) ||
               requiredVariables.Contains(fullPropertyName) ||
               requiredVariables.Any(req => req.StartsWith(propertyName + ".") || req.StartsWith(propertyName + "["));
    }

    /// <summary>
    /// Creates context variables for array/collection elements (filtered version).
    /// </summary>
    private void CreateContextVariablesForArrayFiltered(System.Collections.IEnumerable enumerable, string propertyName, Dictionary<string, object> variables, int depth, HashSet<string> requiredVariables)
    {
        var list = enumerable.Cast<object>().ToList();

        // Add indexed access for each element (only if needed)
        for (int i = 0; i < list.Count; i++)
        {
            var element = list[i];
            if (element == null)
                continue; var indexedName = $"{propertyName}_{i}";

            // Only add if this indexed variable is required
            if (requiredVariables.Any(req => req.StartsWith(propertyName + "_") || req == propertyName))
            {
                variables[indexedName] = element;

                // For array elements, we need to expose their properties as flattened variables 
                // so DollarSign can access them directly (e.g., Items[0].Id -> Items_0___Id)
                if (!IsSimpleType(element.GetType()))
                {
                    var elementProperties = element.GetType().GetProperties();
                    foreach (var prop in elementProperties)
                    {
                        if (!prop.CanRead)
                            continue;

                        try
                        {
                            var flattenedName = $"{indexedName}___{prop.Name}";
                            var propValue = prop.GetValue(element);
                            if (propValue != null)
                            {
                                variables[flattenedName] = propValue;
                            }
                        }
                        catch (Exception ex)
                        {
                            Logger.Warning($"DataBinder: Error flattening array element property '{prop.Name}': {ex.Message}");
                        }
                    }

                    // Also recursively process for deeper nesting
                    CreateContextVariablesRecursiveFiltered(element, indexedName, variables, depth + 1, requiredVariables);
                }
            }
        }

        // Add the array itself (always needed if the property is required)
        if (requiredVariables.Contains(propertyName.Split('.').Last()) || requiredVariables.Contains(propertyName))
        {
            variables[propertyName] = list;
        }
    }

    /// <summary>
    /// Extracts properties from array elements for context operator paths (filtered version).
    /// </summary>
    private void ExtractArrayElementPropertiesFiltered(List<object> list, string arrayPropertyName, Dictionary<string, object> variables, int depth, HashSet<string> requiredVariables)
    {
        if (list.Count == 0 || list[0] == null)
            return;

        var firstElement = list[0];
        var elementProperties = firstElement.GetType().GetProperties();

        foreach (var prop in elementProperties)
        {
            if (!prop.CanRead)
                continue;

            try
            {
                var contextPath = $"{arrayPropertyName}___{prop.Name}";

                // Only process if this context path is required
                if (IsVariableRequired(contextPath, prop.Name, requiredVariables))
                {
                    var values = new List<object?>();
                    foreach (var item in list)
                    {
                        if (item != null)
                        {
                            var value = prop.GetValue(item);
                            values.Add(value);
                        }
                        else
                        {
                            values.Add(null);
                        }
                    }
                    variables[contextPath] = values;
                }
            }
            catch (Exception ex)
            {
                Logger.Warning($"DataBinder: Error extracting array element property '{prop.Name}': {ex.Message}");
            }
        }
    }

    /// <summary>
    /// Binds data to template expressions in the given text with custom functions.
    /// Converts context operators (>) to underscore notation (___) and evaluates using DollarSignEngine.
    /// </summary>
    /// <param name="template">Template text containing expressions</param>
    /// <param name="data">Data object to bind</param>
    /// <param name="usedExpressions">Set of expressions actually used for filtering variables</param>
    /// <param name="customFunctions">Custom functions to register for use in expressions</param>
    /// <param name="indexOffset">Index offset for array expressions</param>
    /// <param name="contextPath">Context path for the current slide (e.g., "Products[0]", "Products[1]")</param>
    /// <returns>Text with all expressions evaluated</returns>
    public string BindData(string template, object data, ISet<string> usedExpressions, Dictionary<string, Func<object, object>>? customFunctions, int indexOffset = 0, string? contextPath = null)
    {
        if (string.IsNullOrEmpty(template) || data == null)
            return template ?? string.Empty;

        try
        {
            Logger.Debug($"DataBinder: Original template: '{template}'");
            var (dollarSignTemplate, variables) = ResolveExpressionAndVariables(template, data, contextPath);

            // Add PPT functions to variables - reuse existing instance if available
            DocuChef.Presentation.Functions.PPTFunctions pptFunctions;
            if (data is Dictionary<string, object> dataDict &&
                dataDict.TryGetValue("ppt", out var existingPpt) &&
                existingPpt is DocuChef.Presentation.Functions.PPTFunctions existingPptFunctions)
            {
                // Reuse existing PPTFunctions instance (preserves image cache)
                pptFunctions = existingPptFunctions;
                Logger.Debug($"DataBinder: Reusing existing PPTFunctions instance with {existingPptFunctions.GetAllImageCache().Count} cached images");
            }
            else
            {
                // Create new PPTFunctions instance
                pptFunctions = new DocuChef.Presentation.Functions.PPTFunctions(variables);
                Logger.Debug($"DataBinder: Created new PPTFunctions instance");
            }
            variables["ppt"] = pptFunctions;

            // Add custom functions to variables
            if (customFunctions != null)
            {
                foreach (var kvp in customFunctions)
                {
                    variables[kvp.Key] = kvp.Value;
                }
                Logger.Debug($"DataBinder: Added {customFunctions.Count} custom functions");
            }

            Logger.Debug($"DataBinder: Created {variables.Count} filtered variables:");
            foreach (var kvp in variables)
            {
                if (kvp.Value is System.Collections.IEnumerable enumerable && !(kvp.Value is string))
                {
                    var list = enumerable.Cast<object>().ToList();
                    Logger.Debug($"  {kvp.Key} = [{string.Join(", ", list.Take(3).Select(x => x?.ToString() ?? "null"))}] (count: {list.Count})");
                }
                else
                {
                    Logger.Debug($"  {kvp.Key} = {kvp.Value}");
                }
            }            // Apply index offset to array expressions
            if (indexOffset > 0)
            {
                dollarSignTemplate = ApplyIndexOffset(dollarSignTemplate, indexOffset);
            }

            // 단일 평가 함수 사용 (현재는 보간식을 그대로 반환)
            var result = EvaluateTemplate(dollarSignTemplate, variables);

            Logger.Debug($"DataBinder: Template '{dollarSignTemplate}' evaluated to '{result}'");
            return result;
        }
        catch (Exception ex)
        {
            Logger.Error($"DataBinder: Error binding data to template '{template}': {ex.Message}");
            return template; // Return original template on error
        }
    }

    private (string dollarSignTemplate, Dictionary<string, object> variables) ResolveExpressionAndVariables(string expression, object data, string contextPath)
    {
        var variables = ResolveVariables(data);

        if (contextPath.Contains('>'))
        {
            object? vData = variables;
            var splits = contextPath.Split(">");
            var vNames = new List<string>();
            foreach (var split in splits)
            {
                if (split.Contains('[') && split.Contains(']'))
                {
                    var vName = split.Substring(0, split.IndexOf("["));
                    var indexText = split.Substring(split.IndexOf("[") + 1, split.IndexOf("]") - split.IndexOf("[") - 1);
                    var index = int.Parse(indexText);

                    vNames.Add(vName);
                    vData = ResolveValue(vData, vName, index);
                }
                else
                {
                    var vName = split;
                    vNames.Add(vName);
                    vData = ResolveValue(vData, vName, null);
                }
            }

            var syntax = string.Join(">", vNames);
            var vNameText = string.Join("___", vNames);

            var newExpression = expression.Replace(syntax, vNameText);
            variables.Add(vNameText, vData ?? string.Empty);
            return (newExpression, variables);
        }

        return (expression, variables);
    }

    private object? ResolveValue(object data, string property, int? index)
    {
        if (data is IDictionary dic)
        {
            var value = dic[property];
            if (value is IEnumerable collValue)
            {
                return collValue.OfType<object>().ElementAt(index ?? 0);
            }
            else
            {
                return value;
            }
        }
        else if (data is IEnumerable coll)
        {
        }
        else if (data.GetType().GetProperty(property) is PropertyInfo pInfo)
        {
            return pInfo.GetValue(data);
        }
        else
        {
            throw new NotImplementedException($"ResolveValue, {data.GetType().FullName}");
        }
        return data;
    }

    private Dictionary<string, object> ResolveVariables(object data)
    {
        if (data == null)
            return new Dictionary<string, object>();

        var variables = new Dictionary<string, object>();
        if (data is Dictionary<string, object> dataDict)
        {
            // If data is already a dictionary, use it directly
            foreach (var kvp in dataDict)
            {
                variables[kvp.Key] = kvp.Value;
            }
        }
        else
        {
            // object to dictionary conversion
            var properties = data.GetType().GetProperties();
            foreach (var property in properties)
            {
                if (property.CanRead)
                {
                    try
                    {
                        var value = property.GetValue(data);
                        if (value != null)
                        {
                            // Use property name as key, convert to underscore format
                            var key = property.Name.Replace(".", "___");
                            variables[key] = value;
                        }
                    }
                    catch (Exception ex)
                    {
                        Logger.Warning($"DataBinder: Error reading property '{property.Name}': {ex.Message}");
                    }
                }
            }
        }

        return variables;
    }

    private string ApplyIndexOffset(string template, int offset)
    {
        // 기본적으로 템플릿을 그대로 반환 (인덱스 오프셋 적용은 현재 비활성화)
        return template;
    }

    private bool IsSimpleType(Type type)
    {
        return type.IsPrimitive ||
               type == typeof(string) ||
               type == typeof(decimal) ||
               type == typeof(DateTime) ||
               type == typeof(DateTimeOffset) ||
               type == typeof(TimeSpan) ||
               type == typeof(Guid) ||
               (type.IsGenericType && type.GetGenericTypeDefinition() == typeof(Nullable<>));
    }
}