using System.Text.RegularExpressions;
using DollarSignEngine;
using DocuChef.Logging;
using DocuChef.Presentation.Models;

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
            // TODO: 보간식이 올바른 인덱스를 가지는지 확인하기 위해 일시적으로 주석처리
            // 실제 평가 대신 템플릿을 그대로 반환하여 보간식을 확인할 수 있도록 함
            var options = new DollarSignOptions
            {
                SupportDollarSignSyntax = true
            };
            var result = DollarSign.Eval(dollarSignTemplate, variables, options);
            return result ?? string.Empty;

            // // 보간식을 그대로 반환 (디버깅용)
            // Logger.Debug($"DataBinder.EvaluateTemplate: 평가 없이 템플릿 반환 - '{dollarSignTemplate}'");
            // return dollarSignTemplate;
        }
        catch (Exception ex)
        {
            Logger.Debug($"DataBinder.EvaluateTemplate: 평가 중 오류 발생 - {ex.Message}");
            return dollarSignTemplate;
        }
    }    /// <summary>
         /// Binds data to template expressions in the given text.
         /// Converts context operators (>) to underscore notation (___) and evaluates using DollarSignEngine.
         /// This overload only creates variables for expressions that are actually used in the template.
         /// </summary>
         /// <param name="template">Template text containing expressions</param>
         /// <param name="data">Data object to bind</param>
         /// <param name="usedExpressions">Set of expressions actually used in the template</param>
         /// <param name="indexOffset">Index offset for array expressions</param>
         /// <param name="contextPath">Context path for the current slide (e.g., "Products[0]", "Products[1]")</param>
         /// <returns>Text with all expressions evaluated</returns>
    public string BindData(string template, object data, ISet<string> usedExpressions, int indexOffset = 0, string? contextPath = null)
    {
        if (string.IsNullOrEmpty(template) || data == null)
            return template ?? string.Empty;

        try
        {
            // Pre-process nested expressions first
            var preprocessedTemplate = PreprocessNestedExpressions(template, data, indexOffset);
            Logger.Debug($"DataBinder: Original template: '{template}'");
            Logger.Debug($"DataBinder: Preprocessed template: '{preprocessedTemplate}'");            // Convert template to DollarSign format
            var dollarSignTemplate = ConvertToTemplate(preprocessedTemplate);
            Logger.Debug($"DataBinder: Converted template to '{dollarSignTemplate}'");            // Create context variables for the data - OPTIMIZED VERSION
            var variables = CreateContextOperatorVariablesFiltered(data, usedExpressions, contextPath);

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
            }

            // Apply index offset to array expressions
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
    }    /// <summary>
         /// Creates context operator variables for only the expressions that are actually used (optimized version).
         /// Extracts variable names from expressions and only creates variables for those.
         /// </summary>
         /// <param name="data">Source data object</param>
         /// <param name="usedExpressions">Set of expressions actually used in the template</param>
         /// <param name="contextPath">Context path for the current slide (e.g., "Products[0]", "Products[1]")</param>
    private Dictionary<string, object> CreateContextOperatorVariablesFiltered(object data, ISet<string> usedExpressions, string? contextPath = null)
    {
        var variables = new Dictionary<string, object>();

        if (data == null || usedExpressions == null || usedExpressions.Count == 0)
            return variables;        // Extract variable names from the used expressions
        var requiredVariables = ExtractVariableNamesFromExpressions(usedExpressions);

        Logger.Debug($"DataBinder: Filtering variables based on {usedExpressions.Count} expressions");
        Logger.Debug($"DataBinder: Required variables: {string.Join(", ", requiredVariables)}");
        Logger.Debug($"DataBinder: Context path: '{contextPath ?? "null"}'");

        // Add the root data object
        variables["Data"] = data;

        // Create flattened variables for only the required properties
        CreateContextVariablesRecursiveFiltered(data, string.Empty, variables, 0, requiredVariables);

        // Create context-specific variables based on contextPath
        if (!string.IsNullOrEmpty(contextPath))
        {
            CreateContextSpecificVariables(data, variables, contextPath, requiredVariables);
        }

        return variables;
    }    /// <summary>
         /// Extracts variable names from expressions (similar to PowerPointProcessor but adapted for DataBinder).
         /// </summary>
    private HashSet<string> ExtractVariableNamesFromExpressions(ISet<string> expressions)
    {
        var variableNames = new HashSet<string>();

        foreach (var expression in expressions)
        {
            if (string.IsNullOrWhiteSpace(expression))
                continue;

            var cleanExpression = expression.Trim().TrimStart('$', '{').TrimEnd('}');

            // Handle array indexing with property access: Items[0].Name -> Items, Name
            var arrayPropertyMatch = System.Text.RegularExpressions.Regex.Match(cleanExpression, @"^([a-zA-Z_]\w*)(?:\[\d+\])\.([a-zA-Z_]\w*)");
            if (arrayPropertyMatch.Success)
            {
                variableNames.Add(arrayPropertyMatch.Groups[1].Value); // Items
                variableNames.Add(arrayPropertyMatch.Groups[2].Value); // Name (property name for filtering)
                continue;
            }

            // Handle array indexing: Items[0] -> Items
            var arrayMatch = System.Text.RegularExpressions.Regex.Match(cleanExpression, @"^([a-zA-Z_]\w*)(?:\[\d+\])?");
            if (arrayMatch.Success)
            {
                variableNames.Add(arrayMatch.Groups[1].Value);
            }

            // Handle property access: object.property -> object
            var propertyMatch = System.Text.RegularExpressions.Regex.Match(cleanExpression, @"^([a-zA-Z_]\w*)(?:\.|:)");
            if (propertyMatch.Success)
            {
                variableNames.Add(propertyMatch.Groups[1].Value);
            }

            // Handle function calls: ppt.Image(LogoPath) -> LogoPath
            var functionMatch = System.Text.RegularExpressions.Regex.Match(cleanExpression, @"ppt\.\w+\(([^)]+)\)");
            if (functionMatch.Success)
            {
                var parameters = functionMatch.Groups[1].Value.Split(',');
                foreach (var param in parameters)
                {
                    var trimmedParam = param.Trim();
                    if (!trimmedParam.StartsWith("\"") && !trimmedParam.StartsWith("'"))
                    {
                        variableNames.Add(trimmedParam);
                    }
                }
            }

            // Handle simple variable names
            if (System.Text.RegularExpressions.Regex.IsMatch(cleanExpression, @"^[a-zA-Z_]\w*$"))
            {
                variableNames.Add(cleanExpression);
            }
        }

        return variableNames;
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
    }    /// <summary>
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
    }    /// <summary>
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
            // Pre-process nested expressions first
            var preprocessedTemplate = PreprocessNestedExpressions(template, data, indexOffset);
            Logger.Debug($"DataBinder: Original template: '{template}'");
            Logger.Debug($"DataBinder: Preprocessed template: '{preprocessedTemplate}'");

            // Convert template to DollarSign format
            var dollarSignTemplate = ConvertToTemplate(preprocessedTemplate);
            Logger.Debug($"DataBinder: Converted template to '{dollarSignTemplate}'");            // Create context variables for the data - OPTIMIZED VERSION            // Create context variables for the data - OPTIMIZED VERSION with context support
            var variables = CreateContextOperatorVariablesFiltered(data, usedExpressions, contextPath);

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

    /// <summary>
    /// Resolves an expression using filtered variables for optimization.
    /// </summary>
    public object? ResolveExpressionWithFilteredVariables(string expression, object data, ISet<string> usedExpressions)
    {
        if (string.IsNullOrEmpty(expression) || data == null)
            return null;

        try
        {
            // Use the optimized BindData method with filtered variables
            var result = BindData(expression, data, usedExpressions);
            return result;
        }
        catch (Exception ex)
        {
            Logger.Error($"DataBinder: Error resolving expression '{expression}' with filtered variables: {ex.Message}");
            return null;
        }
    }

    /// <summary>
    /// Resolves an expression using filtered variables and custom functions for optimization.
    /// </summary>
    public object? ResolveExpressionWithFilteredVariables(string expression, object data, ISet<string> usedExpressions, Dictionary<string, Func<object, object>>? customFunctions)
    {
        if (string.IsNullOrEmpty(expression) || data == null)
            return null;

        try
        {
            // Use the optimized BindData method with filtered variables and custom functions
            var result = BindData(expression, data, usedExpressions, customFunctions);
            return result;
        }
        catch (Exception ex)
        {
            Logger.Error($"DataBinder: Error resolving expression '{expression}' with filtered variables and custom functions: {ex.Message}");
            return null;
        }
    }

    /// <summary>
    /// 누락된 메서드들 추가 - 원래 DataBinder에서 누락된 부분들
    /// </summary>
    /// 
    private string PreprocessNestedExpressions(string template, object data, int indexOffset)
    {
        // 기본적으로 템플릿을 그대로 반환
        return template;
    }

    private string ConvertToTemplate(string template)
    {
        // Context operator (>) 를 underscore (___) 로 변환
        return ContextOperatorRegex.Replace(template, "$1___$2");
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

    /// <summary>
    /// Creates context-specific variables based on the provided context path.
    /// For example, if contextPath is "Products[0]", creates Products___Items variable
    /// that points to Products[0].Items.
    /// </summary>
    private void CreateContextSpecificVariables(object data, Dictionary<string, object> variables, string contextPath, HashSet<string> requiredVariables)
    {
        if (string.IsNullOrEmpty(contextPath) || data == null)
            return;

        Logger.Debug($"DataBinder: Creating context-specific variables for path '{contextPath}'");

        try
        {
            // Parse context path (e.g., "Products[0]" -> rootProperty="Products", index=0)
            var match = System.Text.RegularExpressions.Regex.Match(contextPath, @"^(\w+)\[(\d+)\]$");
            if (!match.Success)
            {
                Logger.Debug($"DataBinder: Context path '{contextPath}' doesn't match expected pattern");
                return;
            }

            string rootProperty = match.Groups[1].Value;
            int contextIndex = int.Parse(match.Groups[2].Value);

            Logger.Debug($"DataBinder: Parsed context - root: '{rootProperty}', index: {contextIndex}");

            // Get the root collection from data
            object? rootCollection = GetPropertyValue(data, rootProperty);
            if (rootCollection is not System.Collections.IEnumerable enumerable || rootCollection is string)
            {
                Logger.Debug($"DataBinder: Root property '{rootProperty}' is not a collection");
                return;
            }

            var list = enumerable.Cast<object>().ToList();
            if (contextIndex >= list.Count)
            {
                Logger.Debug($"DataBinder: Context index {contextIndex} is out of bounds for collection with {list.Count} items");
                return;
            }

            // Get the specific context object
            var contextObject = list[contextIndex];
            if (contextObject == null)
            {
                Logger.Debug($"DataBinder: Context object at index {contextIndex} is null");
                return;
            }

            Logger.Debug($"DataBinder: Context object type: {contextObject.GetType().Name}");

            // Create context-specific variables for properties of the context object
            CreateContextVariablesForObject(contextObject, rootProperty, variables, requiredVariables);
        }
        catch (Exception ex)
        {
            Logger.Error($"DataBinder: Error creating context-specific variables for '{contextPath}': {ex.Message}");
        }
    }

    /// <summary>
    /// Creates variables for properties of a context object with the root property prefix.
    /// For example, if rootProperty is "Products" and object has "Items" property,
    /// creates "Products___Items" variable.
    /// </summary>
    private void CreateContextVariablesForObject(object contextObject, string rootProperty, Dictionary<string, object> variables, HashSet<string> requiredVariables)
    {
        if (contextObject == null)
            return;

        Logger.Debug($"DataBinder: Creating variables for context object properties with prefix '{rootProperty}___'");

        if (contextObject is Dictionary<string, object> dictionary)
        {
            foreach (var kvp in dictionary)
            {
                var variableName = $"{rootProperty}___{kvp.Key}";

                // Only create if this variable is required
                if (IsVariableRequired(variableName, kvp.Key, requiredVariables))
                {
                    variables[variableName] = kvp.Value;
                    Logger.Debug($"DataBinder: Created context variable '{variableName}' = {GetValueDescription(kvp.Value)}");
                }
            }
        }
        else
        {
            // Handle regular objects using reflection
            var properties = contextObject.GetType().GetProperties();
            foreach (var property in properties)
            {
                try
                {
                    var variableName = $"{rootProperty}___{property.Name}";

                    // Only create if this variable is required
                    if (IsVariableRequired(variableName, property.Name, requiredVariables))
                    {
                        var value = property.GetValue(contextObject);
                        variables[variableName] = value ?? new object();
                        Logger.Debug($"DataBinder: Created context variable '{variableName}' = {GetValueDescription(value)}");
                    }
                }
                catch (Exception ex)
                {
                    Logger.Debug($"DataBinder: Error accessing property '{property.Name}': {ex.Message}");
                }
            }
        }
    }

    /// <summary>
    /// Gets a description of a value for logging purposes.
    /// </summary>
    private string GetValueDescription(object? value)
    {
        if (value == null) return "null";
        if (value is System.Collections.IEnumerable enumerable && !(value is string))
        {
            var list = enumerable.Cast<object>().ToList();
            return $"[{list.Count} items]";
        }
        return value.ToString() ?? "null";
    }

    /// <summary>
    /// Gets property value from an object using property name.
    /// </summary>
    private object? GetPropertyValue(object obj, string propertyName)
    {
        if (obj == null) return null;

        if (obj is Dictionary<string, object> dictionary)
        {
            return dictionary.TryGetValue(propertyName, out var value) ? value : null;
        }

        var property = obj.GetType().GetProperty(propertyName);
        return property?.GetValue(obj);
    }
}