using DollarSignEngine;
using DocuChef.Exceptions;
using DocuChef.Logging;
using System.Collections;
using System.Reflection;
using System.Text.RegularExpressions;

namespace DocuChef.Presentation.Processors;

/// <summary>
/// Handles data binding by converting template expressions to DollarSign format and evaluating them.
/// Supports context operators (>) for nested property access.
/// </summary>
public class DataBinder
{
    // PERFORMANCE OPTIMIZATION: Use instance-based caching instead of static to prevent memory leaks
    private readonly Dictionary<string, Dictionary<string, object>> _variableCache = new();
    private Dictionary<string, object>? _baseVariables;

    // PERFORMANCE OPTIMIZATION: Cache reflection property info to avoid repeated reflection calls
    private readonly Dictionary<Type, PropertyInfo[]> _propertyCache = new();

    /// <summary>
    /// Initialize base variables once for performance optimization
    /// </summary>
    public void PrepareBaseVariables(object data)
    {
        if (_baseVariables == null)
        {
            Logger.Debug("DataBinder: Preparing base variables (one-time initialization)");
            _baseVariables = ResolveVariables(data);
            Logger.Debug($"DataBinder: Cached {_baseVariables.Count} base variables");
        }
        Logger.Debug("Base variables prepared for optimized data binding");
    }

    /// <summary>
    /// Get variables for specific context with caching
    /// </summary>
    public Dictionary<string, object> GetVariablesForContext(object data, string? contextPath)
    {
        var cacheKey = contextPath ?? "";

        if (_variableCache.TryGetValue(cacheKey, out var cachedVariables))
        {
            Logger.Debug($"DataBinder: Using cached variables for context '{cacheKey}' ({cachedVariables.Count} variables)");
            return new Dictionary<string, object>(cachedVariables);
        }

        // Create new variables for this context
        var variables = new Dictionary<string, object>(_baseVariables ?? new Dictionary<string, object>());

        // Apply context-specific transformations
        if (!string.IsNullOrEmpty(contextPath))
        {
            ApplyContextPath(variables, data, contextPath);
        }

        // Cache for future use
        _variableCache[cacheKey] = new Dictionary<string, object>(variables);
        Logger.Debug($"DataBinder: Cached {variables.Count} variables for context '{cacheKey}'");

        return variables;
    }    /// <summary>
         /// Apply context path transformation with optimized parsing
         /// </summary>
    private void ApplyContextPath(Dictionary<string, object> variables, object data, string contextPath)
    {
        if (!contextPath.Contains('>'))
            return;

        object? vData = data;
        var splits = contextPath.Split('>');
        var vNames = new List<string>(splits.Length);

        foreach (var split in splits)
        {
            string vName;
            if (TryParseArrayAccess(split, out vName, out var index))
            {
                // Include array index in variable name to distinguish Products[0] from Products[1]
                vNames.Add($"{vName}__{index}");
                vData = ResolveValue(vData!, vName, index);
            }
            else
            {
                vNames.Add(split);
                vData = ResolveValue(vData!, split, null);
            }
        }
        var vNameText = string.Join("__", vNames);
        variables[vNameText] = vData ?? string.Empty;

        Logger.Debug($"DataBinder: Created context variable '{vNameText}' for path '{contextPath}'");
    }

    /// <summary>
    /// Optimized array access parsing
    /// </summary>
    private static bool TryParseArrayAccess(string segment, out string propertyName, out int index)
    {
        index = 0;
        propertyName = segment;

        if (!segment.Contains('[') || !segment.Contains(']'))
            return false;

        var bracketStart = segment.IndexOf('[');
        var bracketEnd = segment.IndexOf(']');

        if (bracketStart >= bracketEnd)
            return false;

        propertyName = segment.Substring(0, bracketStart);
        var indexText = segment.Substring(bracketStart + 1, bracketEnd - bracketStart - 1);

        return int.TryParse(indexText, out index);
    }    /// <summary>
         /// Get cached property info for better performance
         /// </summary>
    private PropertyInfo[] GetCachedProperties(Type type)
    {
        if (!_propertyCache.TryGetValue(type, out var properties))
        {
            properties = type.GetProperties();
            _propertyCache[type] = properties;
        }
        return properties;
    }

    /// <summary>
    /// Clear variable cache (useful for testing or when data changes significantly)
    /// </summary>
    public void ClearCache()
    {
        _variableCache.Clear();
        _baseVariables = null;
        _propertyCache.Clear();
        Logger.Debug("DataBinder: All caches cleared");
    }
    /// <summary>
    /// 단일 평가 함수 - 모든 DollarSign.EvalAsync 호출을 여기서 처리
    /// 현재는 보간식 확인을 위해 주석처리되어 있음
    /// </summary>
    private string EvaluateTemplate(string dollarSignTemplate, Dictionary<string, object> variables)
    {
        try
        {
            // PERFORMANCE OPTIMIZATION: Pre-validate array indices before DollarSign.Eval
            ValidateArrayIndicesInTemplate(dollarSignTemplate, variables);

            // Setup DollarSignEngine options
            var options = new DollarSignOptions
            {
                SupportDollarSignSyntax = true,
                ThrowOnError = true
            };

            var result = DollarSign.Eval(dollarSignTemplate, variables, options);
            Logger.Debug($"DataBinder.EvaluateTemplate: {dollarSignTemplate}: {result}");
            return result ?? string.Empty;
        }
        catch (DollarSignEngineException dollarSignEngineException)
        {
            if (dollarSignEngineException.InnerException is ArgumentOutOfRangeException ||
                dollarSignEngineException.InnerException is IndexOutOfRangeException)
            {
                throw new DocuChefHideException("Array index out of bounds", dollarSignEngineException);
            }

            Logger.Debug($"DataBinder.EvaluateTemplate: Not an array bounds error, returning empty string");
            return string.Empty;
        }
        catch (DocuChefHideException)
        {
            throw;
        }
        catch (Exception ex)
        {
            Logger.Debug($"DataBinder.EvaluateTemplate: 평가 중 오류 발생 - {ex.Message}");
            return string.Empty;
        }
    }

    /// <summary>
    /// PERFORMANCE OPTIMIZATION: Pre-validate array indices in template expressions
    /// Throws DocuChefHideException early if array bounds are exceeded, avoiding expensive DollarSign.Eval
    /// </summary>
    private void ValidateArrayIndicesInTemplate(string template, Dictionary<string, object> variables)
    {
        if (string.IsNullOrEmpty(template) || !template.Contains("${"))
            return;

        // only array check
        if (template.Contains('[') && template.Contains(']'))
        {
            // Extract ${...} expressions using regex
            var dollarSignPattern = @"\$\{([^}]+)\}";
            var matches = Regex.Matches(template, dollarSignPattern);

            foreach (Match match in matches)
            {
                var expression = match.Groups[1].Value.Trim();
                ValidateArrayIndexInExpression(expression, variables);
            }
        }
    }    /// <summary>
         /// Validate array index bounds in a single expression with unified logic
         /// </summary>
    private void ValidateArrayIndexInExpression(string expression, Dictionary<string, object> variables)
    {
        // Look for array access patterns like "Items[0]", "Users[1].Name", "Products[2].Details[0]"
        var arrayAccessPattern = @"(\w+)\[(\d+)\]";
        var matches = Regex.Matches(expression, arrayAccessPattern);

        foreach (Match match in matches)
        {
            var arrayName = match.Groups[1].Value;
            var indexText = match.Groups[2].Value;

            if (int.TryParse(indexText, out var index) &&
                variables.TryGetValue(arrayName, out var arrayValue))
            {
                var bounds = GetCollectionBounds(arrayValue);
                if (bounds.HasValue && (index >= bounds.Value || index < 0))
                {
                    Logger.Debug($"DataBinder: Array bounds validation failed - {arrayName}[{index}] exceeds collection size {bounds.Value}");
                    throw new DocuChefHideException($"Array index out of bounds: {arrayName}[{index}] exceeds collection size {bounds.Value}");
                }
            }
        }
    }    /// <summary>
         /// Get collection bounds for various collection types
         /// </summary>
    private static int? GetCollectionBounds(object? collection)
    {
        return collection switch
        {
            IList list => list.Count,
            ICollection<object> col => col.Count,
            IEnumerable enumerable when enumerable is not string => enumerable.Cast<object>().Count(),
            _ => null
        };
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
        }        // Handle regular objects with cached property info
        var properties = GetCachedProperties(obj.GetType());
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
                variables[indexedName] = element;                    // For array elements, we need to expose their properties as flattened variables 
                                                                     // so DollarSign can access them directly (e.g., Items[0].Id -> Items_0___Id)
                if (!IsSimpleType(element.GetType()))
                {
                    var elementProperties = GetCachedProperties(element.GetType());
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
                        catch (TargetInvocationException ex)
                        {
                            Logger.Warning($"DataBinder: TargetInvocationException accessing property '{prop.Name}': {ex.InnerException?.Message ?? ex.Message}");
                        }
                        catch (ArgumentException ex)
                        {
                            Logger.Warning($"DataBinder: ArgumentException accessing property '{prop.Name}': {ex.Message}");
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
            Logger.Debug($"DataBinder: Original template: '{template}', ContextPath: '{contextPath ?? "null"}'");            // Use the improved variable resolution with context awareness
            var variables = GetVariablesForContext(data, contextPath);

            // Convert context operators in template to use context-specific variable names
            var dollarSignTemplate = ConvertContextOperatorsToContextAware(template, contextPath);

            // Create variables for corrected expressions if needed
            CreateVariablesForCorrectedExpressions(variables, data, dollarSignTemplate);

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
            Logger.Debug($"DataBinder: Created {variables.Count} filtered variables");
            // PERFORMANCE OPTIMIZATION: Limit debug logging output
            var logLimit = Math.Min(variables.Count, 10);
            for (int i = 0; i < logLimit; i++)
            {
                var kvp = variables.ElementAt(i);
                if (kvp.Value is System.Collections.IEnumerable enumerable && !(kvp.Value is string))
                {
                    var list = enumerable.Cast<object>().ToList();
                    Logger.Debug($"  {kvp.Key} = [{string.Join(", ", list.Take(3).Select(x => x?.ToString() ?? "null"))}] (count: {list.Count})");
                }
                else
                {
                    var valueStr = kvp.Value?.ToString();
                    Logger.Debug($"  {kvp.Key} = {(valueStr?.Length > 50 ? valueStr.Substring(0, 50) + "..." : valueStr)}");
                }
            }
            if (variables.Count > 10)
            {
                Logger.Debug($"  ... and {variables.Count - 10} more variables");
            }

            // 단일 평가 함수 사용 (현재는 보간식을 그대로 반환)
            var result = EvaluateTemplate(dollarSignTemplate, variables); Logger.Debug($"DataBinder: Template '{dollarSignTemplate}' evaluated to '{result}'");
            return result;
        }
        catch (DocuChefHideException)
        {
            // Re-throw DocuChefHideException to allow element hiding
            throw;
        }
        catch (Exception ex)
        {
            Logger.Error($"DataBinder: Error binding data to template '{template}': {ex.Message}");
            return template; // Return original template on error
        }
    }

    /// <summary>
    /// Convert context operators in template to use context-aware variable names
    /// </summary>
    private string ConvertContextOperatorsToContextAware(string template, string? contextPath)
    {
        if (string.IsNullOrEmpty(contextPath) || !contextPath.Contains('>'))
            return template;

        Logger.Debug($"DataBinder: Converting context operators for contextPath: '{contextPath}'");

        // Extract array indices from context path for context-specific variable names
        var contextParts = contextPath.Split('>');
        var contextVarNames = new List<string>();

        foreach (var part in contextParts)
        {
            if (part.Contains('[') && part.Contains(']'))
            {
                var baseName = part.Substring(0, part.IndexOf("["));
                var indexText = part.Substring(part.IndexOf("[") + 1, part.IndexOf("]") - part.IndexOf("[") - 1);
                contextVarNames.Add($"{baseName}_{indexText}");
            }
            else
            {
                contextVarNames.Add(part);
            }
        }        // Build context-specific variable name
        var contextVarName = string.Join("__", contextVarNames);

        // Replace context operator expressions with context-specific variable names
        var genericPattern = string.Join(">", contextParts.Select(p => p.Contains('[') ? p.Substring(0, p.IndexOf("[")) : p));
        var result = template.Replace(genericPattern, contextVarName);

        Logger.Debug($"DataBinder: Converted '{genericPattern}' to '{contextVarName}' in template");
        return result;
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

    /// <summary>
    /// Resolves a value from data using property name and optional index
    /// </summary>
    private object? ResolveValue(object data, string property, int? index)
    {
        if (data is IDictionary dic)
        {
            if (!dic.Contains(property))
                return null;

            var value = dic[property];
            if (value is IEnumerable collValue && value is not string)
            {
                var list = collValue.OfType<object>().ToList();
                return index.HasValue && index.Value >= 0 && index.Value < list.Count
                    ? list[index.Value]
                    : null;
            }
            return value;
        }
        else if (data is IEnumerable coll && data is not string)
        {
            var list = coll.OfType<object>().ToList();
            return index.HasValue && index.Value >= 0 && index.Value < list.Count
                ? list[index.Value]
                : null;
        }
        else if (data.GetType().GetProperty(property) is PropertyInfo pInfo)
        {
            return pInfo.GetValue(data);
        }

        Logger.Warning($"DataBinder: Cannot resolve property '{property}' from type {data.GetType().FullName}");
        return null;
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
    }    /// <summary>
         /// Create variables for corrected expressions that contain "__" patterns (e.g., Products__1__Items)
         /// Only creates variables that are actually needed based on the template expressions
         /// </summary>
    private void CreateVariablesForCorrectedExpressions(Dictionary<string, object> variables, object data, string template)
    {
        Logger.Debug($"DataBinder: CreateVariablesForCorrectedExpressions called with template: '{template}'");

        // Find all expressions that contain "__" pattern (corrected expressions)
        var regex = new System.Text.RegularExpressions.Regex(@"\$\{([^}]*__[^}]*)\}");
        var matches = regex.Matches(template);

        Logger.Debug($"DataBinder: Found {matches.Count} expressions with '__' pattern");

        foreach (System.Text.RegularExpressions.Match match in matches)
        {
            var expression = match.Groups[1].Value;
            Logger.Debug($"DataBinder: Processing expression: '{expression}'");

            // Extract the base variable name (e.g., "Products__1__Items" -> "Products__1__Items")
            var baseVarMatch = System.Text.RegularExpressions.Regex.Match(expression, @"^([^[.\(]+)");
            if (baseVarMatch.Success)
            {
                var variableName = baseVarMatch.Groups[1].Value;

                // Skip if variable already exists
                if (variables.ContainsKey(variableName))
                {
                    continue;
                }                // Try to extract data for this corrected expression
                var extractedData = ExtractDataForCorrectedExpression(data, variableName);
                if (extractedData != null)
                {
                    variables[variableName] = extractedData;
                    Logger.Debug($"DataBinder: Created variable '{variableName}' with {(extractedData is System.Collections.IEnumerable enumerable && !(extractedData is string) ? enumerable.Cast<object>().Count() : 1)} items");
                }
                else
                {
                    Logger.Debug($"DataBinder: Failed to extract data for variable '{variableName}'");
                }
            }
        }
    }

    /// <summary>
    /// Extract data for corrected expression by parsing the variable name and navigating the data structure
    /// </summary>
    private object? ExtractDataForCorrectedExpression(object data, string correctedVariableName)
    {
        try
        {
            Logger.Debug($"DataBinder: Extracting data for '{correctedVariableName}'");

            // Parse patterns like "Products__1__Items" 
            var parts = correctedVariableName.Split("__");
            Logger.Debug($"DataBinder: Split into {parts.Length} parts: [{string.Join(", ", parts)}]");

            if (parts.Length < 3)
            {
                Logger.Debug($"DataBinder: Not enough parts for '{correctedVariableName}' (need at least 3)");
                return null;
            }

            var baseName = parts[0]; // "Products"
            var indexStr = parts[1];  // "1"
            var propertyName = parts[2]; // "Items"

            Logger.Debug($"DataBinder: BaseName='{baseName}', Index='{indexStr}', Property='{propertyName}'");

            // Get the base array from data
            var baseProperty = data.GetType().GetProperty(baseName);
            if (baseProperty == null)
            {
                Logger.Debug($"DataBinder: Property '{baseName}' not found in data type {data.GetType().Name}");
                return null;
            }

            var baseValue = baseProperty.GetValue(data);
            if (baseValue is not System.Collections.IEnumerable enumerable || baseValue is string)
            {
                Logger.Debug($"DataBinder: Property '{baseName}' is not enumerable or is string");
                return null;
            }

            var list = enumerable.Cast<object>().ToList();
            Logger.Debug($"DataBinder: Found {list.Count} items in '{baseName}' collection");

            // Parse index
            if (!int.TryParse(indexStr, out var index))
            {
                Logger.Debug($"DataBinder: Cannot parse index '{indexStr}' as integer");
                return null;
            }

            if (index >= list.Count)
            {
                Logger.Debug($"DataBinder: Index {index} is out of bounds for collection with {list.Count} items");
                return null;
            }

            var targetItem = list[index];
            if (targetItem == null)
            {
                Logger.Debug($"DataBinder: Item at index {index} is null");
                return null;
            }

            Logger.Debug($"DataBinder: Found target item at index {index}, type: {targetItem.GetType().Name}");

            // Get the property value from the target item
            var targetProperty = targetItem.GetType().GetProperty(propertyName);
            if (targetProperty == null)
            {
                Logger.Debug($"DataBinder: Property '{propertyName}' not found in target item type {targetItem.GetType().Name}");
                return null;
            }

            var result = targetProperty.GetValue(targetItem);
            Logger.Debug($"DataBinder: Successfully extracted data for '{correctedVariableName}': {result?.GetType().Name ?? "null"}");

            return result;
        }
        catch (Exception ex)
        {
            Logger.Warning($"DataBinder: Error extracting data for corrected expression '{correctedVariableName}': {ex.Message}");
            return null;
        }
    }
}