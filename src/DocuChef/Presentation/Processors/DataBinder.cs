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
    }
    
    /// <summary>
    /// Apply context path transformation with optimized parsing
    /// </summary>
    public void ApplyContextPath(Dictionary<string, object> variables, object data, string contextPath)
    {
        Logger.Debug($"DataBinder: ApplyContextPath called with contextPath='{contextPath}', data type={data?.GetType().Name ?? "null"}");

        if (!contextPath.Contains('>'))
            return;

        object? vData = data;
        var splits = contextPath.Split('>');
        var vNames = new List<string>(splits.Length);

        Logger.Debug($"DataBinder: Processing {splits.Length} splits for contextPath '{contextPath}'");

        foreach (var split in splits)
        {
            if (vData == null)
            {
                Logger.Warning($"DataBinder: vData is null when processing split '{split}' in context path '{contextPath}'");
                break;
            }

            string vName;
            if (TryParseArrayAccess(split, out vName, out var index))
            {
                Logger.Debug($"DataBinder: Processing array access '{split}' -> name='{vName}', index={index}");
                // Include array index in variable name to distinguish Products[0] from Products[1]
                vNames.Add($"{vName}__{index}");
                vData = ResolveValue(vData, vName, index);
                Logger.Debug($"DataBinder: After ResolveValue: vData type={vData?.GetType().Name ?? "null"}");
            }
            else
            {
                Logger.Debug($"DataBinder: Processing simple property '{split}'");
                vNames.Add(split);
                vData = ResolveValue(vData, split, null);
                Logger.Debug($"DataBinder: After ResolveValue: vData type={vData?.GetType().Name ?? "null"}");
            }
        }
        var vNameText = string.Join("__", vNames);
        variables[vNameText] = vData ?? string.Empty;

        Logger.Debug($"DataBinder: Created context variable '{vNameText}' for path '{contextPath}'");        // CRITICAL FIX: For nested contexts like "Brands[0]>Types", we need to create additional variables
        // that can resolve expressions like "${Brands>Types[0].Key}" where "Brands>" gets resolved to current context
        if (splits.Length >= 2 && vData != null)
        {
            // Create a direct mapping from the parent collection to the resolved data
            // e.g., "Brands" -> resolved Types collection for Brands[0]
            var parentCollectionName = splits[0];
            if (TryParseArrayAccess(parentCollectionName, out var cleanParentName, out var parentIndex))
            {
                Logger.Debug($"DataBinder: BEFORE parent mapping - '{cleanParentName}' existing value type: {(variables.ContainsKey(cleanParentName) ? variables[cleanParentName]?.GetType().Name : "not exists")}");

                // Store the original collection if not already stored (but don't expose it in final variables)
                var originalKey = $"{cleanParentName}__Original";
                if (!variables.ContainsKey(originalKey) && variables.ContainsKey(cleanParentName))
                {
                    variables[originalKey] = variables[cleanParentName];
                    Logger.Debug($"DataBinder: Backed up original '{cleanParentName}' collection");
                }

                // Map "Brands" to the Types collection from Brands[0] for context resolution
                variables[cleanParentName] = vData;
                Logger.Debug($"DataBinder: Created parent collection mapping '{cleanParentName}' -> nested collection for context resolution");
                Logger.Debug($"DataBinder: vData type: {vData.GetType().Name}, count: {(vData is System.Collections.IEnumerable enumerable ? enumerable.Cast<object>().Count() : "N/A")}");
            }
        }
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
    /// </summary>
    private string EvaluateTemplate(string dollarSignTemplate, Dictionary<string, object> variables)
    {
        try
        {
            // Filter out __Original variables for evaluation
            var filteredVariables = variables
                .Where(kvp => !kvp.Key.EndsWith("__Original"))
                .ToDictionary(kvp => kvp.Key, kvp => kvp.Value);

            // PERFORMANCE OPTIMIZATION: Pre-validate array indices before DollarSign.Eval
            // Temporarily disable strict validation to allow more flexible array access
            // ValidateArrayIndicesInTemplate(dollarSignTemplate, variables);

            // Setup DollarSignEngine options
            var options = new DollarSignOptions
            {
                SupportDollarSignSyntax = true,
                ThrowOnError = false  // Allow graceful handling of missing values
            };

            var result = DollarSign.Eval(dollarSignTemplate, filteredVariables, options);

            // Check if result contains error message and return empty string instead
            if (result != null && result.Contains("[ERROR:"))
            {
                Logger.Debug($"DataBinder.EvaluateTemplate: Error result detected, returning empty string - {dollarSignTemplate}");
                return string.Empty;
            }

            Logger.Debug($"DataBinder.EvaluateTemplate: {dollarSignTemplate}: {result}");
            return result ?? string.Empty;
        }
        catch (DollarSignEngineException dollarSignEngineException)
        {
            if (dollarSignEngineException.InnerException is ArgumentOutOfRangeException ||
                dollarSignEngineException.InnerException is IndexOutOfRangeException)
            {
                Logger.Debug($"DataBinder.EvaluateTemplate: Array index out of bounds, returning empty string - {dollarSignEngineException.Message}");
                return string.Empty;
            }

            Logger.Debug($"DataBinder.EvaluateTemplate: Not an array bounds error, returning empty string");
            return string.Empty;
        }
        catch (DocuChefHideException hideEx)
        {
            Logger.Debug($"DataBinder.EvaluateTemplate: Array bounds exceeded, returning empty string - {hideEx.Message}");
            return string.Empty;
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
            Logger.Debug($"DataBinder: Original template: '{template}', ContextPath: '{contextPath ?? "null"}'");

            // Adjust context path with index offset for nested collections
            var adjustedContextPath = AdjustContextPathWithOffset(contextPath, indexOffset);
            Logger.Debug($"DataBinder: Adjusted ContextPath: '{adjustedContextPath ?? "null"}' (offset: {indexOffset})");

            // Use the improved variable resolution with context awareness
            var variables = GetVariablesForContext(data, adjustedContextPath);

            // Convert context operators in template to use context-specific variable names
            var dollarSignTemplate = ConvertContextOperatorsToContextAware(template, adjustedContextPath);

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
        if (string.IsNullOrEmpty(contextPath) || !template.Contains('>'))
            return template;

        Logger.Debug($"DataBinder: Converting context operators for contextPath: '{contextPath}'");
        Logger.Debug($"DataBinder: Input template: '{template}'");

        // For nested contexts like "Brands[1]>Types", we need to be more careful
        // The slide offset determines which item in the innermost collection we're processing
        var parts = contextPath.Split('>');
        var result = template;

        if (parts.Length >= 2)
        {
            // Extract parent context index from first part (e.g., "Brands[1]" -> 1)
            var parentPart = parts[0];
            var parentMatch = Regex.Match(parentPart, @"(\w+)\[(\d+)\]");
            if (parentMatch.Success)
            {
                var parentName = parentMatch.Groups[1].Value;
                var parentIndex = parentMatch.Groups[2].Value;

                // Extract current context index dynamically instead of hardcoding 0
                var currentContextIndex = ExtractCurrentContextIndex(contextPath);
                Logger.Debug($"DataBinder: Extracted current context index: {currentContextIndex}");

                // Replace pattern like "Brands__1__Types>Items" with "Brands__1__Types__{currentIndex}__Items"
                var pattern = $@"{parentName}__\d+__(\w+)>(\w+)";
                result = Regex.Replace(result, pattern, $"{parentName}__{parentIndex}__$1__{currentContextIndex}__$2");
                Logger.Debug($"DataBinder: Applied nested pattern replacement with parent index {parentIndex} and current context index {currentContextIndex}");
            }
        }

        Logger.Debug($"DataBinder: After conversion: '{result}'");
        return result;
    }    /// <summary>
         /// Extract current context index from contextPath
         /// This should return the slide offset for the current collection, not the parent index
         /// For nested contexts like "Brands[0]>Types[1]", this should return the Types index (1)
         /// For "Brands[1]>Types", this should return 0 (first Types item for this slide)
         /// </summary>
    private int ExtractCurrentContextIndex(string? contextPath)
    {
        if (string.IsNullOrEmpty(contextPath))
            return 0;

        Logger.Debug($"DataBinder: ExtractCurrentContextIndex - input contextPath: '{contextPath}'");

        // Split by > to get context parts
        var parts = contextPath.Split('>');

        // For nested contexts, we want the index from the last part if it has one
        // Example: "Brands[0]>Types[1]" should return 1 (the Types index)
        // Example: "Brands[1]>Types" should return 0 (default for first Types item)
        var lastPart = parts[parts.Length - 1];
        var match = Regex.Match(lastPart, @"\[(\d+)\]");
        if (match.Success && int.TryParse(match.Groups[1].Value, out var index))
        {
            Logger.Debug($"DataBinder: Extracted current context index {index} from last part '{lastPart}'");
            return index;
        }

        // If no index in last part, this represents the first item (index 0)
        Logger.Debug($"DataBinder: No index found in last part '{lastPart}', using default 0");
        return 0;
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
        if (data == null)
        {
            Logger.Debug($"DataBinder: ResolveValue called with null data for property '{property}'");
            return null;
        }

        Logger.Debug($"DataBinder: ResolveValue - data type: {data.GetType().Name}, property: '{property}', index: {index}");

        if (data is IDictionary dic)
        {
            Logger.Debug($"DataBinder: Processing as IDictionary, contains '{property}': {dic.Contains(property)}");
            if (!dic.Contains(property))
                return null;

            var value = dic[property];

            // IMPORTANT FIX: Check if we have an original backup for this property
            // This handles the case where the current value was overwritten by context processing
            var originalKey = $"{property}__Original";
            if (dic.Contains(originalKey))
            {
                var originalValue = dic[originalKey];
                Logger.Debug($"DataBinder: Found original backup for '{property}', using original instead of current value");
                Logger.Debug($"DataBinder: Original type: {originalValue?.GetType().Name}, Current type: {value?.GetType().Name}");
                value = originalValue;
            }

            if (value is IEnumerable collValue && value is not string)
            {
                var list = collValue.OfType<object>().ToList();
                Logger.Debug($"DataBinder: Dictionary value is enumerable with {list.Count} items, requested index: {index}");
                return index.HasValue && index.Value >= 0 && index.Value < list.Count
                    ? list[index.Value]
                    : null;
            }
            return value;
        }
        else if (data is IEnumerable coll && data is not string)
        {
            var list = coll.OfType<object>().ToList();
            Logger.Debug($"DataBinder: Processing as IEnumerable with {list.Count} items, requested index: {index}");
            return index.HasValue && index.Value >= 0 && index.Value < list.Count
                ? list[index.Value]
                : null;
        }
        else
        {
            var dataType = data.GetType();
            var pInfo = dataType.GetProperty(property);

            // If property not found by name, try to find it in anonymous types by checking all properties
            if (pInfo == null && dataType.Name.Contains("AnonymousType"))
            {
                Logger.Debug($"DataBinder: Searching for property '{property}' in anonymous type {dataType.Name}");
                var properties = dataType.GetProperties();
                Logger.Debug($"DataBinder: Available properties: {string.Join(", ", properties.Select(p => p.Name))}");

                // Try case-insensitive match for anonymous types
                pInfo = properties.FirstOrDefault(p => string.Equals(p.Name, property, StringComparison.OrdinalIgnoreCase));
            }

            if (pInfo != null)
            {
                Logger.Debug($"DataBinder: Found property '{pInfo.Name}' via reflection");
                var result = pInfo.GetValue(data);
                Logger.Debug($"DataBinder: Property value type: {result?.GetType().Name ?? "null"}");
                return result;
            }
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

            // Extract all variable names with __ pattern from the expression
            // This handles both simple cases like "Products__1__Items" and function calls like "ppt.Image(Products__1__Items[0])"
            var variableRegex = new System.Text.RegularExpressions.Regex(@"([A-Za-z_][A-Za-z0-9_]*__[A-Za-z0-9_]*__[A-Za-z0-9_]*)");
            var variableMatches = variableRegex.Matches(expression);

            foreach (System.Text.RegularExpressions.Match varMatch in variableMatches)
            {
                var variableName = varMatch.Groups[1].Value;
                Logger.Debug($"DataBinder: Found variable name: '{variableName}'");

                // Skip if variable already exists
                if (variables.ContainsKey(variableName))
                {
                    Logger.Debug($"DataBinder: Variable '{variableName}' already exists, skipping");
                    continue;
                }

                // Try to extract data for this corrected expression
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

            // Parse patterns like "Brands__0__Types__0__Items" for deeper nesting
            var parts = correctedVariableName.Split("__");
            Logger.Debug($"DataBinder: Split into {parts.Length} parts: [{string.Join(", ", parts)}]");

            if (parts.Length < 3)
            {
                Logger.Debug($"DataBinder: Not enough parts for '{correctedVariableName}' (need at least 3)");
                return null;
            }

            // Handle different patterns:
            // 3 parts: "Brands__0__Types" -> Brands[0].Types
            // 4 parts: "Brands__0__Types__Items" -> Brands[0].Types[0].Items (use index 0 for Items)
            // 5 parts: "Brands__0__Types__0__Items" -> Brands[0].Types[0].Items
            if (parts.Length == 3)
            {
                var baseName = parts[0]; // "Brands"
                var indexStr = parts[1];  // "0"
                var propertyName = parts[2]; // "Types"

                Logger.Debug($"DataBinder: 3-part pattern - BaseName='{baseName}', Index='{indexStr}', Property='{propertyName}'");
                return ExtractSimpleProperty(data, baseName, indexStr, propertyName);
            }
            else if (parts.Length == 4)
            {
                var baseName = parts[0]; // "Brands"
                var baseIndexStr = parts[1];  // "0" or "1"
                var nestedName = parts[2]; // "Types"
                var finalProperty = parts[3]; // "Items"

                Logger.Debug($"DataBinder: 4-part pattern - BaseName='{baseName}', BaseIndex='{baseIndexStr}', NestedName='{nestedName}', FinalProperty='{finalProperty}'");
                // For 4-part pattern, use index 0 for the nested collection (Types[0].Items)
                return ExtractNestedProperty(data, baseName, baseIndexStr, nestedName, "0", finalProperty);
            }
            else if (parts.Length == 5)
            {
                var baseName = parts[0]; // "Brands"
                var baseIndexStr = parts[1];  // "0"
                var nestedName = parts[2]; // "Types"
                var nestedIndexStr = parts[3]; // "0"
                var finalProperty = parts[4]; // "Items"

                Logger.Debug($"DataBinder: 5-part pattern - BaseName='{baseName}', BaseIndex='{baseIndexStr}', NestedName='{nestedName}', NestedIndex='{nestedIndexStr}', FinalProperty='{finalProperty}'");
                return ExtractNestedProperty(data, baseName, baseIndexStr, nestedName, nestedIndexStr, finalProperty);
            }
            else if (parts.Length == 6)
            {
                // Handle deeply nested structure like Brands[0]>Types[1]>Items[1]
                var baseName = parts[0]; // "Brands"
                var baseIndexStr = parts[1];  // "0" 
                var midName = parts[2]; // "Types"
                var midIndexStr = parts[3]; // "1"
                var nestedIndexStr = parts[4]; // "1" (from __1__)
                var finalProperty = parts[5]; // "Items"

                Logger.Debug($"DataBinder: 6-part pattern - BaseName='{baseName}', BaseIndex='{baseIndexStr}', MidName='{midName}', MidIndex='{midIndexStr}', NestedIndex='{nestedIndexStr}', FinalProperty='{finalProperty}'");

                // For 6-part pattern, we need to handle the middle indexing differently
                // This typically comes from context expressions like ${Brands>Types>Items[1]["key"]}
                // where the middle index represents the current context offset
                return ExtractDeeplyNestedProperty(data, baseName, baseIndexStr, midName, midIndexStr, nestedIndexStr, finalProperty);
            }
            else
            {
                Logger.Debug($"DataBinder: Unsupported pattern with {parts.Length} parts");
                return null;
            }
        }
        catch (Exception ex)
        {
            Logger.Debug($"DataBinder: Error extracting data for '{correctedVariableName}': {ex.Message}");
            return null;
        }
    }

    /// <summary>
    /// Extract a property value from an object
    /// </summary>
    private object? ExtractPropertyFromObject(object obj, string propertyName)
    {
        try
        {
            var property = obj.GetType().GetProperty(propertyName);
            if (property == null)
            {
                Logger.Debug($"DataBinder: Property '{propertyName}' not found in type {obj.GetType().Name}");
                return null;
            }

            var result = property.GetValue(obj);
            Logger.Debug($"DataBinder: Extracted '{propertyName}' from {obj.GetType().Name}: {(result != null ? result.GetType().Name : "null")}");
            return result;
        }
        catch (Exception ex)
        {
            Logger.Debug($"DataBinder: Error extracting property '{propertyName}': {ex.Message}");
            return null;
        }
    }

    /// <summary>
    /// Extract a simple property like Brands[0].Types using cached variables when possible
    /// </summary>
    private object? ExtractSimpleProperty(object data, string baseName, string indexStr, string propertyName)
    {
        try
        {
            Logger.Debug($"DataBinder: ExtractSimpleProperty - trying to get {baseName}[{indexStr}].{propertyName}");

            // Try to use cached variables first
            foreach (var cacheEntry in _variableCache)
            {
                var cachedVariables = cacheEntry.Value;

                // Look for a cached variable that matches our base pattern
                var baseVarName = $"{baseName}__{indexStr}";
                if (cachedVariables.ContainsKey(baseVarName))
                {
                    var cachedBaseValue = cachedVariables[baseVarName];
                    Logger.Debug($"DataBinder: Found cached variable '{baseVarName}' for extraction");

                    if (cachedBaseValue != null)
                    {
                        return ExtractPropertyFromObject(cachedBaseValue, propertyName);
                    }
                }

                // Also try the base collection directly
                if (cachedVariables.ContainsKey(baseName))
                {
                    var baseCollection = cachedVariables[baseName];
                    Logger.Debug($"DataBinder: Found cached collection '{baseName}' for extraction");

                    if (baseCollection is System.Collections.IEnumerable enumerable && !(baseCollection is string))
                    {
                        var list = enumerable.Cast<object>().ToList();
                        if (int.TryParse(indexStr, out var index) && index < list.Count)
                        {
                            var item = list[index];
                            if (item != null)
                            {
                                return ExtractPropertyFromObject(item, propertyName);
                            }
                        }
                    }
                }
            }

            // Fallback to original approach if cached variables don't work
            var baseProperty = data.GetType().GetProperty(baseName);
            if (baseProperty == null)
            {
                Logger.Debug($"DataBinder: Property '{baseName}' not found in data type {data.GetType().Name}");
                return null;
            }

            var baseValue = baseProperty.GetValue(data);
            if (baseValue is not System.Collections.IEnumerable enumerable2 || baseValue is string)
            {
                Logger.Debug($"DataBinder: Property '{baseName}' is not enumerable or is string");
                return null;
            }

            var list2 = enumerable2.Cast<object>().ToList();
            Logger.Debug($"DataBinder: Found {list2.Count} items in '{baseName}' collection");

            // Parse index
            if (!int.TryParse(indexStr, out var index2))
            {
                Logger.Debug($"DataBinder: Cannot parse index '{indexStr}' as integer");
                return null;
            }

            if (index2 >= list2.Count)
            {
                Logger.Debug($"DataBinder: Index {index2} is out of bounds for collection size {list2.Count}");
                return null;
            }

            var item2 = list2[index2];
            if (item2 == null)
            {
                Logger.Debug($"DataBinder: Item at index {index2} is null");
                return null;
            }

            return ExtractPropertyFromObject(item2, propertyName);
        }
        catch (Exception ex)
        {
            Logger.Debug($"DataBinder: Error extracting simple property '{baseName}[{indexStr}].{propertyName}': {ex.Message}");
            return null;
        }
    }

    /// <summary>
    /// Extract a nested property like Brands[0].Types[0].Items using cached variables when possible
    /// </summary>
    private object? ExtractNestedProperty(object data, string baseName, string baseIndexStr, string nestedName, string nestedIndexStr, string finalProperty)
    {
        try
        {
            Logger.Debug($"DataBinder: ExtractNestedProperty - trying to get {baseName}[{baseIndexStr}].{nestedName}[{nestedIndexStr}].{finalProperty}");

            // Try to use cached variables first for better performance
            foreach (var cacheEntry in _variableCache)
            {
                var cachedVariables = cacheEntry.Value;

                // Look for a cached nested variable that matches our pattern
                var nestedVarName = $"{baseName}__{baseIndexStr}__{nestedName}";
                if (cachedVariables.ContainsKey(nestedVarName))
                {
                    var nestedCollection = cachedVariables[nestedVarName];
                    Logger.Debug($"DataBinder: Found cached nested variable '{nestedVarName}' for extraction");

                    if (nestedCollection is System.Collections.IEnumerable enumerable && !(nestedCollection is string))
                    {
                        var list = enumerable.Cast<object>().ToList();
                        Logger.Debug($"DataBinder: Cached nested collection has {list.Count} items");

                        if (int.TryParse(nestedIndexStr, out var nestedIndex) && nestedIndex < list.Count)
                        {
                            var nestedItem = list[nestedIndex];
                            if (nestedItem != null)
                            {
                                Logger.Debug($"DataBinder: Successfully found nested item at index {nestedIndex}");
                                return ExtractPropertyFromObject(nestedItem, finalProperty);
                            }
                        }
                        else
                        {
                            Logger.Debug($"DataBinder: Nested index {nestedIndexStr} is out of bounds or invalid for collection with {list.Count} items");
                        }
                    }
                }
            }

            // Fallback: First, get the base object: Brands[0]
            var baseObject = ExtractSimpleProperty(data, baseName, baseIndexStr, nestedName);
            if (baseObject == null)
            {
                Logger.Debug($"DataBinder: Failed to get base object for '{baseName}[{baseIndexStr}].{nestedName}'");
                return null;
            }

            // Then, get the nested property from that object: Types[0].Items
            if (baseObject is not System.Collections.IEnumerable enumerable2 || baseObject is string)
            {
                Logger.Debug($"DataBinder: Nested property '{nestedName}' is not enumerable");
                return null;
            }

            var list2 = enumerable2.Cast<object>().ToList();
            Logger.Debug($"DataBinder: Found {list2.Count} items in nested '{nestedName}' collection");

            // Parse nested index
            if (!int.TryParse(nestedIndexStr, out var nestedIndex2))
            {
                Logger.Debug($"DataBinder: Cannot parse nested index '{nestedIndexStr}' as integer");
                return null;
            }

            if (nestedIndex2 >= list2.Count)
            {
                Logger.Debug($"DataBinder: Nested index {nestedIndex2} is out of bounds for collection size {list2.Count}");
                return null;
            }

            var nestedItem2 = list2[nestedIndex2];
            if (nestedItem2 == null)
            {
                Logger.Debug($"DataBinder: Nested item at index {nestedIndex2} is null");
                return null;
            }

            return ExtractPropertyFromObject(nestedItem2, finalProperty);
        }
        catch (Exception ex)
        {
            Logger.Debug($"DataBinder: Error extracting nested property '{baseName}[{baseIndexStr}].{nestedName}[{nestedIndexStr}].{finalProperty}': {ex.Message}");
            return null;
        }
    }

    /// <summary>
    /// Extract a deeply nested property like Brands[0].Types[1].__1__.Items for 6-part patterns
    /// </summary>
    private object? ExtractDeeplyNestedProperty(object data, string baseName, string baseIndexStr, string midName, string midIndexStr, string nestedIndexStr, string finalProperty)
    {
        try
        {
            Logger.Debug($"DataBinder: ExtractDeeplyNestedProperty - trying to get {baseName}[{baseIndexStr}].{midName}[{midIndexStr}].__{nestedIndexStr}__{finalProperty}");

            // For 6-part patterns, the structure is typically:
            // Brands[0] > Types[1] > __1__ > Items
            // This means we want the Items from the nested context at index 1

            // First get the base collection: Brands[baseIndex].midName
            var baseCollection = ExtractSimpleProperty(data, baseName, baseIndexStr, midName);
            if (baseCollection == null)
            {
                Logger.Debug($"DataBinder: Failed to get base collection for '{baseName}[{baseIndexStr}].{midName}'");
                return null;
            }

            // Convert to enumerable and get the item at midIndex
            if (baseCollection is not System.Collections.IEnumerable enumerable || baseCollection is string)
            {
                Logger.Debug($"DataBinder: Base collection '{midName}' is not enumerable");
                return null;
            }

            var list = enumerable.Cast<object>().ToList();
            Logger.Debug($"DataBinder: Found {list.Count} items in '{midName}' collection");

            if (!int.TryParse(midIndexStr, out var midIndex))
            {
                Logger.Debug($"DataBinder: Cannot parse mid index '{midIndexStr}' as integer");
                return null;
            }

            if (midIndex >= list.Count)
            {
                Logger.Debug($"DataBinder: Mid index {midIndex} is out of bounds for collection size {list.Count}");
                return null;
            }

            var midItem = list[midIndex];
            if (midItem == null)
            {
                Logger.Debug($"DataBinder: Mid item at index {midIndex} is null");
                return null;
            }

            // Now get the final property from that item
            return ExtractPropertyFromObject(midItem, finalProperty);
        }
        catch (Exception ex)
        {
            Logger.Debug($"DataBinder: Error extracting deeply nested property '{baseName}[{baseIndexStr}].{midName}[{midIndexStr}].__{nestedIndexStr}__{finalProperty}': {ex.Message}");
            return null;
        }
    }

    /// <summary>
    /// Adjust context path by applying index offset to the last collection in the path
    /// For example: "Brands[0]>Types" with offset 1 becomes "Brands[0]>Types[1]"
    /// </summary>
    private string? AdjustContextPathWithOffset(string? contextPath, int indexOffset)
    {
        if (string.IsNullOrEmpty(contextPath) || indexOffset <= 0)
            return contextPath;

        Logger.Debug($"DataBinder: AdjustContextPathWithOffset - input: '{contextPath}', offset: {indexOffset}");

        var parts = contextPath.Split('>');
        if (parts.Length == 0)
            return contextPath;

        // Apply offset to the last part (collection name)
        var lastPart = parts[parts.Length - 1];

        // If the last part doesn't have an index, add one with the offset
        if (!lastPart.Contains('['))
        {
            parts[parts.Length - 1] = $"{lastPart}[{indexOffset}]";
        }
        else
        {
            // If it already has an index, adjust it by the offset
            var match = System.Text.RegularExpressions.Regex.Match(lastPart, @"^([^[]+)\[(\d+)\](.*)$");
            if (match.Success)
            {
                var baseName = match.Groups[1].Value;
                var currentIndex = int.Parse(match.Groups[2].Value);
                var suffix = match.Groups[3].Value;
                var newIndex = currentIndex + indexOffset;
                parts[parts.Length - 1] = $"{baseName}[{newIndex}]{suffix}";
            }
        }

        var result = string.Join(">", parts);
        Logger.Debug($"DataBinder: AdjustContextPathWithOffset - result: '{result}'");
        return result;
    }
}