using DollarSignEngine;
using System.Globalization;

namespace DocuChef.PowerPoint.DollarSignEngine;

/// <summary>
/// Improved DollarSignEngine adapter for processing expressions in PowerPoint templates
/// with enhanced nested data structure support
/// </summary>
internal class ExpressionEvaluator : IExpressionEvaluator
{
    protected readonly DollarSignOptions _options;
    private readonly CultureInfo _cultureInfo;
    private readonly PowerPointContext _context;

    /// <summary>
    /// Initializes a new instance of ExpressionEvaluator
    /// </summary>
    public ExpressionEvaluator(CultureInfo cultureInfo = null)
    {
        _cultureInfo = cultureInfo ?? CultureInfo.CurrentCulture;
        _context = null;

        _options = new DollarSignOptions
        {
            CultureInfo = _cultureInfo,
            SupportDollarSignSyntax = true,  // Use ${variable} syntax as per PPT guidelines
            ThrowOnMissingParameter = false, // Don't throw on missing params, show placeholder instead
            EnableDebugLogging = false,      // Disable debug logging by default
            VariableResolver = CustomVariableResolver // Custom resolver for PowerPoint functions and array indices
        };
    }

    /// <summary>
    /// Initializes a new instance of ExpressionEvaluator with context
    /// </summary>
    public ExpressionEvaluator(PowerPointContext context, CultureInfo cultureInfo = null)
        : this(cultureInfo)
    {
        _context = context;
    }

    /// <summary>
    /// Evaluates a complete expression with provided variables
    /// </summary>
    public object EvaluateCompleteExpression(string expression, Dictionary<string, object> variables)
    {
        try
        {
            // If already wrapped in ${...}, evaluate directly
            if (expression.StartsWith("${") && expression.EndsWith("}"))
            {
                var result = DollarSign.EvalAsync(expression, variables, _options).GetAwaiter().GetResult();
                return result;
            }

            // Otherwise, wrap it for evaluation
            string wrappedExpr = "${" + expression + "}";
            var evalResult = DollarSign.EvalAsync(wrappedExpr, variables, _options).GetAwaiter().GetResult();
            return evalResult;
        }
        catch (Exception ex)
        {
            Logger.Error($"Error evaluating expression '{expression}': {ex.Message}", ex);
            return $"[Error: {ex.Message}]";
        }
    }

    /// <summary>
    /// Enhanced custom variable resolver for PowerPoint functions, array indices, and nested data structures
    /// </summary>
    protected virtual object CustomVariableResolver(string expression, object parameters)
    {
        // Convert parameters to dictionary if not already
        Dictionary<string, object> variables;
        if (parameters is Dictionary<string, object> dict)
        {
            variables = dict;
        }
        else
        {
            variables = new Dictionary<string, object>();

            // If parameters is an object, extract its properties
            if (parameters != null)
            {
                var props = parameters.GetType().GetProperties();
                foreach (var prop in props)
                {
                    if (prop.CanRead)
                    {
                        try
                        {
                            var value = prop.GetValue(parameters);
                            variables[prop.Name] = value;
                        }
                        catch
                        {
                            // Skip properties that throw exceptions
                        }
                    }
                }
            }
        }

        // Handle PowerPoint functions (ppt.)
        if (expression.StartsWith("ppt."))
        {
            return HandlePowerPointFunction(expression, variables);
        }

        // First try to use hierarchical path resolution if context is available
        if (_context?.Navigator != null &&
            (expression.Contains('.') || expression.Contains('[') || expression.Contains('_')))
        {
            // For complex paths, use hierarchical path navigation
            var hierarchicalPath = new HierarchicalPath(expression);
            if (hierarchicalPath.Segments.Count > 0)
            {
                // Use the current hierarchical indices from context
                var result = _context.Navigator.ResolveValueWithContext(hierarchicalPath, _context.HierarchicalIndices);
                if (result != null)
                {
                    Logger.Debug($"[EVAL] Resolved hierarchical path: {expression}");
                    return result;
                }
            }
        }

        // Handle dotted path notation (obj.prop.nested)
        if (expression.Contains('.'))
        {
            string[] parts = expression.Split('.');
            if (parts.Length > 1 && variables.TryGetValue(parts[0], out var rootObj) && rootObj != null)
            {
                return ResolveNestedProperty(rootObj, parts, 1);
            }
        }

        // New approach: Handle nested data patterns with multiple levels (A_B_C)
        if (expression.Contains('_') && !expression.StartsWith("_"))
        {
            var result = ResolveMultiLevelNestedData(expression, variables);
            if (result != null)
            {
                Logger.Debug($"[EVAL] Resolved multi-level nested data: {expression}");
                return result;
            }
        }

        // Handle array indexing with multiple dimensions: collections[0].subcollections[1].items[2].property
        var complexIndexerMatch = System.Text.RegularExpressions.Regex.Match(
            expression, @"^([\w]+)(?:\[(\d+)\])((?:\.\w+(?:\[\d+\])?)+)$");

        if (complexIndexerMatch.Success)
        {
            return ResolveComplexIndexerPath(complexIndexerMatch, variables);
        }

        // Handle simple array indexing expressions: item[0].Property and Items[0].Property
        var arrayIndexMatch = System.Text.RegularExpressions.Regex.Match(
            expression, @"^([\w]+)\[(\d+)\](\.(\w+))?$");

        if (arrayIndexMatch.Success)
        {
            return ResolveSimpleArrayIndex(arrayIndexMatch, variables);
        }

        // For direct variable access, let DollarSignEngine handle it
        if (variables.TryGetValue(expression, out var directValue))
        {
            return directValue;
        }

        // For other expressions, let DollarSignEngine handle them
        return null;
    }

    /// <summary>
    /// Resolve nested properties recursively
    /// </summary>
    private object ResolveNestedProperty(object obj, string[] parts, int startIndex)
    {
        if (obj == null || startIndex >= parts.Length)
            return null;

        var currentObj = obj;

        // Navigate through the property chain
        for (int i = startIndex; i < parts.Length; i++)
        {
            string propName = parts[i];

            // Check if property has array index
            var indexMatch = System.Text.RegularExpressions.Regex.Match(propName, @"(\w+)\[(\d+)\]");
            if (indexMatch.Success)
            {
                propName = indexMatch.Groups[1].Value;
                int index = int.Parse(indexMatch.Groups[2].Value);

                // Get the property (which should be a collection)
                var prop = currentObj.GetType().GetProperty(propName);
                if (prop == null)
                    return null;

                var collection = prop.GetValue(currentObj);
                if (collection == null)
                    return null;

                // Get item at index
                currentObj = GetItemAtIndex(collection, index);
                if (currentObj == null)
                    return null;
            }
            else
            {
                // Simple property access
                var prop = currentObj.GetType().GetProperty(propName);
                if (prop == null)
                    return null;

                currentObj = prop.GetValue(currentObj);
                if (currentObj == null)
                    return null;
            }
        }

        return currentObj;
    }

    /// <summary>
    /// Resolves multi-level nested data references (e.g., Parent_Child_Grandchild)
    /// </summary>
    private object ResolveMultiLevelNestedData(string expression, Dictionary<string, object> variables)
    {
        // Split by underscore
        var parts = expression.Split('_');
        if (parts.Length < 2)
            return null;

        // Look for direct variable with this name first
        if (variables.TryGetValue(expression, out var directValue))
        {
            Logger.Debug($"[EVAL] Found direct variable for '{expression}'");
            return directValue;
        }

        // Start with the first part as the root collection
        string rootName = parts[0];
        if (!variables.TryGetValue(rootName, out var rootObj) || rootObj == null)
        {
            Logger.Debug($"[EVAL] Root collection not found: {rootName}");
            return null;
        }

        // Get the current index for the root collection
        int rootIndex = _context?.HierarchicalIndices?.GetValueOrDefault(rootName, 0) ?? 0;
        Logger.Debug($"[EVAL] Using index {rootIndex} for root collection {rootName}");

        // Get item from root collection at current index
        object currentObj = GetItemAtIndex(rootObj, rootIndex);
        if (currentObj == null)
        {
            Logger.Debug($"[EVAL] Item at index {rootIndex} not found in collection {rootName}");
            return null;
        }

        // Process each level of nesting
        for (int i = 1; i < parts.Length - 1; i++)
        {
            string collectionName = parts[i];

            // Get property from current object
            var property = currentObj.GetType().GetProperty(collectionName);
            if (property == null)
            {
                Logger.Debug($"[EVAL] Property {collectionName} not found on object type {currentObj.GetType().Name}");
                return null;
            }

            // Get collection from property
            var collection = property.GetValue(currentObj);
            if (collection == null)
            {
                Logger.Debug($"[EVAL] Collection {collectionName} is null");
                return null;
            }

            // Build the path so far for indexing purposes
            string pathSoFar = string.Join("_", parts.Take(i + 1));

            // Get index for this level
            int index = _context?.HierarchicalIndices?.GetValueOrDefault(pathSoFar, 0) ?? 0;
            Logger.Debug($"[EVAL] Using index {index} for collection {pathSoFar}");

            // Get item at index
            currentObj = GetItemAtIndex(collection, index);
            if (currentObj == null)
            {
                Logger.Debug($"[EVAL] Item at index {index} not found in collection {pathSoFar}");
                return null;
            }
        }

        // Get final property value
        string finalProperty = parts[parts.Length - 1];
        var finalProp = currentObj.GetType().GetProperty(finalProperty);
        if (finalProp == null)
        {
            Logger.Debug($"[EVAL] Final property {finalProperty} not found on object type {currentObj.GetType().Name}");
            return null;
        }

        var result = finalProp.GetValue(currentObj);
        Logger.Debug($"[EVAL] Resolved nested data {expression} to value: {result}");
        return result;
    }

    /// <summary>
    /// Resolves a complex indexer path (collection[0].subcollection[1].property)
    /// </summary>
    private object ResolveComplexIndexerPath(System.Text.RegularExpressions.Match match, Dictionary<string, object> variables)
    {
        string rootName = match.Groups[1].Value;
        int rootIndex = int.Parse(match.Groups[2].Value);
        string restOfPath = match.Groups[3].Value.TrimStart('.');

        Logger.Debug($"[EVAL] Resolving complex indexer path: {rootName}[{rootIndex}].{restOfPath}");

        // Get root collection
        if (!variables.TryGetValue(rootName, out var rootCollection) || rootCollection == null)
        {
            Logger.Debug($"[EVAL] Root collection not found: {rootName}");
            return null;
        }

        // Get item at root index
        object currentObj = GetItemAtIndex(rootCollection, rootIndex);
        if (currentObj == null)
        {
            Logger.Debug($"[EVAL] Item at index {rootIndex} not found in root collection {rootName}");
            return null;
        }

        // Process each path segment
        var segments = restOfPath.Split('.');
        foreach (var segment in segments)
        {
            // Check if segment contains an indexer
            var indexerMatch = System.Text.RegularExpressions.Regex.Match(segment, @"^(\w+)\[(\d+)\]$");
            if (indexerMatch.Success)
            {
                // It's a collection property with index
                string propName = indexerMatch.Groups[1].Value;
                int index = int.Parse(indexerMatch.Groups[2].Value);

                // Get property
                var property = currentObj.GetType().GetProperty(propName);
                if (property == null)
                {
                    Logger.Debug($"[EVAL] Property {propName} not found on object type {currentObj.GetType().Name}");
                    return null;
                }

                // Get collection
                var collection = property.GetValue(currentObj);
                if (collection == null)
                {
                    Logger.Debug($"[EVAL] Collection {propName} is null");
                    return null;
                }

                // Get item at index
                currentObj = GetItemAtIndex(collection, index);
                if (currentObj == null)
                {
                    Logger.Debug($"[EVAL] Item at index {index} not found in collection {propName}");
                    return null;
                }
            }
            else
            {
                // It's a simple property
                var property = currentObj.GetType().GetProperty(segment);
                if (property == null)
                {
                    Logger.Debug($"[EVAL] Property {segment} not found on object type {currentObj.GetType().Name}");
                    return null;
                }

                currentObj = property.GetValue(currentObj);
                if (currentObj == null)
                {
                    Logger.Debug($"[EVAL] Property {segment} is null");
                    return null;
                }
            }
        }

        Logger.Debug($"[EVAL] Resolved complex indexer path to value: {currentObj}");
        return currentObj;
    }

    /// <summary>
    /// Resolves a simple array index expression (Items[0].Property)
    /// </summary>
    private object ResolveSimpleArrayIndex(System.Text.RegularExpressions.Match match, Dictionary<string, object> variables)
    {
        string arrayName = match.Groups[1].Value;
        int index = int.Parse(match.Groups[2].Value);
        string propPath = match.Groups.Count > 3 ? match.Groups[3].Value : null;

        Logger.Debug($"[EVAL] Array expression: arrayName={arrayName}, index={index}, propPath={propPath}");

        // First check if this is a case-insensitive match for "Items"
        string normalizedName = null;
        if (string.Equals(arrayName, "item", StringComparison.OrdinalIgnoreCase))
        {
            normalizedName = "Items";
        }

        // Try both original and normalized names
        object arrayObj = null;
        if (normalizedName != null && variables.TryGetValue(normalizedName, out var normalizedObj))
        {
            arrayObj = normalizedObj;
            arrayName = normalizedName;
        }
        else if (variables.TryGetValue(arrayName, out var originalObj))
        {
            arrayObj = originalObj;
        }

        if (arrayObj != null)
        {
            Logger.Debug($"[EVAL] Found array '{arrayName}' in variables");

            // Get item at index
            object item = GetItemAtIndex(arrayObj, index);

            if (item == null)
            {
                Logger.Warning($"[EVAL] Item at index {index} not found in array {arrayName}");
                return null;
            }

            Logger.Debug($"[EVAL] Retrieved item at index {index}: {item?.GetType().Name}");

            // Return the item or its property
            if (string.IsNullOrEmpty(propPath))
                return item;

            // Remove leading dot
            if (propPath.StartsWith("."))
                propPath = propPath.Substring(1);

            // Get property value
            var property = item?.GetType().GetProperty(propPath);
            if (property != null)
            {
                var propValue = property.GetValue(item);
                Logger.Debug($"[EVAL] Property '{propPath}' value: {propValue}");
                return propValue;
            }
            else
            {
                Logger.Warning($"[EVAL] Property '{propPath}' not found on item type {item?.GetType().Name}");
            }
        }
        else
        {
            Logger.Warning($"[EVAL] Array '{arrayName}' not found in variables");
        }

        return null;
    }

    /// <summary>
    /// Gets an item at a specific index from any collection type
    /// </summary>
    private object GetItemAtIndex(object collection, int index)
    {
        // Array
        if (collection is Array array && index >= 0 && index < array.Length)
        {
            return array.GetValue(index);
        }

        // IList
        if (collection is IList list && index >= 0 && index < list.Count)
        {
            return list[index];
        }

        // IEnumerable
        if (collection is IEnumerable enumerable && !(collection is string))
        {
            int currentIndex = 0;
            foreach (var item in enumerable)
            {
                if (currentIndex == index)
                    return item;
                currentIndex++;
            }
        }

        return null;
    }

    /// <summary>
    /// Handles PowerPoint specific functions (ppt.Image, ppt.Chart, ppt.Table)
    /// </summary>
    private object HandlePowerPointFunction(string expression, Dictionary<string, object> variables)
    {
        // Parse function expression: ppt.Function("arg", param1: value1, param2: value2)
        var match = System.Text.RegularExpressions.Regex.Match(expression, @"ppt\.(\w+)\((.+)??\)");
        if (!match.Success)
        {
            Logger.Warning($"Invalid PowerPoint function format: {expression}");
            return $"[Invalid function: {expression}]";
        }

        string functionName = match.Groups[1].Value;
        string argsString = match.Groups[2].Success ? match.Groups[2].Value : "";

        // Parse arguments
        string[] args = ParseFunctionArguments(argsString);

        // Look up the PowerPoint function
        if (variables.TryGetValue($"ppt.{functionName}", out var funcObj) &&
            funcObj is PowerPointFunction function)
        {
            // Look up PowerPoint context
            PowerPointContext context = null;
            if (variables.TryGetValue("_context", out var ctxObj) &&
                ctxObj is PowerPointContext ctx)
            {
                context = ctx;
            }
            else
            {
                // If no context (rare case), create a temporary one
                context = new PowerPointContext
                {
                    Variables = new Dictionary<string, object>(variables)
                };
            }

            try
            {
                // Execute the function
                return function.Execute(context, null, args);
            }
            catch (Exception ex)
            {
                Logger.Error($"Error executing function {functionName}: {ex.Message}", ex);
                return $"[Error in {functionName}: {ex.Message}]";
            }
        }

        return $"[Unknown function: ppt.{functionName}]";
    }

    /// <summary>
    /// Simple parsing of function arguments
    /// </summary>
    private string[] ParseFunctionArguments(string argsString)
    {
        if (string.IsNullOrEmpty(argsString))
            return Array.Empty<string>();

        var args = new List<string>();
        bool inQuotes = false;
        int startIndex = 0;

        for (int i = 0; i < argsString.Length; i++)
        {
            char c = argsString[i];

            if (c == '"' && (i == 0 || argsString[i - 1] != '\\'))
            {
                inQuotes = !inQuotes;
            }
            else if (c == ',' && !inQuotes)
            {
                args.Add(argsString.Substring(startIndex, i - startIndex).Trim());
                startIndex = i + 1;
            }
        }

        // Add the last argument
        if (startIndex < argsString.Length)
        {
            args.Add(argsString.Substring(startIndex).Trim());
        }

        // Clean up arguments
        for (int i = 0; i < args.Count; i++)
        {
            var arg = args[i].Trim();

            // Remove quotes from string arguments
            if (arg.StartsWith("\"") && arg.EndsWith("\"") && arg.Length > 1)
            {
                arg = arg.Substring(1, arg.Length - 2)
                    .Replace("\\\"", "\"")
                    .Replace("\\\\", "\\");
            }

            args[i] = arg;
        }

        return args.ToArray();
    }
}