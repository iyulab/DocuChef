using System.Reflection;
using DocuChef.Presentation.Models;
using DocuChef.Presentation.Utilities;
using DollarSignEngine;
using System.Globalization;

namespace DocuChef.Presentation.Processors;

/// <summary>
/// Generates a plan for slide creation based on the template analysis and data
/// </summary>
public class SlidePlanGenerator
{
    /// <summary>
    /// Generates a slide plan based on slide infos and data
    /// </summary>
    /// <param name="slideInfos">The analyzed slide information</param>
    /// <param name="data">The data object</param>
    /// <returns>A SlidePlan containing slide instances to be created</returns>
    public SlidePlan GeneratePlan(List<SlideInfo> slideInfos, object data)
    {
        if (slideInfos == null || slideInfos.Count == 0 || data == null)
            return new SlidePlan();

        var slidePlan = new SlidePlan();
        var aliasMap = BuildAliasMap(slideInfos);

        // Process slides with range directives
        ProcessRangeDirectives(slideInfos, slidePlan, data, aliasMap);

        // Process slides with foreach directives that are not part of a range
        ProcessStandaloneSlides(slideInfos, slidePlan, data, aliasMap);

        return slidePlan;
    }

    /// <summary>
    /// Calculates the number of slides required to display all items
    /// </summary>
    public int CalculateRequiredSlides(int itemCount, int itemsPerSlide)
    {
        if (itemCount == 0 || itemsPerSlide <= 0) return 0;
        return (int)Math.Ceiling((double)itemCount / itemsPerSlide);
    }

    /// <summary>
    /// Builds a map of aliases for simpler path resolution
    /// </summary>
    private Dictionary<string, string> BuildAliasMap(List<SlideInfo> slideInfos)
    {
        var aliasMap = new Dictionary<string, string>();
        
        foreach (var slideInfo in slideInfos)
        {
            foreach (var directive in slideInfo.Directives)
            {
                if (directive.Type == DirectiveType.Alias && !string.IsNullOrEmpty(directive.AliasName))
                {
                    aliasMap[directive.AliasName] = directive.SourcePath;
                }
            }
        }
        
        return aliasMap;
    }

    /// <summary>
    /// Processes slides with range directives to generate slide instances
    /// </summary>
    private void ProcessRangeDirectives(List<SlideInfo> slideInfos, SlidePlan slidePlan, object data, Dictionary<string, string> aliasMap)
    {
        // Find all range begin directives
        var rangeBeginSlides = slideInfos
            .Where(si => si.Type == SlideType.Source)
            .ToList();

        foreach (var beginSlide in rangeBeginSlides)
        {
            // Get range directives
            var rangeDirectives = beginSlide.Directives
                .Where(d => d.Type == DirectiveType.Range && d.RangeBoundary == RangeBoundary.Begin)
                .ToList();

            foreach (var rangeDirective in rangeDirectives)
            {
                // Find corresponding range end slide
                var endSlide = slideInfos.FirstOrDefault(si => 
                    si.Type == SlideType.Cloned && 
                    si.Directives.Any(d => d.Type == DirectiveType.Range && 
                                         d.RangeBoundary == RangeBoundary.End && 
                                         d.SourcePath == rangeDirective.SourcePath));

                // Get all slides between range begin and end
                int endIndex = endSlide != null 
                    ? slideInfos.IndexOf(endSlide) 
                    : slideInfos.Count - 1;
                    
                int beginIndex = slideInfos.IndexOf(beginSlide);
                
                var slidesInRange = slideInfos
                    .Skip(beginIndex)
                    .Take(endIndex - beginIndex + 1)
                    .ToList();

                // Process the range
                ProcessRangeWithContextPath(slidesInRange, slidePlan, data, rangeDirective.SourcePath, aliasMap);
            }
        }
    }    /// <summary>
    /// Processes a range of slides with a specific context path
    /// </summary>
    private void ProcessRangeWithContextPath(List<SlideInfo> slidesInRange, SlidePlan slidePlan, object data, string contextPath, Dictionary<string, string> aliasMap)
    {
        // Resolve the collection from the data object
        IEnumerable<object>? collection = ResolveCollection(data, contextPath);
        if (collection == null || !collection.Any()) return; // 컬렉션이 비어있으면 처리하지 않음

        int position = 0;
        foreach (var item in collection)
        {
            string itemContext = contextPath;
            
            // Process each slide in the range for this item
            foreach (var slideInfo in slidesInRange)
            {
                // Skip range boundary slides
                if (slideInfo.Type == SlideType.Cloned)
                    continue;

                // Find foreach directives in this slide
                var foreachDirectives = slideInfo.Directives
                    .Where(d => d.Type == DirectiveType.Foreach)
                    .ToList();                // Process nested collections if any
                if (foreachDirectives.Any(d => d.SourcePath.StartsWith(contextPath + ">")))
                {
                    ProcessNestedCollections(slideInfo, slidePlan, item, position, itemContext ?? "", aliasMap);
                }
                else
                {
                    // Add this slide to the plan
                    slidePlan.AddSlideInstance(new SlideInstance                    {
                        SourceSlideId = slideInfo.SlideId,
                        Position = GetNextPosition(slidePlan),
                        IndexOffset = 0,
                        Type = SlideInstanceType.Generated,
                        ContextPath = itemContext?.Split('>').ToList() ?? new List<string>()
                    });
                }
            }
            
            position++;
        }
    }

    /// <summary>
    /// Processes nested collections within a slide
    /// </summary>
    private void ProcessNestedCollections(SlideInfo slideInfo, SlidePlan slidePlan, object contextItem, int parentOffset, string parentContext, Dictionary<string, string> aliasMap)
    {
        // Get foreach directives that reference nested collections
        var nestedDirectives = slideInfo.Directives
            .Where(d => d.Type == DirectiveType.Foreach && d.SourcePath.Contains(">"))
            .ToList();

        if (nestedDirectives.Count == 0) return;

        foreach (var directive in nestedDirectives)
        {
            // Resolve the nested path relative to the parent context
            string nestedPath = directive.SourcePath.Replace(parentContext + ">", "");
            
            // Resolve the nested collection
            IEnumerable<object>? nestedCollection = ResolveCollection(contextItem, nestedPath);
            if (nestedCollection == null) continue;

            // Calculate how many items per slide
            int itemsPerSlide = directive.MaxItems > 0 ? directive.MaxItems : 1;
            
            // Calculate how many slides we need
            int itemCount = nestedCollection.Count();
            int requiredSlides = CalculateRequiredSlides(itemCount, itemsPerSlide);
            
            // Create slide instances for each required slide
            for (int i = 0; i < requiredSlides; i++)            {
                int offset = i * itemsPerSlide;
                  slidePlan.AddSlideInstance(new SlideInstance
                {
                    SourceSlideId = slideInfo.SlideId,
                    Position = GetNextPosition(slidePlan),
                    IndexOffset = offset,
                    Type = SlideInstanceType.Generated,
                    ContextPath = (parentContext + ">" + nestedPath).Split('>').ToList()
                });
            }
        }
    }    /// <summary>
    /// Processes standalone slides that are not part of a range
    /// </summary>
    private void ProcessStandaloneSlides(List<SlideInfo> slideInfos, SlidePlan slidePlan, object data, Dictionary<string, string> aliasMap)
    {        // Get slides that are not part of a range (static or source)
        var standaloneSlides = slideInfos
            .Where(si => si.Type == SlideType.Static || si.Type == SlideType.Source)
            .ToList();

        foreach (var slideInfo in standaloneSlides)
        {
            var foreachDirectives = slideInfo.Directives
                .Where(d => d.Type == DirectiveType.Foreach)
                .ToList();            if (foreachDirectives.Any())
            {
                // This slide has foreach directives, we need to create clones
                foreach (var directive in foreachDirectives)
                {
                    ProcessForeachDirective(slideInfo, slidePlan, data, directive, aliasMap);
                }
            }
            else if (slideInfo.Type == SlideType.Source && !string.IsNullOrEmpty(slideInfo.CollectionName))
            {
                // This is a source slide with collection binding - infer foreach directive
                var inferredDirective = new Directive
                {
                    Type = DirectiveType.Foreach,
                    SourcePath = slideInfo.CollectionName,
                    MaxItems = slideInfo.MaxArrayIndex + 1
                };
                ProcessForeachDirective(slideInfo, slidePlan, data, inferredDirective, aliasMap);
            }
            else
            {
                // This is a static slide OR a source slide without array binding
                // Design-centered approach: prioritize template design over data structure
                slidePlan.AddSlideInstance(new SlideInstance
                {
                    SourceSlideId = slideInfo.SlideId,
                    Position = GetNextPosition(slidePlan),
                    IndexOffset = 0,
                    Type = SlideInstanceType.Static // Always treat as static when no directives or collection binding
                });
            }
        }
    }

    /// <summary>
    /// Processes a foreach directive to generate slide instances
    /// </summary>
    private void ProcessForeachDirective(SlideInfo slideInfo, SlidePlan slidePlan, object data, Directive directive, Dictionary<string, string> aliasMap)
    {
        // Resolve alias if used
        string resolvedPath = ResolveAliasPath(directive.SourcePath, aliasMap);
          // Resolve the collection
        IEnumerable<object>? collection = ResolveCollection(data, resolvedPath);
        if (collection == null) return;
        
        // Calculate items per slide and offset
        int itemsPerSlide = directive.MaxItems > 0 ? directive.MaxItems : 1;
        int initialOffset = directive.Offset;
        
        // Calculate number of required slides
        int itemCount = collection.Count();
        int requiredSlides = CalculateRequiredSlides(itemCount - initialOffset, itemsPerSlide);
        
        // If collection is empty, create one empty slide
        if (itemCount == 0)
        {
            slidePlan.AddSlideInstance(new SlideInstance
            {
                SourceSlideId = slideInfo.SlideId,
                Position = GetNextPosition(slidePlan),
                IndexOffset = 0,
                Type = SlideInstanceType.Generated,
                ContextPath = resolvedPath.Split('>').ToList(),
                IsEmpty = true
            });
            return;
        }
        
        // Create slide instances
        for (int i = 0; i < requiredSlides; i++)
        {            int offset = initialOffset + (i * itemsPerSlide);
              slidePlan.AddSlideInstance(new SlideInstance
            {
                SourceSlideId = slideInfo.SlideId,
                Position = GetNextPosition(slidePlan),
                IndexOffset = offset,
                Type = SlideInstanceType.Generated,
                ContextPath = resolvedPath.Split('>').ToList()
            });
        }
    }

    /// <summary>
    /// Resolves an alias path to its full path
    /// </summary>
    private string ResolveAliasPath(string path, Dictionary<string, string> aliasMap)
    {
        // Check if this is an alias
        if (aliasMap.TryGetValue(path, out string? resolvedPath))
            return resolvedPath;

        // Handle more complex cases like "Alias[0]"
        foreach (var alias in aliasMap.Keys)
        {
            if (path.StartsWith(alias + "[") || path.StartsWith(alias + "."))
            {
                return path.Replace(alias, aliasMap[alias]);
            }
        }
        
        return path;
    }    /// <summary>
    /// Resolves a collection from a data object using a path with comprehensive context handling
    /// </summary>
    private IEnumerable<object>? ResolveCollection(object data, string path)
    {
        if (data == null || string.IsNullOrEmpty(path))
            return Enumerable.Empty<object>(); // Return empty collection instead of null for design-centered approach
            
        try
        {            // Handle context operator
            if (path.Contains(">"))
            {
                // Try enhanced nested collection resolution first
                var enhancedResult = ResolveNestedCollectionWithContext(data, path);
                if (enhancedResult != null && enhancedResult.Any())
                {
                    return enhancedResult;
                }

                // Fallback to original logic
                string[] segments = path.Split('>');
                object? currentContext = data;
                
                for (int i = 0; i < segments.Length; i++)
                {
                    if (currentContext == null)
                        throw new InvalidOperationException("Encountered null context while resolving path");
                        
                    string segment = segments[i].Trim();
                    
                    // Handle array indexer in path segment
                    var arrayMatch = Regex.Match(segment, @"(.+?)\[(\d+)\]");
                    if (arrayMatch.Success)
                    {
                        string arrayName = arrayMatch.Groups[1].Value;
                        int index = int.Parse(arrayMatch.Groups[2].Value);
                          // Get the array property
                        var property = currentContext.GetType().GetProperty(arrayName);
                        object? arrayObject = null;
                        
                        if (property == null)
                        {
                            // Try dictionary access if it's a Dictionary<string, object>
                            if (currentContext is IDictionary<string, object> dict)
                            {
                                if (!dict.TryGetValue(arrayName, out arrayObject))
                                    throw new InvalidOperationException($"Property '{arrayName}' not found in context object");
                            }
                            else
                            {
                                throw new InvalidOperationException($"Property '{arrayName}' not found in context object");
                            }
                        }
                        else
                        {
                            arrayObject = property.GetValue(currentContext);
                        }
                            
                        var array = arrayObject as IList;
                        if (array == null)
                            throw new InvalidOperationException($"Property '{arrayName}' is not a collection");
                            
                        if (index >= array.Count)
                            throw new InvalidOperationException($"Index {index} is out of range for collection '{arrayName}' (Count: {array.Count})");
                            
                        currentContext = array[index];
                    }                    else
                    {
                        // Regular property access - handle dot notation
                        string[] dotSegments = segment.Split('.');
                        object? tempContext = currentContext;
                          foreach (string dotSegment in dotSegments)
                        {
                            if (tempContext == null)
                                throw new InvalidOperationException($"Encountered null context while resolving dot notation path");
                                
                            var property = tempContext.GetType().GetProperty(dotSegment);
                            if (property == null)
                            {
                                // Try dictionary access if it's a Dictionary<string, object>
                                if (tempContext is IDictionary<string, object> dict)
                                {
                                    if (!dict.TryGetValue(dotSegment, out object? value))
                                        throw new InvalidOperationException($"Property '{dotSegment}' not found in context object");
                                    tempContext = value;
                                }
                                else
                                {
                                    throw new InvalidOperationException($"Property '{dotSegment}' not found in context object");
                                }
                            }
                            else
                            {
                                tempContext = property.GetValue(tempContext);
                            }
                        }
                        
                        currentContext = tempContext;
                    }
                    
                    // If we've reached the last segment, check if it's a collection
                    if (i == segments.Length - 1)
                    {                        if (currentContext == null)
                            return Enumerable.Empty<object>();
                            
                        if (currentContext is IEnumerable enumerableObj && !(currentContext is string))
                            return enumerableObj.Cast<object>();
                            
                        throw new InvalidOperationException($"Path '{path}' does not resolve to a collection");
                    }
                }
            }
            
            // Handle direct property access
            var parts = path.Split('.');
            object? currentObject = data;
            
            foreach (var part in parts)
            {
                if (currentObject == null)
                    throw new InvalidOperationException("Encountered null object while resolving path");
                    
                // Handle array indexer
                var arrayMatch = Regex.Match(part, @"(.+?)\[(\d+)\]");
                if (arrayMatch.Success)
                {
                    string arrayName = arrayMatch.Groups[1].Value;                    int index = int.Parse(arrayMatch.Groups[2].Value);
                    
                    var property = currentObject.GetType().GetProperty(arrayName);
                    object? arrayObject = null;
                    
                    if (property == null)
                    {
                        // Try dictionary access if it's a Dictionary<string, object>
                        if (currentObject is IDictionary<string, object> dict)
                        {
                            if (!dict.TryGetValue(arrayName, out arrayObject))
                                throw new InvalidOperationException($"Property '{arrayName}' not found in data object");
                        }
                        else
                        {
                            throw new InvalidOperationException($"Property '{arrayName}' not found in data object");
                        }
                    }
                    else
                    {
                        arrayObject = property.GetValue(currentObject);
                    }
                        
                    var array = arrayObject as IList;
                    if (array == null)
                        throw new InvalidOperationException($"Property '{arrayName}' is not a collection");
                        
                    if (index >= array.Count)
                        throw new InvalidOperationException($"Index {index} is out of range for collection '{arrayName}' (Count: {array.Count})");
                        
                    currentObject = array[index];
                }                else
                {
                    var property = currentObject.GetType().GetProperty(part);
                    if (property == null)
                    {
                        // Try dictionary access if it's a Dictionary<string, object>
                        if (currentObject is IDictionary<string, object> dict)
                        {
                            if (!dict.TryGetValue(part, out object? value))
                                throw new InvalidOperationException($"Property '{part}' not found in data object");
                            currentObject = value;
                        }
                        else
                        {
                            throw new InvalidOperationException($"Property '{part}' not found in data object");
                        }
                    }
                    else
                    {
                        currentObject = property.GetValue(currentObject);
                        if (currentObject == null)
                            throw new InvalidOperationException($"Property '{part}' returned null");
                    }
                }
            }
              // Check if the resolved object is a collection
            if (currentObject == null)
                return Enumerable.Empty<object>();
                
            if (currentObject is IEnumerable enumerableResult && !(currentObject is string))
                return enumerableResult.Cast<object>();
                throw new InvalidOperationException($"Path '{path}' does not resolve to a collection");
        }
        catch (InvalidOperationException)
        {
            // Rethrow InvalidOperationException as it contains specific info about the collection path
            throw;
        }
        catch (Exception ex)
        {
            // Wrap other exceptions for a consistent error handling
            throw new InvalidOperationException($"Error resolving collection path '{path}': {ex.Message}", ex);
        }
    }
    
    /// <summary>
    /// Resolves a property path to get an object with enhanced array access support
    /// </summary>
    private object? ResolvePropertyPath(object? data, string path)
    {
        if (data == null || string.IsNullOrEmpty(path))
            return null;

        try
        {
            // Handle property path traversal
            string[] parts = path.Split('.');
            object? current = data;
            
            foreach (var part in parts)
            {
                if (current == null)
                    return null;
                    
                // Handle array access: Items[0]
                var arrayMatch = System.Text.RegularExpressions.Regex.Match(part, @"(.+?)\[(\d+)\]");
                if (arrayMatch.Success)
                {
                    string arrayName = arrayMatch.Groups[1].Value;
                    int index = int.Parse(arrayMatch.Groups[2].Value);
                    
                    // Get the array property
                    var property = current.GetType().GetProperty(arrayName);
                    if (property == null)
                    {
                        // 디자인 중심 접근법: 속성을 찾을 수 없으면 필드를 시도
                        var field = current.GetType().GetField(arrayName);
                        if (field != null)
                        {
                            var array = field.GetValue(current);
                            if (array is IList list)
                            {
                                if (index >= 0 && index < list.Count)
                                    current = list[index];
                                else
                                    return null; // 인덱스가 범위를 벗어남
                                continue;
                            }
                        }
                        // 대소문자 구분 없이 속성 검색 시도
                        property = current.GetType().GetProperties()
                            .FirstOrDefault(p => string.Equals(p.Name, arrayName, StringComparison.OrdinalIgnoreCase));
                        
                        if (property == null)
                            return null;
                    }
                      var arrayProperty = property.GetValue(current);
                    if (arrayProperty == null)
                        return null;
                        
                    // Access the array element
                    if (arrayProperty is IList listProperty)
                    {
                        if (index >= 0 && index < listProperty.Count)
                            current = listProperty[index];
                        else
                            return null; // 인덱스가 범위를 벗어남
                    }
                    else
                    {
                        return null;
                    }
                }                else
                {
                    // Regular property access
                    var property = current.GetType().GetProperty(part);
                    if (property == null)
                    {
                        // 디자인 중심 접근법: 대소문자 구분 없이 속성 검색 시도
                        property = current.GetType().GetProperties()
                            .FirstOrDefault(p => string.Equals(p.Name, part, StringComparison.OrdinalIgnoreCase));
                        
                        // 필드 검색 시도
                        if (property == null)
                        {
                            var field = current.GetType().GetField(part) ?? 
                                        current.GetType().GetFields()
                                        .FirstOrDefault(f => string.Equals(f.Name, part, StringComparison.OrdinalIgnoreCase));
                            
                            if (field != null)
                                current = field.GetValue(current);
                            else
                                return null; // 속성이나 필드를 찾을 수 없음
                                
                            continue;
                        }
                    }
                    {
                        // 디자인 중심 접근법: 속성을 찾을 수 없으면 필드를 시도
                        var field = current.GetType().GetField(part);
                        if (field != null)
                        {
                            current = field.GetValue(current);
                            continue;
                        }
                        
                        // 대소문자 구분 없이 속성 검색 시도
                        property = current.GetType().GetProperties()
                            .FirstOrDefault(p => string.Equals(p.Name, part, StringComparison.OrdinalIgnoreCase));
                            
                        if (property == null)
                            return null;
                    }
                    
                    current = property.GetValue(current);
                }
            }
            
            return current;
        }
        catch (Exception)
        {
            // 디자인 중심 접근법: 예외가 발생해도 null을 반환하여 호출자가 처리할 수 있도록 함
            return null;
        }
    }    /// <summary>
    /// Gets the next available position in the slide plan
    /// </summary>
    private int GetNextPosition(SlidePlan slidePlan)
    {
        return slidePlan.SlideInstances.Count > 0 
            ? slidePlan.SlideInstances.Max(i => i.Position) + 1 
            : 0;
    }

    /// <summary>
    /// Converts SlideType to SlideInstanceType
    /// </summary>
    private SlideInstanceType ConvertSlideTypeToInstanceType(SlideType slideType)
    {
        return slideType switch
        {
            SlideType.Static => SlideInstanceType.Static,
            SlideType.Source => SlideInstanceType.Generated,
            SlideType.Cloned => SlideInstanceType.Generated,
            _ => SlideInstanceType.Static
        };
    }    /// <summary>
    /// Enhanced method to resolve nested collection paths using DollarSignEngine
    /// </summary>
    private IEnumerable<object>? ResolveNestedCollectionWithContext(object data, string path, object? contextItem = null)
    {
        if (data == null || string.IsNullOrEmpty(path))
            return Enumerable.Empty<object>();

        try
        {
            // Prepare variables for DollarSignEngine
            var variables = new Dictionary<string, object?>();
            
            // Add root data properties
            if (data is Dictionary<string, object> dict)
            {
                foreach (var kvp in dict)
                {
                    variables[kvp.Key] = kvp.Value;
                }
            }
            else
            {
                var properties = data.GetType().GetProperties();
                foreach (var prop in properties)
                {
                    if (prop.CanRead)
                    {
                        try
                        {
                            variables[prop.Name] = prop.GetValue(data);
                        }
                        catch { /* Skip problematic properties */ }
                    }
                }
            }

            // Add context item if available
            if (contextItem != null)
            {
                variables["__contextItem"] = contextItem;
                
                // Add context item properties
                var contextProperties = contextItem.GetType().GetProperties();
                foreach (var prop in contextProperties)
                {
                    if (prop.CanRead)
                    {
                        try
                        {
                            variables[$"__context_{prop.Name}"] = prop.GetValue(contextItem);
                        }
                        catch { /* Skip problematic properties */ }
                    }
                }
            }

            // Configure DollarSignEngine options
            var options = new DollarSignOptions
            {
                SupportDollarSignSyntax = true,
                ThrowOnError = false,
                UseCache = true,
                CultureInfo = CultureInfo.CurrentCulture
            };

            // Handle nested paths like "Categories>Products" 
            if (path.Contains(">"))
            {
                // If we have a context item, try to resolve from context first
                if (contextItem != null)
                {
                    var segments = path.Split('>');
                    string lastSegment = segments[segments.Length - 1].Trim();
                    
                    // Use DollarSignEngine to evaluate the nested property
                    string expression = $"${{__context_{lastSegment}}}";
                    var result = DollarSign.Eval(expression, variables, options);
                    
                    if (result is IEnumerable enumerable && !(result is string))
                    {
                        return enumerable.Cast<object>();
                    }
                }
                
                // Fallback: resolve the full path using DollarSignEngine
                var pathSegments = path.Split('>');
                string dollarExpression = string.Join(".", pathSegments);
                dollarExpression = $"${{{dollarExpression}}}";
                
                var pathResult = DollarSign.Eval(dollarExpression, variables, options);
                if (pathResult is IEnumerable pathEnumerable && !(pathResult is string))
                {
                    return pathEnumerable.Cast<object>();
                }
            }
            
            return Enumerable.Empty<object>();
        }
        catch (Exception ex)
        {
            Logger.Warning($"Failed to resolve nested collection path '{path}' using DollarSignEngine: {ex.Message}");
            return Enumerable.Empty<object>();
        }
    }
}
