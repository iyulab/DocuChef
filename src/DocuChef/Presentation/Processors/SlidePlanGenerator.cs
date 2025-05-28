using System.Reflection;
using DocuChef.Presentation.Models;
using DocuChef.Presentation.Utilities;
using System.Text.RegularExpressions;
using DocuChef.Logging;

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
                    slidePlan.AddSlideInstance(new SlideInstance
                    {
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
            for (int i = 0; i < requiredSlides; i++)
            {
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
            Logger.Debug($"ProcessStandaloneSlides: Processing slide {slideInfo.SlideId}, Type={slideInfo.Type}, CollectionName='{slideInfo.CollectionName}', MaxArrayIndex={slideInfo.MaxArrayIndex}, HasArrayReferences={slideInfo.HasArrayReferences}");

            var foreachDirectives = slideInfo.Directives
                .Where(d => d.Type == DirectiveType.Foreach)
                .ToList();

            Logger.Debug($"ProcessStandaloneSlides: Found {foreachDirectives.Count} foreach directives for slide {slideInfo.SlideId}");
            Logger.Debug($"ProcessStandaloneSlides: Total directives: {slideInfo.Directives.Count} - Types: {string.Join(", ", slideInfo.Directives.Select(d => d.Type))}");

            if (foreachDirectives.Any())
            {
                // This slide has foreach directives, we need to create clones
                Logger.Debug($"ProcessStandaloneSlides: Processing explicit foreach directives for slide {slideInfo.SlideId}");
                foreach (var directive in foreachDirectives)
                {
                    ProcessForeachDirective(slideInfo, slidePlan, data, directive, aliasMap);
                }
            }
            else if (slideInfo.Type == SlideType.Source && !string.IsNullOrEmpty(slideInfo.CollectionName))
            {
                // This is a source slide with collection binding - infer foreach directive
                Logger.Debug($"ProcessStandaloneSlides: Inferring foreach directive for slide {slideInfo.SlideId}, Collection='{slideInfo.CollectionName}', MaxItems={slideInfo.MaxArrayIndex + 1}");
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
                Logger.Debug($"ProcessStandaloneSlides: Treating slide {slideInfo.SlideId} as static - Type={slideInfo.Type}, CollectionName='{slideInfo.CollectionName}', HasArrayReferences={slideInfo.HasArrayReferences}");
                slidePlan.AddSlideInstance(new SlideInstance
                {
                    SourceSlideId = slideInfo.SlideId,
                    Position = GetNextPosition(slidePlan),
                    IndexOffset = 0,
                    Type = SlideInstanceType.Static // Always treat as static when no directives or collection binding
                });
            }
        }
    }    /// <summary>
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

        Logger.Debug($"ProcessForeachDirective: itemCount={itemCount}, itemsPerSlide={itemsPerSlide}, initialOffset={initialOffset}, requiredSlides={requiredSlides}");

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
        {
            int offset = initialOffset + (i * itemsPerSlide);
            Logger.Debug($"ProcessForeachDirective: Creating slide {i + 1}/{requiredSlides} with offset={offset}");
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
        {
            Logger.Debug($"ResolveCollection: Early return - data={data}, path='{path}'");
            return Enumerable.Empty<object>(); // Return empty collection instead of null for design-centered approach
        }

        Logger.Debug($"ResolveCollection: Starting with path='{path}', data type={data.GetType().Name}");

        try
        {            // Handle context operator
            if (path.Contains(">"))
            {
                Logger.Debug($"ResolveCollection: Handling context operator in path '{path}'");
                string[] segments = path.Split('>');
                object? currentContext = data;

                for (int i = 0; i < segments.Length; i++)
                {
                    if (currentContext == null)
                    {
                        Logger.Debug($"ResolveCollection: CurrentContext is null at segment {i}");
                        return Enumerable.Empty<object>();
                    }

                    string segment = segments[i].Trim();
                    Logger.Debug($"ResolveCollection: Processing segment {i}: '{segment}'");
                    currentContext = ResolvePropertyPath(currentContext, segment);
                    Logger.Debug($"ResolveCollection: Segment {i} resolved to: {currentContext?.GetType().Name ?? "null"}");

                    // If we've reached the last segment, check if it's a collection
                    if (i == segments.Length - 1)
                    {
                        if (currentContext == null)
                        {
                            Logger.Debug($"ResolveCollection: Final segment result is null");
                            return Enumerable.Empty<object>();
                        }

                        if (currentContext is IEnumerable enumerableObj && !(currentContext is string))
                        {
                            var result = enumerableObj.Cast<object>();
                            Logger.Debug($"ResolveCollection: Found enumerable, count={result.Count()}");
                            return result;
                        }

                        Logger.Debug($"ResolveCollection: Final segment is not enumerable: {currentContext.GetType().Name}");
                        return Enumerable.Empty<object>();
                    }
                }
            }            // Handle direct property access
            Logger.Debug($"ResolveCollection: Handling direct property access for path '{path}'");
            var resolvedObject = ResolvePropertyPath(data, path);
            Logger.Debug($"ResolveCollection: ResolvePropertyPath returned: {resolvedObject?.GetType().Name ?? "null"}");

            if (resolvedObject == null)
            {
                Logger.Debug($"ResolveCollection: ResolvePropertyPath returned null");
                return Enumerable.Empty<object>();
            }

            if (resolvedObject is IEnumerable enumerableResult && !(resolvedObject is string))
            {
                var result = enumerableResult.Cast<object>();
                Logger.Debug($"ResolveCollection: Found enumerable result, count={result.Count()}");
                return result;
            }

            Logger.Debug($"ResolveCollection: Resolved object is not enumerable: {resolvedObject.GetType().Name}");
            return Enumerable.Empty<object>();
        }
        catch (InvalidOperationException)
        {
            // Rethrow InvalidOperationException as it contains specific info about the collection path
            throw;
        }
        catch (Exception ex)
        {
            Logger.Debug($"ResolveCollection: Exception caught: {ex.Message}");
            // Wrap other exceptions for a consistent error handling
            throw new InvalidOperationException($"Error resolving collection path '{path}': {ex.Message}", ex);
        }
    }    /// <summary>
         /// Resolves a property path to get an object using simple reflection (no data binding)
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
                var arrayMatch = Regex.Match(part, @"(.+?)\[(\d+)\]");
                if (arrayMatch.Success)
                {
                    string arrayName = arrayMatch.Groups[1].Value;
                    int index = int.Parse(arrayMatch.Groups[2].Value);

                    // Get the array property
                    var property = current.GetType().GetProperty(arrayName);
                    if (property == null)
                    {
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
                            return null;
                    }
                    else
                    {
                        return null;
                    }
                }
                else
                {
                    // Handle Dictionary<string, object> access first
                    if (current is IDictionary<string, object> dict)
                    {
                        if (!dict.TryGetValue(part, out var value))
                            return null;
                        current = value;
                    }
                    // Handle regular property access via reflection
                    else
                    {
                        var property = current.GetType().GetProperty(part);
                        if (property == null)
                            return null;
                        current = property.GetValue(current);
                    }
                }
            }

            return current;
        }
        catch (Exception)
        {
            return null;
        }
    }/// <summary>
     /// Gets the next available position in the slide plan
     /// </summary>
    private int GetNextPosition(SlidePlan slidePlan)
    {
        return slidePlan.SlideInstances.Count > 0
            ? slidePlan.SlideInstances.Max(i => i.Position) + 1
            : 0;
    }    /// <summary>
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
    }
}
