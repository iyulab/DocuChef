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
{    /// <summary>
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
        var aliasMap = BuildAliasMap(slideInfos);        // Check if we have nested context (Products>Items pattern)
        // Look for binding expressions that contain nested paths (with ">")
        var hasNestedContext = slideInfos.Any(si =>
            si.BindingExpressions.Any(expr => expr.DataPath.Contains(">")));

        if (hasNestedContext)
        {
            Logger.Debug("GeneratePlan: Detected nested context, using nested range processing");
            // Process static slides first (before ranges)
            var staticSlides = slideInfos.Where(si => si.Type == SlideType.Static && si.SlideId < slideInfos.Where(s => s.Type == SlideType.Source).Min(s => s.SlideId)).ToList(); foreach (var staticSlide in staticSlides)
            {
                slidePlan.AddSlideInstance(new SlideInstance
                {
                    SourceSlideId = staticSlide.SlideId,
                    Position = GetNextPosition(slidePlan),
                    ContextPath = new List<string>(),
                    IndexOffset = 0
                });
            }

            // Process nested ranges in correct order
            ProcessNestedRangeSlides(slideInfos, slidePlan, data, aliasMap);            // Process static slides after ranges (END slide)
            var endSlides = slideInfos.Where(si => si.Type == SlideType.Static && si.SlideId > slideInfos.Where(s => s.Type == SlideType.Source).Max(s => s.SlideId)).ToList();
            foreach (var endSlide in endSlides)
            {
                slidePlan.AddSlideInstance(new SlideInstance
                {
                    SourceSlideId = endSlide.SlideId,
                    Position = GetNextPosition(slidePlan),
                    ContextPath = new List<string>(),
                    IndexOffset = 0
                });
            }
        }
        else
        {
            // Original processing for non-nested scenarios
            // Process slides with range directives
            ProcessRangeDirectives(slideInfos, slidePlan, data, aliasMap);

            // Process slides with foreach directives that are not part of a range
            ProcessStandaloneSlides(slideInfos, slidePlan, data, aliasMap);
        }

        return slidePlan;
    }

    /// <summary>
    /// Calculates the number of slides required to display all items
    /// </summary>
    public int CalculateRequiredSlides(int itemCount, int itemsPerSlide)
    {
        if (itemCount == 0 || itemsPerSlide <= 0) return 0;
        return (int)Math.Ceiling((double)itemCount / itemsPerSlide);
    }    /// <summary>
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
                    aliasMap[directive.AliasName] = directive.CollectionPath;
                    Logger.Debug($"BuildAliasMap: Added alias mapping '{directive.AliasName}' -> '{directive.CollectionPath}'");
                }
            }
        }

        Logger.Debug($"BuildAliasMap: Created alias map with {aliasMap.Count} entries");
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
            Logger.Debug($"ProcessStandaloneSlides: Total directives: {slideInfo.Directives.Count} - Types: {string.Join(", ", slideInfo.Directives.Select(d => d.Type))}");            // Check if this slide has nested context expressions (Products>Items) first
            bool hasNestedContext = slideInfo.BindingExpressions.Any(e => e.DataPath.Contains(">"));

            if (hasNestedContext)
            {
                Logger.Debug($"ProcessStandaloneSlides: Processing nested context for slide {slideInfo.SlideId}");
                ProcessNestedContextSlide(slideInfo, slidePlan, data, aliasMap);
            }
            else if (foreachDirectives.Any())
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
                if (hasNestedContext)
                {
                    Logger.Debug($"ProcessStandaloneSlides: Processing nested context for slide {slideInfo.SlideId}");
                    ProcessNestedContextSlide(slideInfo, slidePlan, data, aliasMap);
                }
                else
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
    }    /// <summary>
         /// Processes a slide with nested context expressions (e.g., Products>Items)
         /// This handles cases where we need to iterate over parent collections and their nested collections
         /// </summary>
    private void ProcessNestedContextSlide(SlideInfo slideInfo, SlidePlan slidePlan, object data, Dictionary<string, string> aliasMap)
    {
        Logger.Debug($"ProcessNestedContextSlide: Processing slide {slideInfo.SlideId} with nested context");

        // Find nested context expressions in the slide
        var nestedExpressions = slideInfo.BindingExpressions
            .Where(e => e.DataPath.Contains(">"))
            .ToList();

        if (!nestedExpressions.Any())
        {
            Logger.Debug($"ProcessNestedContextSlide: No nested expressions found in slide {slideInfo.SlideId}");
            return;
        }

        // Extract the nested context pattern (e.g., "Products>Items")
        var firstNestedExpression = nestedExpressions.First();
        var contextParts = firstNestedExpression.DataPath.Split('>');

        if (contextParts.Length < 2)
        {
            Logger.Debug($"ProcessNestedContextSlide: Invalid nested context pattern: {firstNestedExpression.DataPath}");
            return;
        }

        string parentPath = contextParts[0]; // "Products"
        string childPath = contextParts[1]; // "Items[0]" or "Items"

        // Extract the child collection name and array index pattern
        var childMatch = Regex.Match(childPath, @"(\w+)(?:\[(\d+)\])?");
        if (!childMatch.Success)
        {
            Logger.Debug($"ProcessNestedContextSlide: Invalid child path pattern: {childPath}");
            return;
        }

        string childCollectionName = childMatch.Groups[1].Value; // "Items"

        // Determine items per slide from the max array index in expressions
        int itemsPerSlide = slideInfo.MaxArrayIndex + 1;
        Logger.Debug($"ProcessNestedContextSlide: Detected {itemsPerSlide} items per slide from max array index {slideInfo.MaxArrayIndex}");

        // Resolve the parent collection (e.g., Products)
        var parentCollection = ResolveCollection(data, parentPath);
        if (parentCollection == null || !parentCollection.Any())
        {
            Logger.Debug($"ProcessNestedContextSlide: Parent collection '{parentPath}' is empty or null");
            return;
        }

        Logger.Debug($"ProcessNestedContextSlide: Found {parentCollection.Count()} items in parent collection '{parentPath}'");        // Check if there's a parent slide that needs to be processed first
                                                                                                                                       // TODO: Implement range-based nested processing structure
                                                                                                                                       // This should be handled by ProcessRangeDirectives method instead

        // For each parent item, create parent slide first, then child slides
        int parentIndex = 0;
        foreach (var parentItem in parentCollection)
        {
            Logger.Debug($"ProcessNestedContextSlide: Processing parent item {parentIndex} of {parentCollection.Count()}");

            // TODO: Implement range-based parent slide creation
            // This functionality should be moved to ProcessRangeDirectives method

            // Resolve the child collection from the parent item
            var childCollection = ResolveCollection(parentItem, childCollectionName);
            if (childCollection == null || !childCollection.Any())
            {
                Logger.Debug($"ProcessNestedContextSlide: Child collection '{childCollectionName}' is empty in parent item {parentIndex}");
                parentIndex++;
                continue;
            }

            int childItemCount = childCollection.Count();
            int requiredSlides = CalculateRequiredSlides(childItemCount, itemsPerSlide);

            Logger.Debug($"ProcessNestedContextSlide: Parent {parentIndex} has {childItemCount} child items, requiring {requiredSlides} slides with {itemsPerSlide} items per slide");

            // Create slide instances for this parent's child collection
            for (int slideIndex = 0; slideIndex < requiredSlides; slideIndex++)
            {
                int offset = slideIndex * itemsPerSlide;
                Logger.Debug($"ProcessNestedContextSlide: Creating child slide {slideIndex + 1}/{requiredSlides} for parent {parentIndex} with offset {offset}");
                slidePlan.AddSlideInstance(new SlideInstance
                {
                    SourceSlideId = slideInfo.SlideId,
                    Position = GetNextPosition(slidePlan),
                    IndexOffset = offset,
                    Type = SlideInstanceType.Generated,
                    ContextPath = new List<string> { parentPath, childCollectionName },
                    ParentIndex = parentIndex
                });
            }

            parentIndex++;
        }

        Logger.Debug($"ProcessNestedContextSlide: Completed processing nested context for slide {slideInfo.SlideId}");
    }/// <summary>
     /// Processes nested range slides in correct hierarchical order
     /// </summary>
    private void ProcessNestedRangeSlides(List<SlideInfo> slideInfos, SlidePlan slidePlan, object data, Dictionary<string, string> aliasMap)
    {
        // Find parent and child slides for nested context
        var parentSlides = slideInfos.Where(si =>
            si.Type == SlideType.Source &&
            si.HasArrayReferences &&
            !si.BindingExpressions.Any(e => e.DataPath.Contains(">"))
        ).ToList();

        var childSlides = slideInfos.Where(si =>
            si.Type == SlideType.Source &&
            si.BindingExpressions.Any(e => e.DataPath.Contains(">"))
        ).ToList();

        if (!parentSlides.Any() || !childSlides.Any())
        {
            Logger.Debug("ProcessNestedRangeSlides: No nested parent-child relationship found");
            return;
        }

        var parentSlide = parentSlides.First();
        var childSlide = childSlides.First();

        Logger.Debug($"ProcessNestedRangeSlides: Found parent slide {parentSlide.SlideId} and child slide {childSlide.SlideId}");

        // Get parent collection path
        var parentCollectionName = parentSlide.CollectionName;
        if (string.IsNullOrEmpty(parentCollectionName))
        {
            Logger.Debug("ProcessNestedRangeSlides: Parent collection name is empty");
            return;
        }

        // Resolve parent collection
        var parentCollection = ResolveCollection(data, parentCollectionName);
        if (parentCollection == null || !parentCollection.Any())
        {
            Logger.Debug($"ProcessNestedRangeSlides: Parent collection '{parentCollectionName}' is empty or null");
            return;
        }
        Logger.Debug($"ProcessNestedRangeSlides: Processing {parentCollection.Count()} parent items");

        // Process each parent item with its child items
        int parentIndex = 0;

        foreach (var parentItem in parentCollection)
        {
            Logger.Debug($"ProcessNestedRangeSlides: Processing parent {parentIndex}");

            // Add parent slide instance
            slidePlan.AddSlideInstance(new SlideInstance
            {
                SourceSlideId = parentSlide.SlideId,
                Position = GetNextPosition(slidePlan),
                ContextPath = new List<string> { parentCollectionName },
                IndexOffset = parentIndex
            });

            // Process child items for this parent with local indexing
            var nestedExpression = childSlide.BindingExpressions.FirstOrDefault(e => e.DataPath.Contains(">"));
            if (nestedExpression != null)
            {
                ProcessNestedChildItems(childSlide, slidePlan, parentItem, parentIndex, $"{parentCollectionName}[{parentIndex}]");
            }

            parentIndex++;
        }

        Logger.Debug("ProcessNestedRangeSlides: Completed nested range processing");
    }    /// <summary>
         /// Processes child items for a specific parent
         /// </summary>
    private void ProcessNestedChildItems(SlideInfo childSlide, SlidePlan slidePlan, object parentItem, int parentIndex, string parentContextPath, int globalChildOffset = 0)
    {
        // Get child collection name from nested expression
        var nestedExpression = childSlide.BindingExpressions.FirstOrDefault(e => e.DataPath.Contains(">"));
        if (nestedExpression == null)
        {
            Logger.Debug("ProcessNestedChildItems: No nested expression found");
            return;
        }

        var parts = nestedExpression.DataPath.Split('>');
        if (parts.Length != 2)
        {
            Logger.Debug($"ProcessNestedChildItems: Invalid nested expression: {nestedExpression.DataPath}");
            return;
        }

        var childCollectionName = parts[1].Split('[')[0]; // Get "Items" from "Items[0]"

        // Resolve child collection from parent item
        var childCollection = ResolveCollection(parentItem, childCollectionName);
        if (childCollection == null || !childCollection.Any())
        {
            Logger.Debug($"ProcessNestedChildItems: Child collection '{childCollectionName}' is empty for parent {parentIndex}");
            return;
        }

        var childItems = childCollection.ToList();
        var itemsPerSlide = childSlide.MaxArrayIndex + 1; // 0-based to count
        var requiredSlides = CalculateRequiredSlides(childItems.Count, itemsPerSlide);

        Logger.Debug($"ProcessNestedChildItems: Parent {parentIndex} has {childItems.Count} child items, requiring {requiredSlides} slides with {itemsPerSlide} items per slide");

        // Create child slides - use local offset within this parent's collection, not global
        for (int slideIndex = 0; slideIndex < requiredSlides; slideIndex++)
        {
            // Use local offset within current parent's collection (0-based for each parent)
            int offset = slideIndex * itemsPerSlide;
            Logger.Debug($"ProcessNestedChildItems: Creating child slide {slideIndex + 1}/{requiredSlides} for parent {parentIndex} with local offset {offset}");

            slidePlan.AddSlideInstance(new SlideInstance
            {
                SourceSlideId = childSlide.SlideId,
                Position = GetNextPosition(slidePlan),
                ContextPath = new List<string> { $"{parts[0]}[{parentIndex}]", childCollectionName },
                IndexOffset = offset
            });
        }
    }
}
