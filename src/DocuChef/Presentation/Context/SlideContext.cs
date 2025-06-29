using DocumentFormat.OpenXml.Packaging;
using DocuChef.Presentation.Models;
using DocuChef.Logging;

namespace DocuChef.Presentation.Context;

/// <summary>
/// 개별 슬라이드 처리 컨텍스트
/// 특정 슬라이드의 데이터 바인딩과 처리 정보를 관리
/// </summary>
public class SlideContext
{
    /// <summary>
    /// 상위 PPT 컨텍스트 참조
    /// </summary>
    public PPTContext PPTContext { get; }

    /// <summary>
    /// 슬라이드 인덱스 (0부터 시작)
    /// </summary>
    public int SlideIndex { get; }

    /// <summary>
    /// 슬라이드 인스턴스 정보 (생성 계획에서 가져온)
    /// </summary>
    public SlideInstance SlideInstance { get; }

    /// <summary>
    /// 원본 템플릿 슬라이드 정보
    /// </summary>
    public SlideInfo TemplateSlideInfo { get; }

    /// <summary>
    /// 실제 슬라이드 파트
    /// </summary>
    public SlidePart? SlidePart { get; set; }

    /// <summary>
    /// 이 슬라이드에서 사용할 바인딩 데이터
    /// (컨텍스트 경로, 인덱스 등이 적용된 최종 데이터)
    /// </summary>
    public Dictionary<string, object> BindingData { get; set; } = new();

    /// <summary>
    /// 이 슬라이드의 컨텍스트 경로 (예: "Items[0]", "Users[1].Orders[0]")
    /// </summary>
    public string ContextPath { get; }

    /// <summary>
    /// 현재 반복문에서의 인덱스 정보
    /// </summary>
    public Dictionary<string, int> IterationIndices { get; set; } = new();

    /// <summary>
    /// 현재 반복문에서의 오프셋 정보
    /// </summary>
    public Dictionary<string, int> IterationOffsets { get; set; } = new();

    /// <summary>
    /// 이 슬라이드에서 처리된 표현식들
    /// </summary>
    public List<string> ProcessedExpressions { get; set; } = new();

    /// <summary>
    /// PERFORMANCE OPTIMIZATION: Cached variables for this slide to avoid recreating them for each expression
    /// </summary>
    private Dictionary<string, object>? _cachedVariables = null; public SlideContext(PPTContext pptContext, int slideIndex, SlideInstance slideInstance)
    {
        PPTContext = pptContext ?? throw new ArgumentNullException(nameof(pptContext));
        SlideIndex = slideIndex;
        SlideInstance = slideInstance ?? throw new ArgumentNullException(nameof(slideInstance));

        // Debug logging to understand ContextPath issue
        Logger.Debug($"SlideContext: Creating for slide {slideIndex}, SourceSlideId: {slideInstance.SourceSlideId}");
        Logger.Debug($"SlideContext: SlideInstance.ContextPath count: {slideInstance.ContextPath?.Count ?? -1}");
        if (slideInstance.ContextPath != null)
        {
            Logger.Debug($"SlideContext: SlideInstance.ContextPath items: [{string.Join(", ", slideInstance.ContextPath)}]");
        }
        Logger.Debug($"SlideContext: SlideInstance.ContextPathString: '{slideInstance.ContextPathString}'");

        ContextPath = slideInstance.ContextPathString;

        // 템플릿 슬라이드 정보 찾기
        TemplateSlideInfo = pptContext.TemplateSlides.FirstOrDefault(s => s.Position == slideInstance.SourceSlideId)
            ?? throw new InvalidOperationException($"Template slide {slideInstance.SourceSlideId} not found");

        InitializeBindingData();
    }

    /// <summary>
    /// 바인딩 데이터 초기화
    /// 컨텍스트 경로와 인덱스를 고려하여 이 슬라이드에서 사용할 데이터 준비
    /// </summary>
    private void InitializeBindingData()
    {
        // 기본 변수들 복사
        foreach (var kvp in PPTContext.Variables)
        {
            BindingData[kvp.Key] = kvp.Value;
        }

        // 컨텍스트 경로가 있으면 해당 데이터로 컨텍스트 설정
        if (!string.IsNullOrEmpty(ContextPath))
        {
            var contextData = ResolveContextData(PPTContext.Variables, ContextPath);
            if (contextData != null)
            {
                // 컨텍스트 데이터의 속성들을 직접 바인딩 데이터에 추가
                if (contextData is IDictionary<string, object> dict)
                {
                    foreach (var kvp in dict)
                    {
                        BindingData[kvp.Key] = kvp.Value;
                    }
                }
                else
                {
                    // 객체의 속성들을 바인딩 데이터에 추가
                    var properties = contextData.GetType().GetProperties();
                    foreach (var prop in properties)
                    {
                        if (prop.CanRead)
                        {
                            try
                            {
                                var value = prop.GetValue(contextData);
                                if (value != null)
                                {
                                    BindingData[prop.Name] = value;
                                }
                            }
                            catch (Exception)
                            {
                                // 속성 읽기 실패 시 무시
                            }
                        }
                    }
                }
            }
        }
        // 반복 인덱스와 오프셋 정보 추가
        if (!string.IsNullOrEmpty(SlideInstance.CollectionName))
        {
            IterationIndices[SlideInstance.CollectionName] = SlideInstance.StartIndex;
            IterationOffsets[SlideInstance.CollectionName] = SlideInstance.IndexOffset;

            // 바인딩 데이터에도 인덱스 정보 추가
            BindingData[$"{SlideInstance.CollectionName}Index"] = SlideInstance.StartIndex;
            BindingData[$"{SlideInstance.CollectionName}Offset"] = SlideInstance.IndexOffset;
        }
    }

    /// <summary>
    /// 컨텍스트 경로로부터 실제 데이터 해석
    /// </summary>
    private object? ResolveContextData(Dictionary<string, object> variables, string contextPath)
    {
        try
        {
            var segments = contextPath.Split('.');
            object? current = null;

            foreach (var segment in segments)
            {
                if (current == null)
                {
                    // 첫 번째 세그먼트는 변수에서 찾기
                    if (segment.Contains('['))
                    {
                        // 배열 인덱스 처리 (예: "Items[0]")
                        var arrayName = segment.Substring(0, segment.IndexOf('['));
                        var indexStr = segment.Substring(segment.IndexOf('[') + 1, segment.IndexOf(']') - segment.IndexOf('[') - 1);

                        if (variables.TryGetValue(arrayName, out var arrayObj) && int.TryParse(indexStr, out var index))
                        {
                            if (arrayObj is System.Collections.IList list && index < list.Count)
                            {
                                current = list[index];
                            }
                        }
                    }
                    else
                    {
                        variables.TryGetValue(segment, out current);
                    }
                }
                else
                {
                    // 중첩 속성 처리
                    if (segment.Contains('['))
                    {
                        // 배열 인덱스 처리
                        var arrayName = segment.Substring(0, segment.IndexOf('['));
                        var indexStr = segment.Substring(segment.IndexOf('[') + 1, segment.IndexOf(']') - segment.IndexOf('[') - 1);

                        var prop = current.GetType().GetProperty(arrayName);
                        if (prop != null && int.TryParse(indexStr, out var index))
                        {
                            var arrayObj = prop.GetValue(current);
                            if (arrayObj is System.Collections.IList list && index < list.Count)
                            {
                                current = list[index];
                            }
                        }
                    }
                    else
                    {
                        var prop = current.GetType().GetProperty(segment);
                        current = prop?.GetValue(current);
                    }
                }
            }

            return current;
        }
        catch (Exception)
        {
            return null;
        }
    }

    /// <summary>
    /// 표현식을 이 슬라이드의 컨텍스트에 맞게 변환
    /// </summary>
    public string TransformExpression(string expression)
    {
        // 인덱스 및 오프셋 정보를 사용하여 표현식 변환
        var transformed = expression;

        foreach (var kvp in IterationIndices)
        {
            // ${Index} -> ${0}, ${Offset} -> ${1} 등으로 변환
            transformed = transformed.Replace($"${{{kvp.Key}Index}}", $"{kvp.Value}");
            transformed = transformed.Replace($"${{{kvp.Key}Offset}}", $"{IterationOffsets[kvp.Key]}");
        }

        return transformed;
    }

    /// <summary>
    /// 표현식 처리 완료 기록
    /// </summary>
    public void MarkExpressionProcessed(string expression)
    {
        if (!ProcessedExpressions.Contains(expression))
        {
            ProcessedExpressions.Add(expression);
        }
    }

    /// <summary>
    /// 이 슬라이드의 바인딩 데이터에서 특정 값 가져오기
    /// </summary>
    public object? GetBindingValue(string key)
    {
        return BindingData.TryGetValue(key, out var value) ? value : null;
    }    /// <summary>
         /// 이 슬라이드의 바인딩 데이터에 값 설정
         /// </summary>
    public void SetBindingValue(string key, object value)
    {
        BindingData[key] = value;
    }    /// <summary>
         /// PERFORMANCE OPTIMIZATION: Get cached variables for this slide context
         /// Variables are prepared once per slide and reused for all expressions within the slide
         /// </summary>
    public Dictionary<string, object> GetCachedVariables()
    {
        if (_cachedVariables == null)
        {
            Logger.Debug($"SlideContext: Preparing variables for slide {SlideIndex} with context '{ContextPath ?? "null"}'");

            // For nested contexts, we need to use the actual data object, not the variables dictionary
            // Start with base variables and apply context-specific transformations
            _cachedVariables = new Dictionary<string, object>(PPTContext.Variables);

            // If we have a context path, resolve it and apply transformations
            if (!string.IsNullOrEmpty(ContextPath))
            {
                // Use the DataBinder's ApplyContextPath method directly with the original data
                // This ensures proper variable creation for nested contexts
                PPTContext.DataBinder.ApplyContextPath(_cachedVariables, PPTContext.Variables, ContextPath);
            }

            Logger.Debug($"SlideContext: Cached {_cachedVariables.Count} variables for slide {SlideIndex}");

            // Log key variables for debugging
            foreach (var kvp in _cachedVariables)
            {
                var valueInfo = kvp.Value?.GetType().Name ?? "null";
                if (kvp.Value is System.Collections.IEnumerable enumerable && !(kvp.Value is string))
                {
                    var count = enumerable.Cast<object>().Count();
                    valueInfo += $" (count: {count})";
                }
                Logger.Debug($"  Variable: {kvp.Key} = {valueInfo}");
            }
        }

        return _cachedVariables;
    }

    /// <summary>
    /// Clear cached variables (used when data changes)
    /// </summary>
    public void ClearVariableCache()
    {
        _cachedVariables = null;
        Logger.Debug($"SlideContext: Cleared variable cache for slide {SlideIndex}");
    }
}
