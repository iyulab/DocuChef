using DocumentFormat.OpenXml.Packaging;
using DocuChef.Presentation.Models;
using DocuChef.Presentation.Functions;
using DocuChef.Presentation.Processors;

namespace DocuChef.Presentation.Context;

/// <summary>
/// PowerPoint 처리 전체 컨텍스트
/// 템플릿 분석부터 최종 생성까지의 모든 정보를 관리
/// </summary>
public class PPTContext
{
    /// <summary>
    /// 원본 템플릿 프레젠테이션 문서
    /// </summary>
    public PresentationDocument TemplateDocument { get; }
    /// <summary>
    /// 작업 중인 프레젠테이션 문서
    /// </summary>
    public PresentationDocument WorkingDocument { get; set; } = null!;    /// <summary>
                                                                          /// 템플릿 분석 결과 - 원본 슬라이드 정보들
                                                                          /// </summary>
    public List<SlideInfo> TemplateSlides { get; set; } = new();

    /// <summary>
    /// 템플릿 분석 결과 - 슬라이드 정보들 (SlideInfos와 동일하지만 명확성을 위해 추가)
    /// </summary>
    public List<SlideInfo> SlideInfos => TemplateSlides;

    /// <summary>
    /// 슬라이드 생성 계획 - 데이터 바인딩을 고려한 최종 슬라이드 계획
    /// </summary>
    public SlidePlan GenerationPlan { get; set; } = new();

    /// <summary>
    /// 바인딩할 변수 및 데이터
    /// </summary>
    public Dictionary<string, object> Variables { get; set; } = new();    /// <summary>
                                                                          /// PowerPoint 함수 인스턴스 (ppt.Image 등)
                                                                          /// </summary>
    public PPTFunctions Functions { get; set; }

    /// <summary>
    /// PERFORMANCE OPTIMIZATION: Shared DataBinder instance for variable preparation
    /// </summary>
    public DataBinder DataBinder { get; set; }

    /// <summary>
    /// 처리 옵션
    /// </summary>
    public PowerPointOptions Options { get; set; }

    /// <summary>
    /// 현재 처리 단계
    /// </summary>
    public ProcessingPhase CurrentPhase { get; set; } = ProcessingPhase.Initialization; public PPTContext(PresentationDocument templateDocument, PowerPointOptions options)
    {
        TemplateDocument = templateDocument ?? throw new ArgumentNullException(nameof(templateDocument));
        Options = options ?? throw new ArgumentNullException(nameof(options));
        Functions = new PPTFunctions(Variables);
        DataBinder = new DataBinder();
    }/// <summary>
     /// 변수 추가
     /// </summary>
    public void AddVariable(string name, object value)
    {
        Variables[name] = value;
        // PPTFunctions 업데이트 (기존 캐시 보존)
        UpdatePPTFunctions();
    }    /// <summary>
         /// 여러 변수 일괄 추가
         /// </summary>
    public void AddVariables(Dictionary<string, object> variables)
    {
        foreach (var kvp in variables)
        {
            Variables[kvp.Key] = kvp.Value;
        }
        // PPTFunctions 업데이트 (기존 캐시 보존)
        UpdatePPTFunctions();
    }

    /// <summary>
    /// PPTFunctions 인스턴스 업데이트 (이미지 캐시 보존)
    /// </summary>
    private void UpdatePPTFunctions()
    {
        // 기존 이미지 캐시 보존
        var existingImageCache = Functions?.GetAllImageCache() ?? new Dictionary<string, string>();

        // 새로운 PPTFunctions 인스턴스 생성
        Functions = new PPTFunctions(Variables);

        // 기존 이미지 캐시를 새 인스턴스에 복원
        if (existingImageCache.Any())
        {
            Functions.RestoreImageCache(existingImageCache);
        }

        Variables["ppt"] = Functions;
    }

    /// <summary>
    /// 특정 슬라이드의 컨텍스트 생성
    /// </summary>
    public SlideContext CreateSlideContext(int slideIndex, SlideInstance slideInstance)
    {
        return new SlideContext(this, slideIndex, slideInstance);
    }
}

/// <summary>
/// PowerPoint 처리 단계 열거형
/// </summary>
public enum ProcessingPhase
{
    Initialization,      // 초기화
    TemplateAnalysis,    // 템플릿 분석
    AliasTransformation, // Alias 표현식 변환
    PlanGeneration,      // 슬라이드 계획 생성
    ExpressionBinding,   // 표현식 바인딩
    DataBinding,         // 데이터 바인딩
    FunctionProcessing,  // 함수 처리 (이미지 등)
    Finalization        // 최종화
}
