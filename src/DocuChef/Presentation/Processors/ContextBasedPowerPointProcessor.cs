using DocumentFormat.OpenXml.Packaging;
using DocuChef.Logging;
using DocuChef.Presentation.Context;
using DocuChef.Presentation.Models;
using DocuChef.Presentation.Processors;
using DocuChef.Presentation.Functions;

namespace DocuChef.Presentation.Processors;

/// <summary>
/// 컨텍스트 기반 PowerPoint 처리기
/// PPTContext와 SlideContext를 활용한 명확한 책임 분리
/// </summary>
public class ContextBasedPowerPointProcessor
{
    private readonly TemplateAnalyzer _templateAnalyzer;
    private readonly SlidePlanGenerator _planGenerator;
    private readonly SlideGenerator _slideGenerator;
    private readonly DataBinder _dataBinder;

    public ContextBasedPowerPointProcessor()
    {
        _templateAnalyzer = new TemplateAnalyzer();
        _planGenerator = new SlidePlanGenerator();
        _slideGenerator = new SlideGenerator();
        _dataBinder = new DataBinder();
    }

    /// <summary>
    /// PowerPoint 문서 생성 전체 프로세스
    /// </summary>
    public IDish ProcessPresentation(PresentationDocument templateDocument, PowerPointOptions options,
        Dictionary<string, object> variables)
    {        // 1. PPTContext 초기화
        var context = new PPTContext(templateDocument, options);
        context.AddVariables(variables);
        context.CurrentPhase = ProcessingPhase.TemplateAnalysis;

        // 작업용 문서 생성
        context.WorkingDocument = CreateWorkingDocument(templateDocument);

        if (options.EnableVerboseLogging)
        {
            Logger.Debug("Starting context-based PowerPoint processing");
        }

        try
        {
            // 2. 템플릿 분석 - SlideInfo List 구성
            AnalyzeTemplate(context);

            // 3. 슬라이드 계획 생성 - 바인딩할 데이터를 기반으로 range, foreach 복제 고려
            GenerateSlidePlan(context);

            // 4. 슬라이드 생성 및 표현식 변환
            GenerateSlides(context);

            // 5. 데이터 바인딩 처리
            ProcessDataBinding(context);

            // 6. 함수 처리 (이미지 등)
            ProcessFunctions(context);

            // 7. 최종화
            return FinalizePresentationTemplate(context);
        }
        catch (Exception ex)
        {
            Logger.Error($"Error in context-based PowerPoint processing: {ex.Message}", ex);
            throw;
        }
    }    /// <summary>
         /// 1단계: 템플릿 분석 - SlideInfo List 구성
         /// </summary>
    private void AnalyzeTemplate(PPTContext context)
    {
        context.CurrentPhase = ProcessingPhase.TemplateAnalysis;

        if (context.Options.EnableVerboseLogging)
        {
            Logger.Debug("Phase 1: Analyzing template slides");
        }

        // Use the existing AnalyzeTemplateSlides approach
        context.TemplateSlides = new List<SlideInfo>();
        var templateAnalyzer = new TemplateAnalyzer();

        if (context.TemplateDocument.PresentationPart?.Presentation?.SlideIdList == null)
        {
            if (context.Options.EnableVerboseLogging)
            {
                Logger.Debug("No slides found in template document");
            }
            return;
        }

        var slideIds = context.TemplateDocument.PresentationPart.Presentation.SlideIdList.ChildElements
            .OfType<DocumentFormat.OpenXml.Presentation.SlideId>();

        int slideIndex = 0;
        foreach (var slideId in slideIds)
        {
            try
            {
                string? relationshipId = slideId.RelationshipId?.Value;
                if (string.IsNullOrEmpty(relationshipId))
                {
                    slideIndex++;
                    continue;
                }
                var slidePart = (SlidePart)context.TemplateDocument.PresentationPart.GetPartById(relationshipId);
                var slideNotes = GetSlideNotes(slidePart);

                // Use the TemplateAnalyzer.Analyze method - pass the SlidePart instead of Slide
                var slideInfo = templateAnalyzer.Analyze(slidePart, slideNotes, slideIndex);
                slideInfo.Position = slideIndex; // Set the position property
                context.TemplateSlides.Add(slideInfo);

                if (context.Options.EnableVerboseLogging)
                {
                    Logger.Debug($"  Slide {slideIndex}: {slideInfo.BindingExpressions?.Count ?? 0} expressions, " +
                               $"{slideInfo.Directives?.Count ?? 0} directives");
                }

                slideIndex++;
            }
            catch (Exception ex)
            {
                if (context.Options.EnableVerboseLogging)
                {
                    Logger.Warning($"Error analyzing slide {slideIndex}: {ex.Message}");
                }
                slideIndex++;
            }
        }

        if (context.Options.EnableVerboseLogging)
        {
            Logger.Debug($"Template analysis complete. Found {context.TemplateSlides.Count} template slides");
        }
    }

    /// <summary>
    /// Get slide notes content
    /// </summary>
    private string GetSlideNotes(SlidePart slidePart)
    {
        try
        {
            var notesSlidePart = slidePart.NotesSlidePart;
            if (notesSlidePart?.NotesSlide?.CommonSlideData?.ShapeTree != null)
            {
                var textElements = notesSlidePart.NotesSlide.CommonSlideData.ShapeTree
                    .Descendants<DocumentFormat.OpenXml.Drawing.Text>();
                return string.Join(" ", textElements.Select(t => t.Text ?? ""));
            }
        }
        catch (Exception ex)
        {
            Logger.Warning($"Error reading slide notes: {ex.Message}");
        }
        return string.Empty;
    }

    /// <summary>
    /// 2단계: 슬라이드 계획 생성 - 데이터 기반 복제 계획
    /// </summary>
    private void GenerateSlidePlan(PPTContext context)
    {
        context.CurrentPhase = ProcessingPhase.PlanGeneration;

        if (context.Options.EnableVerboseLogging)
        {
            Logger.Debug("Phase 2: Generating slide plan based on data");
        }

        context.GenerationPlan = _planGenerator.GeneratePlan(context.TemplateSlides, context.Variables);

        if (context.Options.EnableVerboseLogging)
        {
            Logger.Debug($"Slide plan generated. Will create {context.GenerationPlan.SlideInstances.Count} slides");
            foreach (var instance in context.GenerationPlan.SlideInstances)
            {
                Logger.Debug($"  Slide {instance.Position}: Template {instance.SourceSlideId}, " +
                           $"Context: '{instance.ContextPathString}'");
            }
        }
    }

    /// <summary>
    /// 3단계: 슬라이드 생성 및 표현식 변환
    /// </summary>
    private void GenerateSlides(PPTContext context)
    {
        context.CurrentPhase = ProcessingPhase.ExpressionBinding;

        if (context.Options.EnableVerboseLogging)
        {
            Logger.Debug("Phase 3: Generating slides and transforming expressions");
        }

        // 작업용 문서 생성
        context.WorkingDocument = CreateWorkingDocument(context.TemplateDocument);        // 슬라이드 생성 및 표현식 변환
        _slideGenerator.GenerateSlides(context.WorkingDocument, context.GenerationPlan, context.SlideInfos, context.Variables);

        if (context.Options.EnableVerboseLogging)
        {
            Logger.Debug("Slide generation and expression transformation complete");
        }
    }

    /// <summary>
    /// 4단계: 데이터 바인딩 처리
    /// </summary>
    private void ProcessDataBinding(PPTContext context)
    {
        context.CurrentPhase = ProcessingPhase.DataBinding;

        if (context.Options.EnableVerboseLogging)
        {
            Logger.Debug("Phase 4: Processing data binding");
        }

        var presentationPart = context.WorkingDocument.PresentationPart;
        if (presentationPart?.Presentation?.SlideIdList == null)
        {
            Logger.Warning("No slides found for data binding");
            return;
        }

        var slideIds = presentationPart.Presentation.SlideIdList.ChildElements
            .OfType<DocumentFormat.OpenXml.Presentation.SlideId>().ToList();

        // 각 슬라이드에 대해 컨텍스트 생성하고 바인딩 처리
        for (int i = 0; i < slideIds.Count && i < context.GenerationPlan.SlideInstances.Count; i++)
        {
            var slideId = slideIds[i];
            var slideInstance = context.GenerationPlan.SlideInstances[i];

            if (string.IsNullOrEmpty(slideId.RelationshipId?.Value))
                continue; try
            {
                // 슬라이드 컨텍스트 생성
                var slideContext = context.CreateSlideContext(i, slideInstance);
                if (context.Options.EnableVerboseLogging)
                {
                    Logger.Debug($"Attempting to get SlidePart for slideId {i} with RelationshipId: {slideId.RelationshipId.Value}");
                }

                var slidePart = (SlidePart)presentationPart.GetPartById(slideId.RelationshipId.Value);

                if (context.Options.EnableVerboseLogging)
                {
                    Logger.Debug($"Retrieved SlidePart: {slidePart != null}, Slide: {slidePart?.Slide != null}");
                }

                slideContext.SlidePart = slidePart;

                // 이 슬라이드의 데이터 바인딩 처리
                ProcessSlideDataBinding(slideContext);

                if (context.Options.EnableVerboseLogging)
                {
                    Logger.Debug($"Data binding complete for slide {i} (context: {slideInstance.ContextPath})");
                }
            }
            catch (Exception ex)
            {
                Logger.Error($"Error processing data binding for slide {i}: {ex.Message}", ex);
            }
        }

        if (context.Options.EnableVerboseLogging)
        {
            Logger.Debug("Data binding phase complete");
        }
    }    /// <summary>
         /// Process data binding for individual slide
         /// </summary>
    private void ProcessSlideDataBinding(SlideContext slideContext)
    {
        if (slideContext.SlidePart?.Slide == null)
        {
            if (slideContext.PPTContext.Options.EnableVerboseLogging)
            {
                Logger.Debug($"Slide {slideContext.SlideIndex}: SlidePart or Slide is null, skipping data binding");
            }
            return;
        }
        var slide = slideContext.SlidePart.Slide;
        var textElements = slide.Descendants<DocumentFormat.OpenXml.Drawing.Text>().ToList();

        if (slideContext.PPTContext.Options.EnableVerboseLogging)
        {
            Logger.Debug($"Processing data binding for slide {slideContext.SlideIndex}: found {textElements.Count} text elements");

            // Debug: Check what elements exist in the slide
            var allTextElements = slide.Descendants().Where(e => e.InnerText.Contains("$")).ToList();
            Logger.Debug($"  Found {allTextElements.Count} elements containing '$':");
            foreach (var element in allTextElements)
            {
                Logger.Debug($"    {element.GetType().Name}: '{element.InnerText}'");
            }

            // Check for other text-like elements
            var runElements = slide.Descendants<DocumentFormat.OpenXml.Drawing.Run>().ToList();
            Logger.Debug($"  Found {runElements.Count} Run elements");

            var textBodyElements = slide.Descendants<DocumentFormat.OpenXml.Drawing.TextBody>().ToList();
            Logger.Debug($"  Found {textBodyElements.Count} TextBody elements");
        }
        foreach (var textElement in textElements)
        {
            if (string.IsNullOrEmpty(textElement.Text))
                continue;

            if (slideContext.PPTContext.Options.EnableVerboseLogging)
            {
                Logger.Debug($"  Processing text element: '{textElement.Text}'");
            }

            // Note: Data binding is intentionally NOT performed here
            // All data binding is handled exclusively in DataBinder.cs via DollarSignEngine
        }

        if (slideContext.PPTContext.Options.EnableVerboseLogging)
        {
            Logger.Debug($"Data binding completed for slide {slideContext.SlideIndex}");
        }
    }

    /// <summary>
    /// 5단계: 함수 처리 (이미지 등)
    /// </summary>
    private void ProcessFunctions(PPTContext context)
    {
        context.CurrentPhase = ProcessingPhase.FunctionProcessing;

        if (context.Options.EnableVerboseLogging)
        {
            Logger.Debug("Phase 5: Processing PowerPoint functions");
        }        // PowerPoint 함수 처리 (이미지 삽입 등)
        DocuChef.Presentation.Functions.PowerPointFunctionHandler.ProcessFunctions(context.WorkingDocument, context.Functions);

        if (context.Options.EnableVerboseLogging)
        {
            Logger.Debug("Function processing complete");
        }
    }

    /// <summary>
    /// 6단계: 최종화
    /// </summary>
    private IDish FinalizePresentationTemplate(PPTContext context)
    {
        context.CurrentPhase = ProcessingPhase.Finalization;

        if (context.Options.EnableVerboseLogging)
        {
            Logger.Debug("Phase 6: Finalizing presentation");
        }        // 문서 저장
        context.WorkingDocument.Save();

        // 임시 파일로 저장하여 PowerPointDocument 생성
        var tempPath = Path.GetTempFileName() + ".pptx";
        using (var stream = new FileStream(tempPath, FileMode.Create))
        {
            context.WorkingDocument.Clone(stream);
        }

        // PowerPointDocument 생성 및 반환
        var document = new PowerPointDocument(tempPath);

        if (context.Options.EnableVerboseLogging)
        {
            Logger.Debug("Context-based PowerPoint processing complete");
        }

        return document;
    }

    /// <summary>
    /// 템플릿으로부터 작업용 문서 생성
    /// </summary>
    private PresentationDocument CreateWorkingDocument(PresentationDocument template)
    {
        // 임시 파일 생성
        var tempPath = Path.GetTempFileName() + ".pptx";

        // 템플릿 복사
        using (var templateStream = new MemoryStream())
        {
            template.Clone(templateStream);
            templateStream.Position = 0;

            using (var tempFileStream = File.Create(tempPath))
            {
                templateStream.CopyTo(tempFileStream);
            }
        }

        // 작업용 문서 열기
        return PresentationDocument.Open(tempPath, true);
    }

    // Note: ExtractExpressionsFromText method removed
    // All expression extraction and data binding is handled exclusively in DataBinder.cs
}
