using DocumentFormat.OpenXml.Packaging;
using DocuChef.Exceptions;
using DocuChef.Logging;
using DocuChef.Presentation.Context;
using DocuChef.Presentation.Exceptions;
using DocuChef.Presentation.Models;
using DocuChef.Presentation.Processors;
using DocuChef.Presentation.Functions;
using DocuChef.Presentation.Utilities;
using System.Text.RegularExpressions;

namespace DocuChef.Presentation.Processors;

/// <summary>
/// 컨텍스트 기반 PowerPoint 처리기
/// PPTContext와 SlideContext를 활용한 명확한 책임 분리
/// </summary>
public class ContextBasedPowerPointProcessor
{
    private static readonly Regex DollarSignExpressionRegex = new(@"\$\{([^}]+)\}", RegexOptions.Compiled); private readonly TemplateAnalyzer _templateAnalyzer;
    private readonly SlidePlanGenerator _planGenerator;
    private readonly SlideGenerator _slideGenerator;
    private readonly DataBinder _dataBinder;
    private readonly ElementHider _elementHider;

    public ContextBasedPowerPointProcessor()
    {
        _templateAnalyzer = new TemplateAnalyzer();
        _planGenerator = new SlidePlanGenerator();
        _slideGenerator = new SlideGenerator();
        _dataBinder = new DataBinder();
        _elementHider = new ElementHider();
    }    /// <summary>
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
        {            // 1. 스마트 따옴표 전처리 - PowerPoint 특수 따옴표를 일반 따옴표로 변환
            PreprocessSmartQuotes(context);

            // 2. 템플릿 분석 - SlideInfo List 구성
            AnalyzeTemplate(context);

            // 3. Alias 표현식 변환 - 템플릿의 모든 표현식을 원래 경로로 변환
            ApplyAliasTransformations(context);

            // 4. 슬라이드 계획 생성 - 바인딩할 데이터를 기반으로 range, foreach 복제 고려
            GenerateSlidePlan(context);

            // 5. 슬라이드 생성
            GenerateSlides(context);

            // 6. 데이터 바인딩 처리
            ProcessDataBinding(context);

            // 7. 함수 처리 (이미지 등)
            ProcessFunctions(context);

            // 8. 최종화
            return FinalizePresentationTemplate(context);
        }
        catch (Exception ex)
        {
            Logger.Error($"Error in context-based PowerPoint processing: {ex.Message}", ex);
            throw;
        }
    }    /// <summary>
         /// PowerPoint 스마트 따옴표를 일반 따옴표로 변환하는 전처리
         /// ${...} 보간 문자열 내의 특수 따옴표들을 표준 따옴표로 치환
         /// </summary>
    private void PreprocessSmartQuotes(PPTContext context)
    {
        if (context.Options.EnableVerboseLogging)
        {
            Logger.Debug("Preprocessing smart quotes in template slides");
        }

        var presentationPart = context.TemplateDocument.PresentationPart;
        if (presentationPart?.Presentation?.SlideIdList == null)
        {
            return;
        }

        var slideIds = presentationPart.Presentation.SlideIdList.ChildElements
            .OfType<DocumentFormat.OpenXml.Presentation.SlideId>();

        int processedSlides = 0;
        int processedExpressions = 0;

        foreach (var slideId in slideIds)
        {
            try
            {
                string? relationshipId = slideId.RelationshipId?.Value;
                if (string.IsNullOrEmpty(relationshipId))
                    continue;

                var slidePart = (SlidePart)presentationPart.GetPartById(relationshipId);
                if (slidePart?.Slide == null)
                    continue;

                // 슬라이드의 모든 텍스트 요소에서 스마트 따옴표 전처리
                var textElements = slidePart.Slide.Descendants<DocumentFormat.OpenXml.Drawing.Text>().ToList();

                foreach (var textElement in textElements)
                {
                    if (string.IsNullOrEmpty(textElement.Text))
                        continue;

                    string originalText = textElement.Text;
                    string processedText = NormalizeQuotesInExpressions(originalText);

                    if (originalText != processedText)
                    {
                        textElement.Text = processedText;
                        processedExpressions++;

                        if (context.Options.EnableVerboseLogging)
                        {
                            Logger.Debug($"  Quote normalization: '{originalText}' → '{processedText}'");
                        }
                    }
                }

                // 슬라이드 노트에서도 스마트 따옴표 전처리
                var notesPart = slidePart.NotesSlidePart;
                if (notesPart?.NotesSlide != null)
                {
                    var notesTextElements = notesPart.NotesSlide.Descendants<DocumentFormat.OpenXml.Drawing.Text>().ToList();

                    foreach (var textElement in notesTextElements)
                    {
                        if (string.IsNullOrEmpty(textElement.Text))
                            continue;

                        string originalText = textElement.Text;
                        string processedText = NormalizeQuotesInExpressions(originalText);

                        if (originalText != processedText)
                        {
                            textElement.Text = processedText;
                            processedExpressions++;

                            if (context.Options.EnableVerboseLogging)
                            {
                                Logger.Debug($"  Quote normalization in notes: '{originalText}' → '{processedText}'");
                            }
                        }
                    }
                }

                processedSlides++;
            }
            catch (Exception ex)
            {
                if (context.Options.EnableVerboseLogging)
                {
                    Logger.Warning($"Error preprocessing quotes in slide: {ex.Message}");
                }
            }
        }

        if (context.Options.EnableVerboseLogging)
        {
            Logger.Debug($"Smart quote preprocessing complete. Processed {processedSlides} slides, normalized {processedExpressions} expressions");
        }
    }

    /// <summary>
    /// ${...} 표현식 내부의 스마트 따옴표를 일반 따옴표로 변환
    /// </summary>
    private string NormalizeQuotesInExpressions(string text)
    {
        if (string.IsNullOrEmpty(text))
            return text;

        // ${...} 패턴을 찾아서 각각 처리
        return DollarSignExpressionRegex.Replace(text, match =>
        {
            string expression = match.Value;
            string innerExpression = match.Groups[1].Value;
            // 표현식 내부의 스마트 따옴표들을 일반 따옴표로 변환
            string normalizedInner = innerExpression
                .Replace("\u201C", "\"")  // 왼쪽 스마트 큰따옴표
                .Replace("\u201D", "\"")  // 오른쪽 스마트 큰따옴표
                .Replace("\u2018", "'")   // 왼쪽 스마트 작은따옴표
                .Replace("\u2019", "'");  // 오른쪽 스마트 작은따옴표

            return "${" + normalizedInner + "}";
        });
    }

    /// <summary>
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

    private string GetSlideNotes(SlidePart slidePart)
    {
        try
        {
            var notesSlidePart = slidePart.NotesSlidePart;
            if (notesSlidePart?.NotesSlide?.CommonSlideData?.ShapeTree != null)
            {
                var textElements = notesSlidePart.NotesSlide.CommonSlideData.ShapeTree
                    .Descendants<DocumentFormat.OpenXml.Drawing.Text>()
                    .Select(t => t.Text ?? "")
                    .Where(text => !string.IsNullOrWhiteSpace(text))
                    .ToList();

                var filteredTexts = textElements
                    .Where(text => !IsSlideNumber(text.Trim()))
                    .ToList();

                var result = string.Join("", filteredTexts);

                Logger.Debug($"GetSlideNotes: Found {textElements.Count} text elements");
                Logger.Debug($"GetSlideNotes: Filtered to {filteredTexts.Count} elements");
                Logger.Debug($"GetSlideNotes: Combined result: '{result}'");

                return result;
            }
        }
        catch (Exception ex)
        {
            Logger.Warning($"Error reading slide notes: {ex.Message}");
        }
        return string.Empty;
    }

    /// <summary>
    /// Check if a text element appears to be a slide number
    /// </summary>
    private static bool IsSlideNumber(string text)
    {
        // Check if it's a number (typically 1-3 digits for slide numbers)
        return text.Length <= 3 && int.TryParse(text, out _);
    }    /// <summary>
         /// 4단계: 슬라이드 계획 생성 - 데이터 기반 복제 계획
         /// </summary>
    private void GenerateSlidePlan(PPTContext context)
    {
        context.CurrentPhase = ProcessingPhase.PlanGeneration;

        if (context.Options.EnableVerboseLogging)
        {
            Logger.Debug("Phase 4: Generating slide plan based on data");
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
    }    /// <summary>
         /// 3단계: Alias 표현식 변환 - 템플릿의 모든 표현식을 원래 경로로 변환
         /// </summary>
    private void ApplyAliasTransformations(PPTContext context)
    {
        context.CurrentPhase = ProcessingPhase.AliasTransformation;

        if (context.Options.EnableVerboseLogging)
        {
            Logger.Debug("Phase 3: Applying alias transformations");
        }

        // Build alias map from all template slides
        var aliasMap = BuildAliasMap(context.TemplateSlides);

        if (aliasMap.Count == 0)
        {
            if (context.Options.EnableVerboseLogging)
            {
                Logger.Debug("No aliases found to apply");
            }
            return;
        }

        if (context.Options.EnableVerboseLogging)
        {
            Logger.Debug($"Found {aliasMap.Count} aliases to apply:");
            foreach (var alias in aliasMap)
            {
                Logger.Debug($"  {alias.Key} -> {alias.Value}");
            }
        }

        // Apply aliases to all slides using ExpressionUpdater
        var expressionUpdater = new ExpressionUpdater();
        var totalTransformations = 0;

        // Transform expressions in working document
        var workingSlides = context.WorkingDocument.PresentationPart?.Presentation?.SlideIdList?.ChildElements
            .OfType<DocumentFormat.OpenXml.Presentation.SlideId>().ToList();

        if (workingSlides != null)
        {
            for (int slideIndex = 0; slideIndex < workingSlides.Count; slideIndex++)
            {
                var slideId = workingSlides[slideIndex];
                string? relationshipId = slideId.RelationshipId?.Value;
                if (string.IsNullOrEmpty(relationshipId)) continue;

                var slidePart = (SlidePart)context.WorkingDocument.PresentationPart!.GetPartById(relationshipId);

                if (context.Options.EnableVerboseLogging)
                {
                    Logger.Debug($"Applying aliases to slide {slideIndex + 1}");
                }

                var transformedCount = TransformSlideExpressions(slidePart, aliasMap, context.Options.EnableVerboseLogging);
                totalTransformations += transformedCount;

                if (context.Options.EnableVerboseLogging && transformedCount > 0)
                {
                    Logger.Debug($"  Transformed {transformedCount} expressions in slide {slideIndex + 1}");
                }
            }
        }

        // Also update the template slides' BindingExpressions for plan generation
        foreach (var templateSlide in context.TemplateSlides)
        {
            if (templateSlide.BindingExpressions != null)
            {
                foreach (var bindingExpression in templateSlide.BindingExpressions)
                {
                    if (!string.IsNullOrEmpty(bindingExpression.OriginalExpression))
                    {
                        var originalExpression = bindingExpression.OriginalExpression;
                        var transformedExpression = expressionUpdater.ApplyAliases(originalExpression, aliasMap);
                        if (originalExpression != transformedExpression)
                        {
                            bindingExpression.OriginalExpression = transformedExpression;

                            if (context.Options.EnableVerboseLogging)
                            {
                                Logger.Debug($"  Template binding: {originalExpression} -> {transformedExpression}");
                            }
                        }
                    }

                    // Also transform DataPath if it uses aliases
                    if (!string.IsNullOrEmpty(bindingExpression.DataPath))
                    {
                        var originalDataPath = bindingExpression.DataPath;
                        var transformedDataPath = TransformDataPath(originalDataPath, aliasMap);

                        if (originalDataPath != transformedDataPath)
                        {
                            bindingExpression.DataPath = transformedDataPath;

                            if (context.Options.EnableVerboseLogging)
                            {
                                Logger.Debug($"  Template DataPath: {originalDataPath} -> {transformedDataPath}");
                            }
                        }
                    }
                }
            }
        }

        if (context.Options.EnableVerboseLogging)
        {
            Logger.Debug($"Alias transformation complete. Total transformations: {totalTransformations}");
        }
    }    /// <summary>
         /// 5단계: 슬라이드 생성
         /// </summary>
    private void GenerateSlides(PPTContext context)
    {
        context.CurrentPhase = ProcessingPhase.ExpressionBinding;

        if (context.Options.EnableVerboseLogging)
        {
            Logger.Debug("Phase 5: Generating slides");
        }

        // 슬라이드 생성 (alias는 이미 이전 단계에서 적용됨)
        _slideGenerator.GenerateSlides(context.WorkingDocument, context.GenerationPlan, context.SlideInfos, context.Variables, new Dictionary<string, string>());

        if (context.Options.EnableVerboseLogging)
        {
            Logger.Debug("Slide generation and expression transformation complete");
        }
    }    /// <summary>
         /// 6단계: 데이터 바인딩 처리
         /// </summary>
    private void ProcessDataBinding(PPTContext context)
    {
        context.CurrentPhase = ProcessingPhase.DataBinding;

        if (context.Options.EnableVerboseLogging)
        {
            Logger.Debug("Phase 6: Processing data binding");
        }

        var presentationPart = context.WorkingDocument.PresentationPart;
        if (presentationPart?.Presentation?.SlideIdList == null)
        {
            Logger.Warning("No slides found for data binding");
            return;
        }        // PERFORMANCE OPTIMIZATION: Prepare base variables once for all slides using context's DataBinder
        context.DataBinder.PrepareBaseVariables(context.Variables);

        if (context.Options.EnableVerboseLogging)
        {
            Logger.Debug("Base variables prepared for optimized data binding");
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

        // Process data binding at paragraph level to handle Korean text properly
        var paragraphs = slide.Descendants<DocumentFormat.OpenXml.Drawing.Paragraph>().ToList();

        if (slideContext.PPTContext.Options.EnableVerboseLogging)
        {
            Logger.Debug($"Processing data binding for slide {slideContext.SlideIndex}: found {paragraphs.Count} paragraphs");
        }

        foreach (var paragraph in paragraphs)
        {
            ProcessParagraphDataBinding(paragraph, slideContext);
        }

        if (slideContext.PPTContext.Options.EnableVerboseLogging)
        {
            Logger.Debug($"Data binding completed for slide {slideContext.SlideIndex}");

            // Log final bound results for verification
            LogFinalSlideContent(slideContext);
        }
    }    /// <summary>
         /// Process data binding for a single paragraph with enhanced formatting preservation
         /// </summary>
    private void ProcessParagraphDataBinding(DocumentFormat.OpenXml.Drawing.Paragraph paragraph, SlideContext slideContext)
    {
        var textElements = paragraph.Descendants<DocumentFormat.OpenXml.Drawing.Text>().ToList();
        if (textElements.Count == 0)
            return;

        // Extract the complete paragraph text
        var paragraphText = string.Join("", textElements.Select(t => t.Text));

        if (slideContext.PPTContext.Options.EnableVerboseLogging)
        {
            Logger.Debug($"  Processing paragraph: '{paragraphText}'");
        }

        // Check if this paragraph contains any binding expressions
        if (string.IsNullOrEmpty(paragraphText) || !paragraphText.Contains("${"))
            return;

        // Enhanced strategy: Process complete expressions first, then handle incomplete ones
        var processedText = ProcessCompleteAndIncompleteExpressions(paragraph, textElements, paragraphText, slideContext);

        if (processedText != paragraphText)
        {
            if (slideContext.PPTContext.Options.EnableVerboseLogging)
            {
                Logger.Debug($"  Paragraph bound from '{paragraphText}' to '{processedText}'");
            }
        }
    }

    /// <summary>
    /// Enhanced expression processing with formatting preservation strategy
    /// 1. Process complete ${...} expressions within single spans first
    /// 2. Handle incomplete expressions by connecting subsequent spans
    /// </summary>
    private string ProcessCompleteAndIncompleteExpressions(
        DocumentFormat.OpenXml.Drawing.Paragraph paragraph,
        IList<DocumentFormat.OpenXml.Drawing.Text> textElements,
        string paragraphText,
        SlideContext slideContext)
    {
        // First pass: Process complete expressions within single spans
        var modifiedElements = new HashSet<DocumentFormat.OpenXml.Drawing.Text>();

        Logger.Debug($"ProcessCompleteAndIncompleteExpressions: Starting first pass with {textElements.Count} text elements");

        foreach (var textElement in textElements)
        {
            var elementText = textElement.Text ?? "";
            if (string.IsNullOrEmpty(elementText))
                continue;

            Logger.Debug($"ProcessCompleteAndIncompleteExpressions: Checking element text: '{elementText}'");

            // Check for complete expressions in this element
            var completeExpressions = DollarSignExpressionRegex.Matches(elementText);
            Logger.Debug($"ProcessCompleteAndIncompleteExpressions: Found {completeExpressions.Count} complete expressions in element");

            if (completeExpressions.Count > 0)
            {
                try
                {
                    Logger.Debug($"ProcessCompleteAndIncompleteExpressions: Processing complete expressions in element: '{elementText}'");
                    var processedElementText = ProcessElementExpressions(elementText, slideContext);
                    if (processedElementText != elementText)
                    {
                        textElement.Text = processedElementText;
                        modifiedElements.Add(textElement);

                        if (slideContext.PPTContext.Options.EnableVerboseLogging)
                        {
                            Logger.Debug($"    Processed complete expression in span: '{elementText}' → '{processedElementText}'");
                        }
                    }
                }
                catch (ElementHideException ex)
                {
                    Logger.Debug($"Array bounds exceeded, setting element to empty string: {ex.Message}");
                    // Set element text to empty string instead of hiding
                    textElement.Text = "";
                    modifiedElements.Add(textElement);
                    Logger.Debug($"    Set element text to empty string due to array bounds: '{elementText}' → ''");
                }
            }
        }

        // Second pass: Handle incomplete expressions spanning multiple elements
        var updatedParagraphText = string.Join("", textElements.Select(t => t.Text));

        Logger.Debug($"ProcessCompleteAndIncompleteExpressions: Second pass - updatedParagraphText: '{updatedParagraphText}'");

        // Check if there are still incomplete expressions or complete expressions that weren't processed
        if (HasIncompleteExpressions(updatedParagraphText))
        {
            Logger.Debug($"ProcessCompleteAndIncompleteExpressions: Has incomplete expressions, processing...");
            ProcessIncompleteExpressions(paragraph, textElements, slideContext, modifiedElements);
        }
        else if (DollarSignExpressionRegex.IsMatch(updatedParagraphText))
        {
            // Even if expressions are "complete", they might be fragmented across elements
            // Process them as incomplete expressions to properly bind them
            Logger.Debug($"ProcessCompleteAndIncompleteExpressions: Found complete expressions that were fragmented, processing as incomplete...");
            ProcessIncompleteExpressions(paragraph, textElements, slideContext, modifiedElements);
        }
        else
        {
            Logger.Debug($"ProcessCompleteAndIncompleteExpressions: No expressions found to process");
        }

        // Return the final processed text
        var finalText = string.Join("", textElements.Select(t => t.Text));
        Logger.Debug($"ProcessCompleteAndIncompleteExpressions: Final text: '{finalText}'");        // Check if all text elements are empty and remove the paragraph if so
        if (string.IsNullOrWhiteSpace(finalText))
        {
            // DISABLED: Don't automatically remove empty paragraphs
            // Individual line-break separated elements should be preserved even if they become empty
            // Only remove in very specific cases (e.g., template errors, not data-driven emptiness)
            Logger.Debug($"ProcessCompleteAndIncompleteExpressions: Text is empty, but keeping paragraph to preserve document structure");
        }

        return finalText;
    }    /// <summary>
         /// Process expressions within a single text element
         /// </summary>
    private string ProcessElementExpressions(string elementText, SlideContext slideContext)
    {
        try
        {
            var indexOffset = slideContext.SlideInstance?.IndexOffset ?? 0;
            var usedExpressions = new HashSet<string> { elementText };

            // PERFORMANCE OPTIMIZATION: Use pre-cached variables from SlideContext instead of recreating them
            var variables = slideContext.GetCachedVariables();

            return slideContext.PPTContext.DataBinder.BindData(elementText, variables, usedExpressions, null, indexOffset, slideContext.ContextPath);
        }
        catch (DocuChefHideException ex)
        {
            Logger.Debug($"ProcessElementExpressions: Array bounds exceeded, returning empty string: {ex.Message}");
            return string.Empty;
        }
        catch (ElementHideException ex)
        {
            Logger.Debug($"ProcessElementExpressions: Element should be empty due to: {ex.Message}");
            return string.Empty;
        }
    }    /// <summary>
         /// Check if text contains incomplete expressions (e.g., "${" without matching "}")
         /// </summary>
    private bool HasIncompleteExpressions(string text)
    {
        // Count opening patterns "${" and closing patterns "}"
        var openBraces = 0;
        var closeBraces = text.Count(c => c == '}');

        // Count "${" patterns properly
        for (int i = 0; i < text.Length - 1; i++)
        {
            if (text[i] == '$' && text[i + 1] == '{')
            {
                openBraces++;
            }
        }

        // Check for incomplete patterns:
        // 1. "${" at the end without closing "}"
        // 2. More opening than closing braces
        // 3. Contains "${" but text is fragmented (doesn't form complete expressions)
        var hasIncomplete = text.Contains("${") &&
                           (openBraces > closeBraces ||
                            text.EndsWith("${") ||
                            (openBraces == closeBraces && !IsCompleteExpression(text)));

        Logger.Debug($"HasIncompleteExpressions: text='{text}', openBraces={openBraces}, closeBraces={closeBraces}, hasIncomplete={hasIncomplete}");

        return hasIncomplete;
    }    /// <summary>
         /// Check if the text contains only complete, well-formed expressions
         /// </summary>
    private bool IsCompleteExpression(string text)
    {
        Logger.Debug($"IsCompleteExpression: Checking text='{text}'");

        // Use regex to find complete expressions
        var completeExpressionPattern = @"\$\{[^{}]*\}";
        var matches = System.Text.RegularExpressions.Regex.Matches(text, completeExpressionPattern);

        Logger.Debug($"IsCompleteExpression: Found {matches.Count} regex matches");

        if (matches.Count == 0)
        {
            Logger.Debug($"IsCompleteExpression: No matches found, returning false");
            return false;
        }

        // Check if the entire text is covered by complete expressions (allowing whitespace)
        var coveredLength = matches.Cast<System.Text.RegularExpressions.Match>()
                                  .Sum(m => m.Length);
        var textWithoutWhitespace = text.Replace(" ", "").Replace("\n", "").Replace("\r", "").Replace("\t", "");

        var isComplete = coveredLength >= textWithoutWhitespace.Length;
        Logger.Debug($"IsCompleteExpression: coveredLength={coveredLength}, textWithoutWhitespace.Length={textWithoutWhitespace.Length}, isComplete={isComplete}");

        return isComplete;
    }

    /// <summary>
    /// Process incomplete expressions that span multiple text elements
    /// </summary>
    private void ProcessIncompleteExpressions(
        DocumentFormat.OpenXml.Drawing.Paragraph paragraph,
        IList<DocumentFormat.OpenXml.Drawing.Text> textElements,
        SlideContext slideContext,
        HashSet<DocumentFormat.OpenXml.Drawing.Text> alreadyModified)
    {
        var fullText = string.Join("", textElements.Select(t => t.Text));
        var matches = DollarSignExpressionRegex.Matches(fullText);

        Logger.Debug($"ProcessIncompleteExpressions: fullText='{fullText}', matches.Count={matches.Count}");

        if (matches.Count == 0)
            return;

        foreach (Match match in matches)
        {
            var expression = match.Value;
            var startIndex = match.Index;
            var endIndex = startIndex + match.Length;

            Logger.Debug($"ProcessIncompleteExpressions: Found expression '{expression}' at [{startIndex}, {endIndex})");

            // Find which elements this expression spans
            var spanningElements = FindSpanningElements(textElements, startIndex, endIndex);

            Logger.Debug($"ProcessIncompleteExpressions: Expression spans {spanningElements.Count} elements");

            if (spanningElements.Count > 1)
            {
                Logger.Debug($"ProcessIncompleteExpressions: Processing spanning expression '{expression}'");
                ProcessSpanningExpression(expression, spanningElements, slideContext, alreadyModified);
            }
        }
    }

    /// <summary>
    /// Find text elements that contain parts of an expression
    /// </summary>
    private List<(DocumentFormat.OpenXml.Drawing.Text element, int relativeStart, int relativeEnd)> FindSpanningElements(
        IList<DocumentFormat.OpenXml.Drawing.Text> textElements, int globalStart, int globalEnd)
    {
        var result = new List<(DocumentFormat.OpenXml.Drawing.Text, int, int)>();
        var currentPosition = 0;

        foreach (var element in textElements)
        {
            var elementText = element.Text ?? "";
            var elementStart = currentPosition;
            var elementEnd = currentPosition + elementText.Length;

            if (elementStart < globalEnd && elementEnd > globalStart)
            {
                var relativeStart = Math.Max(0, globalStart - elementStart);
                var relativeEnd = Math.Min(elementText.Length, globalEnd - elementStart);
                result.Add((element, relativeStart, relativeEnd));
            }

            currentPosition = elementEnd;
        }

        return result;
    }

    /// <summary>
    /// Process an expression that spans multiple text elements with formatting preservation
    /// </summary>
    private void ProcessSpanningExpression(
        string expression,
        List<(DocumentFormat.OpenXml.Drawing.Text element, int relativeStart, int relativeEnd)> spanningElements,
        SlideContext slideContext,
        HashSet<DocumentFormat.OpenXml.Drawing.Text> alreadyModified)
    {
        if (spanningElements.Count == 0)
            return;

        // CRITICAL: Ensure variables are cached for nested context expressions
        // This is needed because spanning expressions may not trigger GetCachedVariables() in the first pass
        Logger.Debug($"ProcessSpanningExpression: Ensuring variables are cached for context '{slideContext.ContextPath}'");
        slideContext.GetCachedVariables();

        // Process the expression
        var processedExpression = ProcessElementExpressions(expression, slideContext);

        if (processedExpression == expression)
            return; // No change needed

        // Strategy: Preserve formatting by intelligently distributing the processed text
        var isFirstElement = true;
        var remainingText = processedExpression;

        foreach (var (element, relativeStart, relativeEnd) in spanningElements)
        {
            if (alreadyModified.Contains(element))
                continue;

            var originalElement = element.Text ?? "";

            if (isFirstElement)
            {
                // First element: keep prefix, replace expression part, handle distribution
                var prefix = originalElement.Substring(0, relativeStart);
                var suffix = originalElement.Substring(relativeEnd);

                // Intelligent text distribution based on formatting patterns
                var distributedText = DistributeTextIntelligently(remainingText, spanningElements.Count, 0);

                element.Text = prefix + distributedText + suffix;
                remainingText = remainingText.Substring(Math.Min(distributedText.Length, remainingText.Length));
                isFirstElement = false;

                if (slideContext.PPTContext.Options.EnableVerboseLogging)
                {
                    Logger.Debug($"    Updated first spanning element: '{originalElement}' → '{element.Text}'");
                }
            }
            else
            {
                // Subsequent elements: distribute remaining text or clear
                if (!string.IsNullOrEmpty(remainingText))
                {
                    var elementIndex = spanningElements.FindIndex(se => se.element == element);
                    var distributedText = DistributeTextIntelligently(remainingText, spanningElements.Count, elementIndex);

                    element.Text = distributedText;
                    remainingText = remainingText.Substring(Math.Min(distributedText.Length, remainingText.Length));

                    if (slideContext.PPTContext.Options.EnableVerboseLogging)
                    {
                        Logger.Debug($"    Updated spanning element {elementIndex}: '{originalElement}' → '{element.Text}'");
                    }
                }
                else
                {
                    element.Text = "";
                }
            }

            alreadyModified.Add(element);
        }
    }

    /// <summary>
    /// Intelligently distribute text across elements to preserve formatting intention
    /// </summary>
    private string DistributeTextIntelligently(string text, int totalElements, int currentElementIndex)
    {
        if (string.IsNullOrEmpty(text) || totalElements <= 1)
            return text;

        // For the pattern "BOLD {content} Italic", try to preserve structure
        if (currentElementIndex == 0 && text.Length > 10)
        {
            // First element gets reasonable portion, not everything
            var firstPortionLength = Math.Min(text.Length * 2 / 3, text.Length - 5);
            return text.Substring(0, firstPortionLength);
        }
        else if (currentElementIndex == totalElements - 1)
        {
            // Last element gets the remainder
            return text;
        }
        else
        {
            // Middle elements get proportional distribution
            var portionLength = text.Length / (totalElements - currentElementIndex);
            return text.Substring(0, Math.Min(portionLength, text.Length));
        }
    }

    /// <summary>
    /// 5단계: 함수 처리 (이미지 등)
    /// </summary>
    private void ProcessFunctions(PPTContext context)
    {
        context.CurrentPhase = ProcessingPhase.FunctionProcessing; if (context.Options.EnableVerboseLogging)
        {
            Logger.Debug("Phase 7: Processing PowerPoint functions");
        }// PowerPoint 함수 처리 (이미지 삽입 등)
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

    /// <summary>
    /// Builds a map of aliases for expression transformation
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
                    Logger.Debug($"ContextBasedPowerPointProcessor: Added alias mapping '{directive.AliasName}' -> '{directive.CollectionPath}'");
                }
            }
        }

        if (aliasMap.Count > 0)
        {
            Logger.Debug($"ContextBasedPowerPointProcessor: Created alias map with {aliasMap.Count} entries");
        }

        return aliasMap;
    }    /// <summary>
         /// Transform expressions in a single slide
         /// </summary>
    private int TransformSlideExpressions(SlidePart slidePart, Dictionary<string, string> aliasMap, bool enableVerboseLogging)
    {
        var transformedCount = 0;
        var expressionUpdater = new ExpressionUpdater();        // Stage 1: Process complete expressions at text element level
        var textElements = slidePart.Slide.Descendants<DocumentFormat.OpenXml.Drawing.Text>().ToList();
        var processedElements = new HashSet<DocumentFormat.OpenXml.Drawing.Text>();

        foreach (var textElement in textElements)
        {
            if (!string.IsNullOrEmpty(textElement.Text))
            {
                var originalText = textElement.Text;
                var transformedText = expressionUpdater.ApplyAliases(originalText, aliasMap);

                if (originalText != transformedText)
                {
                    textElement.Text = transformedText;
                    transformedCount++;
                    processedElements.Add(textElement);

                    if (enableVerboseLogging)
                    {
                        Logger.Debug($"    Text element: '{originalText}' -> '{transformedText}'");
                    }
                }
            }
        }

        // Stage 2: Process incomplete expressions at paragraph level for unprocessed text
        var paragraphs = slidePart.Slide.Descendants<DocumentFormat.OpenXml.Drawing.Paragraph>().ToList();

        foreach (var paragraph in paragraphs)
        {
            var textRuns = paragraph.Descendants<DocumentFormat.OpenXml.Drawing.Text>().ToList();
            if (textRuns.Count > 1)
            {
                // Check if this paragraph has any unprocessed text elements
                var hasUnprocessedElements = textRuns.Any(t => !processedElements.Contains(t) && !string.IsNullOrEmpty(t.Text));

                if (hasUnprocessedElements)
                {
                    // Combine all text elements in the paragraph
                    var combinedText = string.Join("", textRuns.Select(t => t.Text));
                    var transformedCombinedText = expressionUpdater.ApplyAliases(combinedText, aliasMap);

                    if (combinedText != transformedCombinedText && enableVerboseLogging)
                    {
                        Logger.Debug($"    Paragraph combined: '{combinedText}' -> '{transformedCombinedText}'");
                    }

                    if (combinedText != transformedCombinedText)
                    {
                        // Update the first text element with the transformed text and clear others
                        if (textRuns.Count > 0)
                        {
                            textRuns[0].Text = transformedCombinedText;
                            for (int i = 1; i < textRuns.Count; i++)
                            {
                                textRuns[i].Text = "";
                            }
                            transformedCount++;

                            if (enableVerboseLogging)
                            {
                                Logger.Debug($"    Paragraph transformation applied to {textRuns.Count} text elements");
                            }
                        }
                    }
                }
            }
        }

        return transformedCount;
    }

    /// <summary>
    /// Transform a data path using alias mapping
    /// </summary>
    private string TransformDataPath(string dataPath, Dictionary<string, string> aliasMap)
    {
        foreach (var alias in aliasMap)
        {
            var aliasName = alias.Key;
            var aliasPath = alias.Value;

            if (dataPath.StartsWith(aliasName))
            {
                var remainingPart = dataPath.Substring(aliasName.Length);
                return aliasPath + remainingPart;
            }
        }

        return dataPath;
    }

    /// <summary>
    /// Log final slide content for verification
    /// </summary>
    private void LogFinalSlideContent(SlideContext slideContext)
    {
        try
        {
            var slidePart = slideContext.SlidePart;
            if (slidePart?.Slide == null) return;

            var paragraphs = slidePart.Slide.Descendants<DocumentFormat.OpenXml.Drawing.Paragraph>().ToList();
            var slideContent = new List<string>();

            foreach (var paragraph in paragraphs)
            {
                var textElements = paragraph.Descendants<DocumentFormat.OpenXml.Drawing.Text>().ToList();
                if (textElements.Any())
                {
                    var paragraphText = string.Join("", textElements.Select(t => t.Text));
                    if (!string.IsNullOrWhiteSpace(paragraphText))
                    {
                        slideContent.Add(paragraphText);
                    }
                }
            }

            if (slideContent.Any())
            {
                Logger.Debug($"[SLIDE {slideContext.SlideIndex} FINAL CONTENT]:");
                foreach (var content in slideContent)
                {
                    Logger.Debug($"  → {content}");
                }
            }
        }
        catch (Exception ex)
        {
            Logger.Debug($"Error logging slide content: {ex.Message}");
        }
    }

    /// <summary>
    /// Check if text contains only expressions (no literal text content)
    /// </summary>
    private bool HasOnlyExpressions(string text)
    {
        if (string.IsNullOrEmpty(text))
            return true;

        // Remove all expressions from text
        var textWithoutExpressions = DollarSignExpressionRegex.Replace(text, "");

        // If what remains is only whitespace, then the text contained only expressions
        return string.IsNullOrWhiteSpace(textWithoutExpressions);
    }
}
