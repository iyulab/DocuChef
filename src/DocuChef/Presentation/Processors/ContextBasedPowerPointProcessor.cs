using DocumentFormat.OpenXml.Packaging;
using DocuChef.Logging;
using DocuChef.Presentation.Context;
using DocuChef.Presentation.Models;
using DocuChef.Presentation.Processors;
using DocuChef.Presentation.Functions;
using DocuChef.Presentation.Utilities;

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
        {            // 2. 템플릿 분석 - SlideInfo List 구성
            AnalyzeTemplate(context);

            // 3. Alias 표현식 변환 - 템플릿의 모든 표현식을 원래 경로로 변환
            ApplyAliasTransformations(context);

            // 4. 슬라이드 계획 생성 - 바인딩할 데이터를 기반으로 range, foreach 복제 고려
            GenerateSlidePlan(context);

            // 5. 슬라이드 생성
            GenerateSlides(context);

            // 6. 데이터 바인딩 처리
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
    }    /// <summary>
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
                    .Descendants<DocumentFormat.OpenXml.Drawing.Text>()
                    .Select(t => t.Text ?? "")
                    .Where(text => !string.IsNullOrWhiteSpace(text))
                    .ToList();

                // Filter out text that looks like slide numbers (single digits)
                var filteredTexts = textElements
                    .Where(text => !IsSlideNumber(text.Trim()))
                    .ToList();

                return string.Join("", filteredTexts); // Join without spaces to preserve original structure
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
        }
    }

    /// <summary>
    /// Process data binding for a single paragraph, handling Korean text that may be split across spans
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
        if (!paragraphText.Contains("${"))
            return;        // Apply data binding to the complete paragraph text
        var indexOffset = slideContext.SlideInstance?.IndexOffset ?? 0;
        var usedExpressions = new HashSet<string> { paragraphText };
        var boundText = _dataBinder.BindData(paragraphText, slideContext.BindingData, usedExpressions, indexOffset);

        if (boundText != paragraphText)
        {
            if (slideContext.PPTContext.Options.EnableVerboseLogging)
            {
                Logger.Debug($"  Paragraph bound from '{paragraphText}' to '{boundText}'");
            }            // Replace the text while preserving formatting as much as possible
            if (slideContext.SlidePart != null)
            {
                ReplaceTextInParagraph(paragraph, paragraphText, boundText);
            }
        }
    }    /// <summary>
         /// Replace text in a paragraph while preserving formatting
         /// </summary>
    private void ReplaceTextInParagraph(DocumentFormat.OpenXml.Drawing.Paragraph paragraph, string oldText, string newText)
    {
        var currentText = ExtractParagraphText(paragraph);
        if (!currentText.Contains(oldText))
            return;

        var textElements = paragraph.Descendants<DocumentFormat.OpenXml.Drawing.Text>().ToList();

        // If the old text spans multiple text elements, we need a more sophisticated approach
        if (textElements.Count == 1)
        {
            // Simple case: text is in one element
            var textElement = textElements[0];
            if (textElement.Text.Contains(oldText))
            {
                textElement.Text = textElement.Text.Replace(oldText, newText);
            }
        }
        else
        {
            // Complex case: text spans multiple elements
            // Enhanced approach to preserve formatting when replacing text
            ReplaceTextPreservingFormatting(paragraph, textElements, currentText, oldText, newText);
        }
    }

    /// <summary>
    /// Extracts text from a paragraph, handling Korean text that may be split across spans
    /// </summary>
    private string ExtractParagraphText(DocumentFormat.OpenXml.Drawing.Paragraph paragraph)
    {
        var textBuilder = new System.Text.StringBuilder();
        var textElements = paragraph.Descendants<DocumentFormat.OpenXml.Drawing.Text>();

        foreach (var textElement in textElements)
        {
            if (!string.IsNullOrEmpty(textElement.Text))
            {
                textBuilder.Append(textElement.Text);
            }
        }

        return textBuilder.ToString();
    }

    /// <summary>
    /// Replaces text while preserving the original formatting structure
    /// </summary>
    private void ReplaceTextPreservingFormatting(DocumentFormat.OpenXml.Drawing.Paragraph paragraph,
        IList<DocumentFormat.OpenXml.Drawing.Text> textElements,
        string currentText, string oldText, string newText)
    {
        // Find the position where oldText starts and ends in the combined text
        var oldTextIndex = currentText.IndexOf(oldText);
        if (oldTextIndex == -1)
            return;

        var oldTextEndIndex = oldTextIndex + oldText.Length;

        // Calculate the position of each text element in the combined text
        var elementPositions = new List<(DocumentFormat.OpenXml.Drawing.Text element, int start, int end)>();
        var currentPosition = 0;

        foreach (var textElement in textElements)
        {
            var elementText = textElement.Text ?? "";
            var elementStart = currentPosition;
            var elementEnd = currentPosition + elementText.Length;
            elementPositions.Add((textElement, elementStart, elementEnd));
            currentPosition = elementEnd;
        }

        // Determine which elements need to be modified
        var elementsToModify = elementPositions
            .Where(ep => ep.start < oldTextEndIndex && ep.end > oldTextIndex)
            .ToList();

        if (elementsToModify.Count == 0)
            return;

        // If the replacement affects only one element, handle it simply
        if (elementsToModify.Count == 1)
        {
            var element = elementsToModify[0].element;
            var elementStart = elementsToModify[0].start;
            var relativeOldStart = Math.Max(0, oldTextIndex - elementStart);
            var relativeOldEnd = Math.Min(element.Text.Length, oldTextEndIndex - elementStart);

            if (relativeOldStart < element.Text.Length && relativeOldEnd > relativeOldStart)
            {
                var before = element.Text.Substring(0, relativeOldStart);
                var after = element.Text.Substring(relativeOldEnd);
                element.Text = before + newText + after;
            }
        }
        else
        {
            // Multiple elements are affected - this is the complex case
            // Strategy: Keep original formatting structure as much as possible

            // Clear all affected elements first
            foreach (var (element, _, _) in elementsToModify)
            {
                element.Text = "";
            }

            // Try to distribute the new text while preserving some formatting structure
            if (elementsToModify.Count >= 2)
            {
                var firstElement = elementsToModify[0].element;
                var lastElement = elementsToModify[elementsToModify.Count - 1].element;

                // If we can reasonably split the newText, do so
                if (newText.Contains("(") && newText.Contains(")"))
                {
                    // Special case for "Product Catalogs(2025-05-29)" pattern
                    var parenIndex = newText.IndexOf('(');
                    if (parenIndex > 0)
                    {
                        firstElement.Text = newText.Substring(0, parenIndex);
                        lastElement.Text = newText.Substring(parenIndex);
                        return;
                    }
                }

                // Fallback: put most text in first element, minimal in last
                var splitPoint = Math.Min(newText.Length, Math.Max(1, newText.Length * 2 / 3));
                firstElement.Text = newText.Substring(0, splitPoint);
                if (splitPoint < newText.Length)
                {
                    lastElement.Text = newText.Substring(splitPoint);
                }
            }
            else
            {
                // Fallback to simple replacement
                elementsToModify[0].element.Text = newText;
            }
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
        var expressionUpdater = new ExpressionUpdater();

        // First try to process complete expressions at text element level
        var textElements = slidePart.Slide.Descendants<DocumentFormat.OpenXml.Drawing.Text>().ToList();

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

                    if (enableVerboseLogging)
                    {
                        Logger.Debug($"    Text element: '{originalText}' -> '{transformedText}'");
                    }
                }
            }
        }

        // If no transformations occurred, try paragraph-level processing for split expressions
        if (transformedCount == 0)
        {
            var paragraphs = slidePart.Slide.Descendants<DocumentFormat.OpenXml.Drawing.Paragraph>().ToList();

            foreach (var paragraph in paragraphs)
            {
                var textRuns = paragraph.Descendants<DocumentFormat.OpenXml.Drawing.Text>().ToList();
                if (textRuns.Count > 1)
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
}
