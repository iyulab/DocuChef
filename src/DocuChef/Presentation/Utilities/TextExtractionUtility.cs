using DocumentFormat.OpenXml.Drawing;
using DocuChef.Logging;
using System.Text;
using DrawingText = DocumentFormat.OpenXml.Drawing.Text;

namespace DocuChef.Presentation.Utilities;

/// <summary>
/// 집중화된 텍스트 추출 유틸리티
/// Span 단위와 Paragraph 단위의 텍스트 추출을 담당
/// Span이 우선되는 계층적 추출 전략 구현
/// </summary>
public static class TextExtractionUtility
{    /// <summary>
     /// Span 단위 텍스트 추출 (최우선)
     /// 각 Text 요소를 개별적으로 추출하여 서식 정보 보존
     /// </summary>
     /// <param name="paragraph">추출할 단락</param>
     /// <param name="enableVerboseLogging">상세 로깅 활성화</param>
     /// <returns>Span 정보와 텍스트 목록</returns>
    public static List<SpanTextInfo> ExtractSpanTexts(Paragraph paragraph, bool enableVerboseLogging = false)
    {
        var spanTexts = new List<SpanTextInfo>();
        var textElements = paragraph.Descendants<DrawingText>().ToList();

        for (int i = 0; i < textElements.Count; i++)
        {
            var textElement = textElements[i];
            var text = textElement.Text ?? "";

            if (!string.IsNullOrEmpty(text))
            {
                var spanInfo = new SpanTextInfo
                {
                    Index = i,
                    Text = text,
                    Element = textElement,
                    HasCompleteExpression = HasCompleteExpression(text),
                    HasIncompleteExpression = HasIncompleteExpression(text),
                    ExpressionFragments = ExtractExpressionFragments(text)
                };

                spanTexts.Add(spanInfo);

                if (enableVerboseLogging)
                {
                    Logger.Debug($"SlideTextExtractor: Extracted span text [{i}]: '{text}' " +
                               $"(Complete: {spanInfo.HasCompleteExpression}, Incomplete: {spanInfo.HasIncompleteExpression})");
                }
            }
        }

        return spanTexts;
    }

    /// <summary>
    /// Paragraph 단위 텍스트 추출 (보조)
    /// 전체 단락 텍스트를 하나로 결합하여 추출
    /// </summary>
    /// <param name="paragraph">추출할 단락</param>
    /// <param name="enableVerboseLogging">상세 로깅 활성화</param>
    /// <returns>단락 텍스트 정보</returns>
    public static ParagraphTextInfo ExtractParagraphText(Paragraph paragraph, bool enableVerboseLogging = false)
    {
        var textElements = paragraph.Descendants<DrawingText>().ToList();
        var combinedText = string.Join("", textElements.Select(t => t.Text ?? ""));

        var paragraphInfo = new ParagraphTextInfo
        {
            Text = combinedText,
            SpanCount = textElements.Count,
            HasExpressions = combinedText.Contains("${"),
            CompleteExpressions = ExtractCompleteExpressions(combinedText),
            TextElements = textElements
        };

        if (enableVerboseLogging)
        {
            Logger.Debug($"SlideTextExtractor: Extracted paragraph text: '{combinedText}' " +
                       $"(Spans: {paragraphInfo.SpanCount}, Expressions: {paragraphInfo.CompleteExpressions.Count})");
        }

        return paragraphInfo;
    }

    /// <summary>
    /// 계층적 텍스트 추출 전략
    /// 1. Span 단위에서 완전한 표현식 우선 처리
    /// 2. 불완전한 표현식은 Paragraph 단위에서 재구성
    /// </summary>
    /// <param name="paragraph">추출할 단락</param>
    /// <param name="enableVerboseLogging">상세 로깅 활성화</param>
    /// <returns>계층적 텍스트 정보</returns>
    public static HierarchicalTextInfo ExtractHierarchicalText(Paragraph paragraph, bool enableVerboseLogging = false)
    {
        var spanTexts = ExtractSpanTexts(paragraph, enableVerboseLogging);
        var paragraphText = ExtractParagraphText(paragraph, enableVerboseLogging);

        // Span 단위에서 완전한 표현식 찾기
        var completeSpans = spanTexts.Where(s => s.HasCompleteExpression).ToList();
        var incompleteSpans = spanTexts.Where(s => s.HasIncompleteExpression && !s.HasCompleteExpression).ToList();

        var hierarchicalInfo = new HierarchicalTextInfo
        {
            SpanTexts = spanTexts,
            ParagraphText = paragraphText,
            CompleteExpressionSpans = completeSpans,
            IncompleteExpressionSpans = incompleteSpans,
            ProcessingStrategy = DetermineProcessingStrategy(completeSpans, incompleteSpans, paragraphText)
        };

        if (enableVerboseLogging)
        {
            Logger.Debug($"SlideTextExtractor: Hierarchical analysis - Strategy: {hierarchicalInfo.ProcessingStrategy}, " +
                       $"Complete spans: {completeSpans.Count}, Incomplete spans: {incompleteSpans.Count}");
        }

        return hierarchicalInfo;
    }

    /// <summary>
    /// 완전한 표현식이 있는지 확인
    /// </summary>
    private static bool HasCompleteExpression(string text)
    {
        return System.Text.RegularExpressions.Regex.IsMatch(text, @"\$\{[^}]+\}");
    }

    /// <summary>
    /// 불완전한 표현식이 있는지 확인
    /// </summary>
    private static bool HasIncompleteExpression(string text)
    {
        // "${" 가 있지만 완전한 표현식이 아닌 경우
        return text.Contains("${") && !HasCompleteExpression(text);
    }

    /// <summary>
    /// 표현식 조각들 추출
    /// </summary>
    private static List<string> ExtractExpressionFragments(string text)
    {
        var fragments = new List<string>();

        // 완전한 표현식 추출
        var completeMatches = System.Text.RegularExpressions.Regex.Matches(text, @"\$\{[^}]+\}");
        foreach (System.Text.RegularExpressions.Match match in completeMatches)
        {
            fragments.Add(match.Value);
        }

        // 불완전한 표현식 시작 부분 추출
        var incompleteMatches = System.Text.RegularExpressions.Regex.Matches(text, @"\$\{[^}]*$");
        foreach (System.Text.RegularExpressions.Match match in incompleteMatches)
        {
            fragments.Add(match.Value);
        }

        return fragments;
    }

    /// <summary>
    /// 완전한 표현식들 추출
    /// </summary>
    private static List<string> ExtractCompleteExpressions(string text)
    {
        var expressions = new List<string>();
        var matches = System.Text.RegularExpressions.Regex.Matches(text, @"\$\{[^}]+\}");

        foreach (System.Text.RegularExpressions.Match match in matches)
        {
            expressions.Add(match.Value);
        }

        return expressions;
    }

    /// <summary>
    /// 처리 전략 결정
    /// </summary>
    private static TextProcessingStrategy DetermineProcessingStrategy(
        List<SpanTextInfo> completeSpans,
        List<SpanTextInfo> incompleteSpans,
        ParagraphTextInfo paragraphText)
    {
        if (completeSpans.Count > 0 && incompleteSpans.Count == 0)
        {
            return TextProcessingStrategy.SpanOnly;
        }
        else if (completeSpans.Count == 0 && incompleteSpans.Count > 0)
        {
            return TextProcessingStrategy.ParagraphOnly;
        }
        else if (completeSpans.Count > 0 && incompleteSpans.Count > 0)
        {
            return TextProcessingStrategy.Hybrid;
        }
        else
        {
            return TextProcessingStrategy.None;
        }
    }
}

/// <summary>
/// Span 단위 텍스트 정보
/// </summary>
public class SpanTextInfo
{
    public int Index { get; set; }
    public string Text { get; set; } = "";
    public DrawingText Element { get; set; } = null!;
    public bool HasCompleteExpression { get; set; }
    public bool HasIncompleteExpression { get; set; }
    public List<string> ExpressionFragments { get; set; } = new();
}

/// <summary>
/// Paragraph 단위 텍스트 정보
/// </summary>
public class ParagraphTextInfo
{
    public string Text { get; set; } = "";
    public int SpanCount { get; set; }
    public bool HasExpressions { get; set; }
    public List<string> CompleteExpressions { get; set; } = new();
    public IList<DrawingText> TextElements { get; set; } = new List<DrawingText>();
}

/// <summary>
/// 계층적 텍스트 정보
/// </summary>
public class HierarchicalTextInfo
{
    public List<SpanTextInfo> SpanTexts { get; set; } = new();
    public ParagraphTextInfo ParagraphText { get; set; } = null!;
    public List<SpanTextInfo> CompleteExpressionSpans { get; set; } = new();
    public List<SpanTextInfo> IncompleteExpressionSpans { get; set; } = new();
    public TextProcessingStrategy ProcessingStrategy { get; set; }

    /// <summary>
    /// 결합된 텍스트 - SpanTexts 또는 ParagraphText.Text 사용
    /// </summary>
    public string CombinedText =>
        SpanTexts.Any() ? string.Join("", SpanTexts.Select(s => s.Text)) : ParagraphText?.Text ?? "";
}

/// <summary>
/// 텍스트 처리 전략
/// </summary>
public enum TextProcessingStrategy
{
    None,           // 표현식 없음
    SpanOnly,       // Span 단위만 처리
    ParagraphOnly,  // Paragraph 단위만 처리  
    Hybrid          // Span 우선, Paragraph 보조
}
