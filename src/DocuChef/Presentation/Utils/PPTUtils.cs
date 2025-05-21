namespace DocuChef.Presentation.Utils;

internal static class PPTUtils
{
    /// <summary>
    /// Normalizes quotes in all text elements throughout the entire template
    /// </summary>
    public static void NormalizeTemplateQuotes(string filePath)
    {
        Logger.Info("Normalizing smart quotes in template...");

        try
        {
            // 수정 가능한 모드로 템플릿 열기
            using (PresentationDocument presentationDoc = PresentationDocument.Open(filePath, true))
            {
                if (presentationDoc.PresentationPart == null)
                {
                    Logger.Warning("Cannot normalize quotes: Missing presentation part");
                    return;
                }

                // Process all slides
                int normalizedCount = 0;
                bool templateModified = false;
                var slideIdList = presentationDoc.PresentationPart.Presentation.SlideIdList;
                if (slideIdList != null)
                {
                    foreach (SlideId slideId in slideIdList.Elements<SlideId>())
                    {
                        string relationshipId = slideId.RelationshipId?.Value;
                        if (string.IsNullOrEmpty(relationshipId))
                            continue;

                        try
                        {
                            SlidePart slidePart = (SlidePart)presentationDoc.PresentationPart.GetPartById(relationshipId);
                            if (slidePart == null)
                                continue;

                            // Process all text elements in this slide
                            var textElements = slidePart.Slide.Descendants<D.Text>().ToList();
                            foreach (var textElement in textElements)
                            {
                                string originalText = textElement.Text;
                                if (string.IsNullOrEmpty(originalText))
                                    continue;

                                // 표현식이 포함된 텍스트에서만 처리
                                if (originalText.Contains("[“") || originalText.Contains("”]"))
                                {
                                    // 직접 스마트 따옴표 [" 와 "] 를 표준 따옴표로 교체
                                    string normalizedText = originalText.Replace("[“", "[\"").Replace("”]", "\"]");

                                    if (normalizedText != originalText)
                                    {
                                        textElement.Text = normalizedText;
                                        normalizedCount++;
                                        templateModified = true;
                                        Logger.Debug($"Normalized smart quotes in expression: '{originalText}' -> '{normalizedText}'");
                                    }
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            Logger.Warning($"Error normalizing text in slide {slideId.Id}: {ex.Message}");
                        }
                    }
                }

                // 변경사항이 있을 때만 저장
                if (templateModified)
                {
                    Logger.Debug("Saving template with normalized quotes");
                    presentationDoc.Save();
                    Logger.Info($"Successfully normalized smart quotes in {normalizedCount} expressions");
                }
                else
                {
                    Logger.Info("No smart quotes found in expressions, template not modified");
                }
            }
        }
        catch (Exception ex)
        {
            Logger.Error($"Error normalizing template quotes: {ex.Message}", ex);
            throw new DocuChefException($"Failed to normalize template quotes: {ex.Message}", ex);
        }
    }

    /// <summary>
    /// Normalizes smart quotes in binding expressions to standard quotes
    /// </summary>
    private static string NormalizeQuotesInExpressions(string text)
    {
        if (string.IsNullOrEmpty(text))
            return text;

        // Check if we have potential expressions
        if (!text.Contains("${"))
            return text;

        // Use a regex to find all expressions and normalize quotes only within them
        return Regex.Replace(text, @"\$\{([^}]*)\}", match =>
        {
            string expressionContent = match.Groups[1].Value;

            // Replace smart quotes with straight quotes within expression content
            expressionContent = expressionContent.Replace("\u201C", "\"")  // Left double quote
                                               .Replace("\u201D", "\"")  // Right double quote
                                               .Replace("\u2018", "'")   // Left single quote
                                               .Replace("\u2019", "'");  // Right single quote

            return "${" + expressionContent + "}";
        });
    }
}
