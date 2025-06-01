using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using System;
using System.IO;
using System.Linq;
using System.Text;

namespace DocuChef.TestConsoleApp
{
    public static class VerifyOutput
    {
        public static void CheckDataBinding(string filePath)
        {
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found: {filePath}");
                return;
            }

            Console.WriteLine("Verifying Data Binding in Generated PowerPoint");
            Console.WriteLine("==============================================");

            try
            {
                using (var presentationDocument = PresentationDocument.Open(filePath, false))
                {
                    var presentationPart = presentationDocument.PresentationPart;
                    if (presentationPart?.Presentation?.SlideIdList == null)
                    {
                        Console.WriteLine("No slides found in presentation");
                        return;
                    }

                    var slideIds = presentationPart.Presentation.SlideIdList.Elements<SlideId>().ToList();
                    Console.WriteLine($"Found {slideIds.Count} slides in the presentation\n");

                    for (int i = 0; i < slideIds.Count; i++)
                    {
                        var slideId = slideIds[i];
                        var slidePart = (SlidePart)presentationPart.GetPartById(slideId.RelationshipId!);

                        Console.WriteLine($"Slide {i + 1} Content:");
                        Console.WriteLine("-------------------");

                        // Extract all text from the slide
                        var allText = ExtractTextFromSlide(slidePart);
                        Console.WriteLine(allText);

                        // Check for unresolved template expressions
                        var unresolvedExpressions = FindUnresolvedExpressions(allText);
                        if (unresolvedExpressions.Any())
                        {
                            Console.WriteLine("\n🚨 UNRESOLVED EXPRESSIONS FOUND:");
                            foreach (var expr in unresolvedExpressions)
                            {
                                Console.WriteLine($"   - {expr}");
                            }
                        }
                        else
                        {
                            Console.WriteLine("\n✅ No unresolved template expressions found!");
                        }

                        // Check for expected resolved values
                        CheckExpectedValues(allText, i + 1);

                        Console.WriteLine(new string('=', 50));
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error reading PowerPoint file: {ex.Message}");
            }
        }

        private static string ExtractTextFromSlide(SlidePart slidePart)
        {
            var allText = new StringBuilder();

            if (slidePart.Slide != null)
            {
                // Get all text elements
                var textElements = slidePart.Slide.Descendants<DocumentFormat.OpenXml.Drawing.Text>();
                foreach (var textElement in textElements)
                {
                    if (!string.IsNullOrEmpty(textElement.Text))
                    {
                        allText.AppendLine(textElement.Text);
                    }
                }
            }

            return allText.ToString();
        }

        private static string[] FindUnresolvedExpressions(string text)
        {
            var expressions = new List<string>();
            var lines = text.Split('\n', StringSplitOptions.RemoveEmptyEntries);

            foreach (var line in lines)
            {
                var trimmed = line.Trim();
                if (trimmed.StartsWith("${") && trimmed.EndsWith("}"))
                {
                    expressions.Add(trimmed);
                }

                // Also check for expressions within text
                var start = 0;
                while ((start = trimmed.IndexOf("${", start)) >= 0)
                {
                    var end = trimmed.IndexOf("}", start);
                    if (end > start)
                    {
                        var expr = trimmed.Substring(start, end - start + 1);
                        if (!expressions.Contains(expr))
                        {
                            expressions.Add(expr);
                        }
                        start = end + 1;
                    }
                    else
                    {
                        break;
                    }
                }
            }

            return expressions.ToArray();
        }

        private static void CheckExpectedValues(string text, int slideNumber)
        {
            Console.WriteLine("\nChecking for expected resolved values:");

            var expectedValues = new Dictionary<string, string>
            {
            };

            var foundValues = 0;
            foreach (var kvp in expectedValues)
            {
                if (text.Contains(kvp.Key))
                {
                    Console.WriteLine($"   ✅ Found {kvp.Value}: {kvp.Key}");
                    foundValues++;
                }
                else
                {
                    Console.WriteLine($"   ❌ Missing {kvp.Value}: {kvp.Key}");
                }
            }

            Console.WriteLine($"\nData binding success rate: {foundValues}/{expectedValues.Count} values found");
        }
    }
}
