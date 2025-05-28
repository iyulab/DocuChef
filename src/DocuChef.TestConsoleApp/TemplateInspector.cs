using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using System;
using System.IO;
using System.Linq;
using System.Text;

namespace DocuChef.TestConsoleApp
{
    public static class TemplateInspector
    {
        public static void InspectTemplate(string templatePath)
        {
            Console.WriteLine("Template Inspection");
            Console.WriteLine("==================");
            Console.WriteLine($"Template file: {templatePath}");
            
            if (!File.Exists(templatePath))
            {
                Console.WriteLine("❌ Template file not found!");
                return;
            }

            try
            {
                using (var presentationDocument = PresentationDocument.Open(templatePath, false))
                {
                    var presentationPart = presentationDocument.PresentationPart;
                    if (presentationPart?.Presentation?.SlideIdList == null)
                    {
                        Console.WriteLine("❌ No slides found in template");
                        return;
                    }

                    var slideIds = presentationPart.Presentation.SlideIdList.Elements<SlideId>().ToList();
                    Console.WriteLine($"✅ Found {slideIds.Count} slides in template\n");

                    for (int i = 0; i < slideIds.Count; i++)
                    {
                        var slideId = slideIds[i];
                        var slidePart = (SlidePart)presentationPart.GetPartById(slideId.RelationshipId!);
                        
                        Console.WriteLine($"Slide {i + 1} (SlideId={slideId.Id}) Template Content:");
                        Console.WriteLine("-------------------------------------------");
                        
                        // Extract all text from the slide
                        var allText = ExtractAllTextFromSlide(slidePart);
                        if (string.IsNullOrWhiteSpace(allText))
                        {
                            Console.WriteLine("   ❌ No text content found");
                        }
                        else
                        {
                            var lines = allText.Split('\n', StringSplitOptions.RemoveEmptyEntries);
                            foreach (var line in lines)
                            {
                                var trimmed = line.Trim();
                                if (!string.IsNullOrWhiteSpace(trimmed))
                                {
                                    Console.WriteLine($"   📝 \"{trimmed}\"");
                                    
                                    // Check for template expressions
                                    if (trimmed.Contains("${"))
                                    {
                                        Console.WriteLine($"       🎯 Contains template expression!");
                                    }
                                }
                            }
                        }
                        
                        Console.WriteLine();
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error reading template: {ex.Message}");
            }
        }

        private static string ExtractAllTextFromSlide(SlidePart slidePart)
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
    }
}
