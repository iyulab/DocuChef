using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocuChef.Presentation.Utilities;
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

                        // Extract slide notes
                        var notesText = ExtractNotesFromSlide(slidePart);
                        if (!string.IsNullOrWhiteSpace(notesText))
                        {
                            Console.WriteLine($"   📄 Notes: \"{notesText.Trim()}\"");

                            // Check for directives in notes
                            if (notesText.Contains("#"))
                            {
                                Console.WriteLine($"       🔧 Contains directive!");
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

        private static string ExtractNotesFromSlide(SlidePart slidePart)
        {
            var notesText = new StringBuilder();

            try
            {
                var notesSlidePart = slidePart.NotesSlidePart;
                if (notesSlidePart?.NotesSlide?.CommonSlideData?.ShapeTree != null)
                {
                    var textElements = notesSlidePart.NotesSlide.CommonSlideData.ShapeTree
                        .Descendants<DocumentFormat.OpenXml.Drawing.Text>();
                    foreach (var textElement in textElements)
                    {
                        if (!string.IsNullOrEmpty(textElement.Text))
                        {
                            notesText.AppendLine(textElement.Text);
                        }
                    }
                }
            }
            catch (Exception)
            {
                // Notes may not exist for all slides
                return string.Empty;
            }

            return notesText.ToString();
        }

        /// <summary>
        /// Inspects a generated PowerPoint file to see actual content transformations
        /// </summary>
        public static void InspectGeneratedFile(string filePath)
        {
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"❌ File not found: {filePath}");
                return;
            }

            Console.WriteLine($"\nGenerated File Inspection");
            Console.WriteLine($"==========================");
            Console.WriteLine($"File: {filePath}");

            try
            {
                using var presentation = DocumentFormat.OpenXml.Packaging.PresentationDocument.Open(filePath, false);
                var presentationPart = presentation.PresentationPart;
                if (presentationPart?.Presentation?.SlideIdList == null)
                {
                    Console.WriteLine("❌ No slides found in presentation");
                    return;
                }

                var slideIdList = presentationPart.Presentation.SlideIdList.Elements<SlideId>().ToList();
                Console.WriteLine($"✅ Found {slideIdList.Count} slides in generated file");

                for (int i = 0; i < slideIdList.Count; i++)
                {
                    var slideId = slideIdList[i];
                    var relationshipId = slideId.RelationshipId?.Value;

                    if (relationshipId != null)
                    {
                        var slidePart = (SlidePart)presentationPart.GetPartById(relationshipId);
                        Console.WriteLine($"\nSlide {i + 1} (SlideId={slideId.Id}) Generated Content:");
                        Console.WriteLine($"-------------------------------------------");

                        var slideText = SlideTextExtractor.GetText(slidePart.Slide);
                        var lines = slideText.Split('\n', StringSplitOptions.RemoveEmptyEntries);

                        foreach (var line in lines)
                        {
                            var trimmedLine = line.Trim();
                            if (!string.IsNullOrEmpty(trimmedLine))
                            {
                                if (trimmedLine.Contains("${") && trimmedLine.Contains("}"))
                                {
                                    Console.WriteLine($"   📝 \"{trimmedLine}\"");
                                    Console.WriteLine($"       🎯 Contains template expression!");
                                }
                                else
                                {
                                    Console.WriteLine($"   📝 \"{trimmedLine}\"");
                                }
                            }
                        }

                        // Extract notes if any
                        if (slidePart.NotesSlidePart != null)
                        {
                            var notesText = ExtractNotesFromSlide(slidePart);
                            if (!string.IsNullOrEmpty(notesText))
                            {
                                Console.WriteLine($"   📄 Notes: \"{notesText}\"");
                                if (notesText.Contains("#"))
                                {
                                    Console.WriteLine($"       🔧 Contains directive!");
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error inspecting generated file: {ex.Message}");
            }
        }
    }
}
