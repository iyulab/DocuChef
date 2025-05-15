using DocuChef.PowerPoint;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using FluentAssertions;
using System.Text;
using Xunit.Abstractions;

namespace DocuChef.Tests
{
    /// <summary>
    /// Tests for PowerPoint group slide functionality
    /// </summary>
    public class PPTGroupTests : TestBase
    {
        private readonly string _tempDirectory;

        public PPTGroupTests(ITestOutputHelper output) : base(output)
        {
            _tempDirectory = Path.Combine(Path.GetTempPath(), "DocuChefTests", Guid.NewGuid().ToString());
            Directory.CreateDirectory(_tempDirectory);
        }

        public override void Dispose()
        {
            try { if (Directory.Exists(_tempDirectory)) Directory.Delete(_tempDirectory, true); }
            catch { }
            base.Dispose();
        }

        [Fact]
        public void Basic_Group_Slide_Creates_Multiple_Slides_With_Proper_Binding()
        {
            // Arrange
            var chef = CreateNewChef();
            var templatePath = Path.Combine(_tempDirectory, "basic_group_template.pptx");

            // Create a simple template with a Groups[0].Name reference
            using (var presentationDoc = PPTHelper.CreateBasicPresentation(templatePath))
            {
                var slidePart = PPTHelper.AddSlide(presentationDoc);

                // Add a shape with a group array reference
                PPTHelper.AddTextShape(slidePart,
                    "Group: ${Groups[0].Name}",
                    "GroupNameShape", 1, 1524000, 1524000, 6096000, 800000);

                // Add directive to the slide notes
                PPTHelper.AddNotesSlide(slidePart, "#foreach: Groups");

                presentationDoc.Save();
            }

            // Create test data with 2 groups
            var groups = new List<Group>
            {
                new Group { Name = "Group A" },
                new Group { Name = "Group B" }
            };

            var recipe = chef.LoadPowerPointTemplate(templatePath);
            recipe.AddVariable("Groups", groups);

            // Act
            var document = recipe.Generate();
            var outputPath = Path.Combine(_tempDirectory, "basic_group_output.pptx");
            document.SaveAs(outputPath);

            // Assert
            using var resultDocument = PresentationDocument.Open(outputPath, false);

            // Verify we have 2 slides
            var slideIds = resultDocument.PresentationPart.Presentation.SlideIdList.Elements<SlideId>().ToList();
            slideIds.Count.Should().Be(2, "Each group should generate its own slide");

            // Get text from first slide
            var firstSlidePart = (SlidePart)resultDocument.PresentationPart.GetPartById(slideIds[0].RelationshipId);
            var firstSlideText = string.Join(" ", PPTHelper.GetTextElements(firstSlidePart));
            _output.WriteLine($"First slide text: {firstSlideText}");
            firstSlideText.Should().Contain("Group A", "First slide should show first group name");

            // Get text from second slide
            var secondSlidePart = (SlidePart)resultDocument.PresentationPart.GetPartById(slideIds[1].RelationshipId);
            var secondSlideText = string.Join(" ", PPTHelper.GetTextElements(secondSlidePart));
            _output.WriteLine($"Second slide text: {secondSlideText}");
            secondSlideText.Should().Contain("Group B", "Second slide should show second group name");
        }

        [Fact]
        public void Nested_Group_Items_Creates_Multiple_Slides_With_Proper_Binding()
        {
            // Arrange
            var chef = CreateNewChef();

            // Create a template with two slides - a group slide and an items slide
            var templatePath = Path.Combine(_tempDirectory, "nested_group_template.pptx");
            using (var presentationDoc = PPTHelper.CreateBasicPresentation(templatePath))
            {
                // First slide template - Group overview
                var groupSlidePart = PPTHelper.AddSlide(presentationDoc, 1);
                PPTHelper.AddTextShape(groupSlidePart,
                    "Group: ${Groups[0].Name}",
                    "GroupNameShape", 1, 1524000, 1524000, 6096000, 800000);

                // Add directive to group slide notes
                PPTHelper.AddNotesSlide(groupSlidePart, "#foreach: Groups");

                // Second slide template - Items for a group
                // Important: Use direct indexing for the base level first
                var itemsSlidePart = PPTHelper.AddSlide(presentationDoc, 2);
                PPTHelper.AddTextShape(itemsSlidePart,
                    "Group: ${Groups[0].Name}\nItems for this group:",
                    "GroupTitleShape", 1, 1524000, 1024000, 6096000, 800000);

                // Use direct array indexing for items
                PPTHelper.AddTextShape(itemsSlidePart,
                    "Item 1: ${Groups[0].Items[0].Name}\nItem 2: ${Groups[0].Items[1].Name}",
                    "ItemsListShape", 2, 1524000, 2000000, 6096000, 800000);

                // Add directive to items slide notes - max 2 items per slide
                PPTHelper.AddNotesSlide(itemsSlidePart, "#foreach-items: Groups.Items, max: 2");

                presentationDoc.Save();
            }

            // Create test data with 2 groups, each with 2 items
            var groups = new List<GroupWithItems>
            {
                new GroupWithItems
                {
                    Name = "Group A",
                    Items = new List<Item>
                    {
                        new Item { Name = "A-Item 1" },
                        new Item { Name = "A-Item 2" }
                    }
                },
                new GroupWithItems
                {
                    Name = "Group B",
                    Items = new List<Item>
                    {
                        new Item { Name = "B-Item 1" },
                        new Item { Name = "B-Item 2" }
                    }
                }
            };

            var recipe = chef.LoadPowerPointTemplate(templatePath);
            recipe.AddVariable("Groups", groups);

            // Act
            var document = recipe.Generate();
            var outputPath = Path.Combine(_tempDirectory, "nested_group_output.pptx");
            document.SaveAs(outputPath);

            // Assert
            using var resultDocument = PresentationDocument.Open(outputPath, false);

            // Get all slide texts for analysis
            var slideIds = resultDocument.PresentationPart.Presentation.SlideIdList.Elements<SlideId>().ToList();
            foreach (var slideId in slideIds)
            {
                var slidePart = (SlidePart)resultDocument.PresentationPart.GetPartById(slideId.RelationshipId);
                var slideText = string.Join(" ", PPTHelper.GetTextElements(slidePart));
                _output.WriteLine($"Slide text: '{slideText}'");
            }

            // We should have 4 slides (2 groups + 2 item slides, one per group)
            slideIds.Count.Should().Be(4, "Should have 2 group slides and 2 item slides");

            // Test both approaches for verification
            // 1. Check all text content at once
            var allSlideTexts = new StringBuilder();
            foreach (var slideId in slideIds)
            {
                var slidePart = (SlidePart)resultDocument.PresentationPart.GetPartById(slideId.RelationshipId);
                allSlideTexts.Append(string.Join(" ", PPTHelper.GetTextElements(slidePart)));
            }

            string allText = allSlideTexts.ToString();

            // Verify at least one group name appears
            allText.Should().Contain("Group A", "Group A should appear in the slides");

            // Verify at least one item name appears
            allText.Should().Contain("A-Item", "At least one A-Item should appear in the slides");
        }

        [Fact]
        public void Exceeding_Design_Capacity_Creates_Correct_Number_Of_Slides()
        {
            // Arrange
            var chef = CreateNewChef();
            var templatePath = Path.Combine(_tempDirectory, "capacity_test_template.pptx");

            // Create a template with slides that can only show 2 items per slide
            using (var presentationDoc = PPTHelper.CreateBasicPresentation(templatePath))
            {
                var slidePart = PPTHelper.AddSlide(presentationDoc);

                // Design slide to show only 2 items
                PPTHelper.AddTextShape(slidePart,
                    "Items per slide: 2",
                    "HeaderShape", 1, 1524000, 1024000, 6096000, 800000);

                PPTHelper.AddTextShape(slidePart,
                    "Item 1: ${Items[0].Name}",
                    "Item1Shape", 2, 1524000, 2000000, 6096000, 800000);

                PPTHelper.AddTextShape(slidePart,
                    "Item 2: ${Items[1].Name}",
                    "Item2Shape", 3, 1524000, 3000000, 6096000, 800000);

                // Add directive to slide notes - max 2 items per slide
                PPTHelper.AddNotesSlide(slidePart, "#foreach: Items, max: 2");

                presentationDoc.Save();
            }

            // Create test data with 5 items (which exceeds the 2 items per slide design)
            var items = new List<Item>
            {
                new Item { Name = "Item 1" },
                new Item { Name = "Item 2" },
                new Item { Name = "Item 3" },
                new Item { Name = "Item 4" },
                new Item { Name = "Item 5" }
            };

            var recipe = chef.LoadPowerPointTemplate(templatePath);
            recipe.AddVariable("Items", items);

            // Act
            var document = recipe.Generate();
            var outputPath = Path.Combine(_tempDirectory, "capacity_test_output.pptx");
            document.SaveAs(outputPath);

            // Assert
            using var resultDocument = PresentationDocument.Open(outputPath, false);

            // Get all slides
            var slideIds = resultDocument.PresentationPart.Presentation.SlideIdList.Elements<SlideId>().ToList();

            // Log all slide texts
            for (int i = 0; i < slideIds.Count; i++)
            {
                var slidePart = (SlidePart)resultDocument.PresentationPart.GetPartById(slideIds[i].RelationshipId);
                var slideText = string.Join(" ", PPTHelper.GetTextElements(slidePart));
                _output.WriteLine($"Slide {i + 1} text: '{slideText}'");
            }

            // Verify we have 3 slides (ceiling(5/2) = 3)
            slideIds.Count.Should().Be(3, "We need 3 slides to display 5 items with 2 items per slide");

            // Verify content of each slide
            // First slide should have items 1 and 2
            var slide1Part = (SlidePart)resultDocument.PresentationPart.GetPartById(slideIds[0].RelationshipId);
            var slide1Text = string.Join(" ", PPTHelper.GetTextElements(slide1Part));
            slide1Text.Should().Contain("Item 1");
            slide1Text.Should().Contain("Item 2");

            // Second slide should have items 3 and 4
            var slide2Part = (SlidePart)resultDocument.PresentationPart.GetPartById(slideIds[1].RelationshipId);
            var slide2Text = string.Join(" ", PPTHelper.GetTextElements(slide2Part));
            slide2Text.Should().Contain("Item 3");
            slide2Text.Should().Contain("Item 4");

            // Third slide should have item 5 only (and possibly empty item 6)
            var slide3Part = (SlidePart)resultDocument.PresentationPart.GetPartById(slideIds[2].RelationshipId);
            var slide3Text = string.Join(" ", PPTHelper.GetTextElements(slide3Part));
            slide3Text.Should().Contain("Item 5");
        }

        // Helper method for PPTHelper class to add notes to slides
        private static class PPTHelperAdditions
        {
            public static void AddNotesSlide(SlidePart slidePart, string notesText)
            {
                // Check if notes slide already exists
                if (slidePart.NotesSlidePart == null)
                {
                    // Create a new notes slide part
                    NotesSlidePart notesSlidePart = slidePart.AddNewPart<NotesSlidePart>();

                    // Generate the notes slide with the specified text
                    GenerateNotesSlidePart(notesSlidePart, notesText);

                    // Create a relationship between the slide and the notes slide
                    slidePart.AddPart(notesSlidePart);
                }
                else
                {
                    // Just update existing notes with the new text
                    UpdateNotesSlidePart(slidePart.NotesSlidePart, notesText);
                }
            }

            private static void GenerateNotesSlidePart(NotesSlidePart notesSlidePart, string notesText)
            {
                var notesSlide = new NotesSlide(
                    new CommonSlideData(
                        new ShapeTree(
                            new NonVisualGroupShapeProperties(
                                new NonVisualDrawingProperties() { Id = 1U, Name = "" },
                                new NonVisualGroupShapeDrawingProperties(),
                                new ApplicationNonVisualDrawingProperties()),
                            new GroupShapeProperties(new DocumentFormat.OpenXml.Drawing.TransformGroup()),
                            new Shape(
                                new NonVisualShapeProperties(
                                    new NonVisualDrawingProperties() { Id = 2U, Name = "Notes Placeholder 1" },
                                    new NonVisualShapeDrawingProperties(new DocumentFormat.OpenXml.Drawing.ShapeLocks() { NoGrouping = true }),
                                    new ApplicationNonVisualDrawingProperties(new PlaceholderShape())),
                                new ShapeProperties(),
                                new DocumentFormat.OpenXml.Presentation.TextBody(
                                    new DocumentFormat.OpenXml.Drawing.BodyProperties(),
                                    new DocumentFormat.OpenXml.Drawing.ListStyle(),
                                    new DocumentFormat.OpenXml.Drawing.Paragraph(
                                        new DocumentFormat.OpenXml.Drawing.Run(
                                            new DocumentFormat.OpenXml.Drawing.RunProperties(),
                                            new DocumentFormat.OpenXml.Drawing.Text() { Text = notesText }
                                        )
                                    )
                                )
                            )
                        )
                    )
                );

                notesSlidePart.NotesSlide = notesSlide;
            }

            private static void UpdateNotesSlidePart(NotesSlidePart notesSlidePart, string notesText)
            {
                // Find the text body
                var textBody = notesSlidePart.NotesSlide.Descendants<DocumentFormat.OpenXml.Presentation.TextBody>().FirstOrDefault();
                if (textBody == null)
                {
                    // Add a new text body if not found
                    var shape = notesSlidePart.NotesSlide.Descendants<Shape>().FirstOrDefault();
                    if (shape != null)
                    {
                        textBody = new DocumentFormat.OpenXml.Presentation.TextBody(
                            new DocumentFormat.OpenXml.Drawing.BodyProperties(),
                            new DocumentFormat.OpenXml.Drawing.ListStyle());
                        shape.Append(textBody);
                    }
                }

                if (textBody != null)
                {
                    // Clear existing paragraphs
                    textBody.RemoveAllChildren<DocumentFormat.OpenXml.Drawing.Paragraph>();

                    // Add new paragraph with the notes text
                    textBody.Append(
                        new DocumentFormat.OpenXml.Drawing.Paragraph(
                            new DocumentFormat.OpenXml.Drawing.Run(
                                new DocumentFormat.OpenXml.Drawing.RunProperties(),
                                new DocumentFormat.OpenXml.Drawing.Text() { Text = notesText }
                            )
                        )
                    );
                }
            }
        }

        // Test data classes
        private class Group
        {
            public string Name { get; set; }
        }

        private class GroupWithItems
        {
            public string Name { get; set; }
            public List<Item> Items { get; set; } = new List<Item>();
        }

        private class Item
        {
            public string Name { get; set; }
        }
    }
}