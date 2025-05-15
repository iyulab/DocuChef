using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

namespace DocuChef.Tests;

/// <summary>
/// Helper class for creating PowerPoint test templates
/// </summary>
internal static class PPTHelper
{
    /// <summary>
    /// Creates a basic PowerPoint document with a single slide
    /// </summary>
    public static PresentationDocument CreateBasicPresentation(string path)
    {
        var presentationDocument = PresentationDocument.Create(path, DocumentFormat.OpenXml.PresentationDocumentType.Presentation);
        var presentationPart = presentationDocument.AddPresentationPart();
        presentationPart.Presentation = new Presentation();

        var slideMasterPart = presentationPart.AddNewPart<SlideMasterPart>();
        slideMasterPart.SlideMaster = new SlideMaster();

        var slideLayoutPart = slideMasterPart.AddNewPart<SlideLayoutPart>();
        slideLayoutPart.SlideLayout = new SlideLayout();

        presentationPart.Presentation.SlideIdList = new SlideIdList();

        return presentationDocument;
    }

    /// <summary>
    /// Adds a slide to the presentation
    /// </summary>
    public static SlidePart AddSlide(PresentationDocument presentationDocument, uint slideIndex = 1)
    {
        var presentationPart = presentationDocument.PresentationPart;
        var slidePart = presentationPart.AddNewPart<SlidePart>();
        slidePart.Slide = new Slide();

        slidePart.Slide.CommonSlideData = new CommonSlideData();
        slidePart.Slide.CommonSlideData.ShapeTree = new ShapeTree();

        var slideId = new SlideId();
        slideId.Id = 255U + slideIndex;
        slideId.RelationshipId = presentationPart.GetIdOfPart(slidePart);
        presentationPart.Presentation.SlideIdList.Append(slideId);

        return slidePart;
    }

    /// <summary>
    /// Creates a text shape and adds it to the slide
    /// </summary>
    public static Shape AddTextShape(SlidePart slidePart, string text, string shapeName, uint shapeId, int x, int y, int width, int height)
    {
        var shape = CreateTextShape(text, shapeName, shapeName, shapeId, x, y, width, height);
        slidePart.Slide.CommonSlideData.ShapeTree.Append(shape);
        return shape;
    }

    /// <summary>
    /// Creates a text shape
    /// </summary>
    public static Shape CreateTextShape(string text, string altText, string shapeName, uint shapeId, int x, int y, int width, int height)
    {
        var shape = new Shape();

        // NonVisualDrawingProperties can include the description in the Name property
        var nvdp = new NonVisualDrawingProperties()
        {
            Id = shapeId,
            Name = shapeName,
            Title = altText // Using Title as alternative for Description
        };

        var nvProps = new NonVisualShapeProperties(
            nvdp,
            new NonVisualShapeDrawingProperties(new A.ShapeLocks() { NoGrouping = true }),
            new ApplicationNonVisualDrawingProperties()
        );

        var shapeProps = new ShapeProperties();
        var transform = new A.Transform2D();
        transform.Offset = new A.Offset() { X = x, Y = y };
        transform.Extents = new A.Extents() { Cx = width, Cy = height };
        shapeProps.Transform2D = transform;

        var textBody = new TextBody();
        var paragraph = new A.Paragraph();
        var run = new A.Run();
        var textElement = new A.Text(text);

        run.Append(textElement);
        paragraph.Append(run);
        textBody.Append(paragraph);

        shape.Append(nvProps);
        shape.Append(shapeProps);
        shape.Append(textBody);

        return shape;
    }

    /// <summary>
    /// Adds a directive to the slide notes
    /// </summary>
    public static void AddDirectiveToNotes(SlidePart slidePart, string directiveText)
    {
        var notesSlidePart = slidePart.AddNewPart<NotesSlidePart>();
        notesSlidePart.NotesSlide = new NotesSlide();

        var notesCommonSlideData = new CommonSlideData();
        var notesShapeTree = new ShapeTree();

        var notesShape = new Shape();
        var notesTextBody = new TextBody();
        var notesParagraph = new A.Paragraph();
        var notesRun = new A.Run();
        var notesText = new A.Text(directiveText);

        notesRun.Append(notesText);
        notesParagraph.Append(notesRun);
        notesTextBody.Append(notesParagraph);
        notesShape.Append(notesTextBody);

        notesShapeTree.Append(notesShape);
        notesCommonSlideData.Append(notesShapeTree);
        notesSlidePart.NotesSlide.Append(notesCommonSlideData);
    }

    /// <summary>
    /// Adds a notes slide to a slide with the specified text
    /// </summary>
    public static NotesSlidePart AddNotesSlide(SlidePart slidePart, string notesText)
    {
        // Check if notes slide already exists
        NotesSlidePart notesSlidePart = slidePart.NotesSlidePart;
        if (notesSlidePart == null)
        {
            // Create a new notes slide part
            notesSlidePart = slidePart.AddNewPart<NotesSlidePart>();
            notesSlidePart.NotesSlide = new NotesSlide(
                new CommonSlideData(
                    new ShapeTree(
                        new NonVisualGroupShapeProperties(
                            new NonVisualDrawingProperties() { Id = 1U, Name = "" },
                            new NonVisualGroupShapeDrawingProperties(),
                            new ApplicationNonVisualDrawingProperties()),
                        new GroupShapeProperties(new A.TransformGroup()),
                        new Shape(
                            new NonVisualShapeProperties(
                                new NonVisualDrawingProperties() { Id = 2U, Name = "Notes Placeholder 1" },
                                new NonVisualShapeDrawingProperties(new A.ShapeLocks() { NoGrouping = true }),
                                new ApplicationNonVisualDrawingProperties(new PlaceholderShape())),
                            new ShapeProperties(),
                            new TextBody(
                                new A.BodyProperties(),
                                new A.ListStyle(),
                                new A.Paragraph(
                                    new A.Run(
                                        new A.RunProperties(),
                                        new A.Text() { Text = notesText }
                                    )
                                )
                            )
                        )
                    )
                )
            );
        }
        else
        {
            // Update existing notes with the new text
            A.Text textElement = notesSlidePart.NotesSlide.Descendants<A.Text>().FirstOrDefault();
            if (textElement != null)
            {
                textElement.Text = notesText;
            }
            else
            {
                // If no text element exists, create one
                var shape = notesSlidePart.NotesSlide.Descendants<Shape>().FirstOrDefault();
                if (shape != null)
                {
                    var textBody = shape.GetFirstChild<TextBody>();
                    if (textBody == null)
                    {
                        textBody = new TextBody();
                        shape.Append(textBody);
                    }

                    var paragraph = new A.Paragraph();
                    var run = new A.Run();
                    var text = new A.Text() { Text = notesText };

                    run.Append(text);
                    paragraph.Append(run);
                    textBody.Append(paragraph);
                }
            }
        }

        return notesSlidePart;
    }

    /// <summary>
    /// Creates a template for basic syntax testing
    /// </summary>
    public static void CreateBasicSyntaxTemplate(string path)
    {
        using var presentationDocument = CreateBasicPresentation(path);
        var slidePart = AddSlide(presentationDocument);

        // Title shape - basic variable binding
        AddTextShape(slidePart, "${Title}", "TitleShape", 1, 1524000, 1524000, 6096000, 1008000);

        // Price shape - format specifier
        AddTextShape(slidePart, "Price: ${Price:C2}", "PriceShape", 2, 1524000, 2900000, 4000000, 500000);

        // Object property access
        AddTextShape(slidePart, "Product: ${Product.Name}\nSKU: ${Product.Details.SKU}",
            "ProductShape", 3, 1524000, 3600000, 6096000, 800000);

        // Conditional expression
        AddTextShape(slidePart,
            "Status: ${InStock ? \"In Stock\" : \"Out of Stock\"}\n" +
            "Quantity: ${Quantity < 10 ? \"Low Stock\" : \"Well Stocked\"}",
            "ConditionalShape", 4, 1524000, 4500000, 6096000, 800000);

        // Visibility test shape
        AddTextShape(slidePart, "Visibility Test", "TestShape", 5, 1524000, 5500000, 4000000, 800000);

        // Add directive in slide notes
        AddNotesSlide(slidePart, "#if: ShowElement, target: \"TestShape\"");

        presentationDocument.PresentationPart.Presentation.Save();
    }

    /// <summary>
    /// Creates a template for array syntax testing
    /// </summary>
    public static void CreateArraySyntaxTemplate(string path)
    {
        using var presentationDocument = CreateBasicPresentation(path);
        var slidePart = AddSlide(presentationDocument);

        // Title shape
        AddTextShape(slidePart, "Product List", "TitleShape", 1, 1524000, 1024000, 6096000, 800000);

        // Products container
        var containerShape = new Shape();
        var containerNvProps = new NonVisualShapeProperties(
            new NonVisualDrawingProperties() { Id = 2U, Name = "ProductsContainer" },
            new NonVisualShapeDrawingProperties(),
            new ApplicationNonVisualDrawingProperties()
        );

        var containerShapeProps = new ShapeProperties();
        var containerTransform = new A.Transform2D();
        containerTransform.Offset = new A.Offset() { X = 1524000, Y = 2000000 };
        containerTransform.Extents = new A.Extents() { Cx = 6096000, Cy = 4000000 };
        containerShapeProps.Transform2D = containerTransform;

        containerShape.Append(containerNvProps);
        containerShape.Append(containerShapeProps);
        slidePart.Slide.CommonSlideData.ShapeTree.Append(containerShape);

        // Product items (2 per slide)
        AddTextShape(slidePart, "Product 1: ${Products[0].Name} - $${Products[0].Price}",
            "Product1Shape", 3, 1524000, 2200000, 6096000, 800000);

        AddTextShape(slidePart, "Product 2: ${Products[1].Name} - $${Products[1].Price}",
            "Product2Shape", 4, 1524000, 3200000, 6096000, 800000);

        // Add directive in slide notes
        AddNotesSlide(slidePart, "#foreach: Products, max: 2");

        presentationDocument.PresentationPart.Presentation.Save();
    }

    /// <summary>
    /// Creates a template for nested data structure testing
    /// </summary>
    public static void CreateNestedSyntaxTemplate(string path)
    {
        using var presentationDocument = CreateBasicPresentation(path);

        // Slide 1: Department template
        var deptSlidePart = AddSlide(presentationDocument, 1);
        AddTextShape(deptSlidePart, "Department: ${Departments[0].Name}", "DepartmentTitleShape", 1, 1524000, 1524000, 6096000, 1008000);
        // Add directive for foreach Departments
        AddNotesSlide(deptSlidePart, "#foreach: Departments");

        // Slide 2: Team template
        var teamSlidePart = AddSlide(presentationDocument, 2);
        AddTextShape(teamSlidePart, "Department: ${Departments_Name}", "CurrentDeptShape", 1, 1524000, 1524000, 6096000, 800000);
        AddTextShape(teamSlidePart, "Team: ${Departments_Teams[0].Name}", "TeamInfoShape", 2, 1524000, 2500000, 6096000, 800000);
        // Add directive for foreach Departments_Teams
        AddNotesSlide(teamSlidePart, "#foreach: Departments_Teams");

        // Slide 3: Member template
        var memberSlidePart = AddSlide(presentationDocument, 3);
        AddTextShape(memberSlidePart, "Department: ${Departments_Name}\nTeam: ${Departments_Teams_Name}",
            "HeaderShape", 1, 1524000, 1024000, 6096000, 1000000);
        AddTextShape(memberSlidePart,
            "Member 1: ${Departments_Teams_Members[0].Name}, ${Departments_Teams_Members[0].Role}\n" +
            "Member 2: ${Departments_Teams_Members[1].Name}, ${Departments_Teams_Members[1].Role}",
            "MembersShape", 2, 1524000, 2500000, 6096000, 1500000);
        // Add directive for foreach Departments_Teams_Members, max 2 per slide
        AddNotesSlide(memberSlidePart, "#foreach: Departments_Teams_Members, max: 2");

        // Slide 4: Category template
        var categorySlidePart = AddSlide(presentationDocument, 4);
        AddTextShape(categorySlidePart,
            "Category: ${Categories[0].Name}\n" +
            "Description: ${Categories[0].Description}\n" +
            "Product Count: ${Categories[0].Products.length}",
            "CategoryInfoShape", 1, 1524000, 1524000, 6096000, 1500000);
        // Add directive for foreach Categories
        AddNotesSlide(categorySlidePart, "#foreach: Categories");

        // Slide 5: Products template
        var productsSlidePart = AddSlide(presentationDocument, 5);
        AddTextShape(productsSlidePart, "${Categories_Name} Products", "CategoryNameShape", 1, 1524000, 1024000, 6096000, 800000);
        AddTextShape(productsSlidePart,
            "Product 1: ${Categories_Products[0].Name} - $${Categories_Products[0].Price:N0}\n" +
            "Product 2: ${Categories_Products[1].Name} - $${Categories_Products[1].Price:N0}\n" +
            "Product 3: ${Categories_Products[2].Name} - $${Categories_Products[2].Price:N0}",
            "ProductsListShape", 2, 1524000, 2000000, 6096000, 2000000);
        // Add directive for foreach Categories_Products, max 3 per slide
        AddNotesSlide(productsSlidePart, "#foreach: Categories_Products, max: 3");

        presentationDocument.PresentationPart.Presentation.Save();
    }

    /// <summary>
    /// Gets the first slide part from the presentation
    /// </summary>
    public static SlidePart GetFirstSlidePart(PresentationDocument document)
    {
        var slideId = document.PresentationPart.Presentation.SlideIdList.Elements<SlideId>().First();
        return (SlidePart)document.PresentationPart.GetPartById(slideId.RelationshipId);
    }

    /// <summary>
    /// Gets a specific slide part by slide index
    /// </summary>
    public static SlidePart GetSlidePart(PresentationDocument document, uint slideIndex)
    {
        var slideId = document.PresentationPart.Presentation.SlideIdList.Elements<SlideId>()
            .FirstOrDefault(s => s.Id.Value == (255U + slideIndex));

        if (slideId == null)
            return GetFirstSlidePart(document);

        return (SlidePart)document.PresentationPart.GetPartById(slideId.RelationshipId);
    }

    /// <summary>
    /// Gets all text elements from a slide part
    /// </summary>
    public static List<string> GetTextElements(SlidePart slidePart)
    {
        return slidePart.Slide.Descendants<A.Text>()
            .Select(t => t.Text)
            .Where(text => !string.IsNullOrEmpty(text))
            .ToList();
    }

    /// <summary>
    /// Finds a shape by name
    /// </summary>
    public static Shape FindShapeByName(SlidePart slidePart, string name)
    {
        return slidePart.Slide.Descendants<Shape>()
            .FirstOrDefault(s => s.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value == name);
    }
}