///*
//template_2.pptx description
//# First slide:
//- Top center shape 1: ${ppt.Image(LogoPath)}
//- Title 1: ${Title}
//- Subtitle 2: hello ${Subtitle} world
//  - Formatting required: <bold>hello</bold><fontsize:16>${Subtitle}</fontsize><italic>world</italic>
//- TextBox 3: Created By: ${Date:yyyy-MM-dd}
//# Second slide:
//- Top-left rectangle 1: ${ppt.Image(LogoPath)}
//- Top-right rectangle 5: ${CompanyName}
//- List rectangle 1:
//${Items[0].Id}. ${Items[0].Name} - ${Items[0].Description}
//Price: ${Items[0].Price:C0} KRW
//- List rectangle 2:
//${Items[1].Id}. ${Items[1].Name} - ${Items[1].Description}
//Price: ${Items[1].Price:C0} KRW
//- List rectangle 3:
//${Items[2].Id}. ${Items[2].Name} - ${Items[2].Description}
//Price: ${Items[2].Price:C0} KRW
//- List rectangle 4:
//${Items[3].Id}. ${Items[3].Name} - ${Items[3].Description}
//Price: ${Items[3].Price:C0} KRW
//- List rectangle 5:
//${Items[4].Id}. ${Items[4].Name} - ${Items[4].Description}
//Price: ${Items[4].Price:C0} KRW
// */

//using DocuChef;
//using DocuChef.PowerPoint;
//using System.Diagnostics;

//Console.WriteLine("DocuChef PowerPoint Template Test - Multi-Slide and Data Binding");
//Console.WriteLine("=======================================================");

//// File path setup
//string basePath = AppDomain.CurrentDomain.BaseDirectory;
//string templatePath = Path.Combine(basePath, "files", "ppt", "template_3.pptx");
//string logoPath = Path.Combine(basePath, "files", "logo.png");
//string outputPath = Path.Combine(basePath, "output_multi_slides.pptx");

//// Check if template file exists
//if (!File.Exists(templatePath))
//{
//    Console.WriteLine($"Template file not found: {templatePath}");
//    return;
//}

//// Check if logo file exists
//if (!File.Exists(logoPath))
//{
//    Console.WriteLine($"Logo file not found: {logoPath}");
//    Console.WriteLine("Continuing, but the logo may not be displayed.");
//}

//Console.WriteLine($"Template file: {templatePath}");
//Console.WriteLine($"Logo file: {logoPath}");

//try
//{
//    // Create Chef instance
//    using var chef = new Chef(new RecipeOptions()
//    {
//        EnableVerboseLogging = true,
//        PowerPoint = new PowerPointOptions()
//    });

//    // Load PowerPoint template
//    Console.WriteLine("Loading template...");
//    var recipe = chef.LoadPowerPointTemplate(templatePath);

//    // Add basic variables
//    Console.WriteLine("Adding variables...");
//    recipe.AddVariable("Title", "DocuChef Test");
//    recipe.AddVariable("Subtitle", "Multi-Slide and Data Binding Test");
//    recipe.AddVariable("Date", DateTime.Now);
//    recipe.AddVariable("LogoPath", logoPath);
//    recipe.AddVariable("CompanyName", "DocuChef Technology Lab");

//    // Create Items array
//    var items = new List<Item>();
//    for (int i = 1; i <= 13; i++)
//    {
//        items.Add(new Item
//        {
//            Id = i,
//            Name = $"Product {i}",
//            Description = $"Description for Product {i}.",
//            Price = 10000 * i,
//            //ImageUrl = logoPath
//            ImageUrl = $"https://placehold.co/60x60?text=Item{i}"
//        });
//    }

//    // Add Items variable
//    recipe.AddVariable("Items", items);
//    Console.WriteLine($"Added {items.Count} product items.");

//    // Generate document
//    Console.WriteLine("Generating document...");
//    var document = recipe.Generate();

//    // Save document
//    Console.WriteLine($"Saving document: {outputPath}");
//    document.SaveAs(outputPath);
//    Console.WriteLine("Document generation completed!");

//    // Automatically open the generated document
//    Console.WriteLine("Opening the generated document...");
//    Process.Start(new ProcessStartInfo
//    {
//        FileName = outputPath,
//        UseShellExecute = true
//    });
//}
//catch (Exception ex)
//{
//    Console.WriteLine($"Error occurred: {ex.Message}");
//    Console.WriteLine(ex.StackTrace);
//}

//Console.WriteLine("Program completed. Press any key to exit...");
//Console.ReadKey();

//// Product item class
//public class Item
//{
//    public int Id { get; set; }
//    public string Name { get; set; }
//    public string Description { get; set; }
//    public decimal Price { get; set; }
//    public string ImageUrl { get; set; }
//}