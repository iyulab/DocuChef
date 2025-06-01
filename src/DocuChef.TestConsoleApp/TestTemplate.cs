using DocuChef.Presentation;
using System.Diagnostics;

namespace DocuChef.TestConsoleApp;

internal class TestTemplate
{
    internal static void Run(string fileName)
    {
        Console.WriteLine("DocuChef PowerPoint Template Test - NEW SYSTEM (Multi-Slide and Data Binding)");
        Console.WriteLine("===============================================================================");

        // File path setup
        string basePath = AppDomain.CurrentDomain.BaseDirectory;
        string templatePath = Path.Combine(basePath, "files", "ppt", fileName);
        string logoPath = Path.Combine(basePath, "files", "logo.png");
        string outputPath = Path.Combine(basePath, "output_multi_slides_new.pptx");

        // Check if template file exists
        if (!File.Exists(templatePath))
        {
            Console.WriteLine($"Template file not found: {templatePath}");
            return;
        }

        // Check if logo file exists
        if (!File.Exists(logoPath))
        {
            Console.WriteLine($"Logo file not found: {logoPath}");
            Console.WriteLine("Continuing, but the logo may not be displayed.");
        }
        Console.WriteLine($"Template file: {templatePath}");
        Console.WriteLine($"Logo file: {logoPath}");

        // First, inspect the template to see what we're working with
        Console.WriteLine("\n" + new string('=', 60));
        TemplateInspector.InspectTemplate(templatePath);
        Console.WriteLine(new string('=', 60) + "\n");

        try
        {            // Create Chef instance with PowerPoint options
            using var chef = new Chef(new RecipeOptions()
            {
                EnableVerboseLogging = true,
                PowerPoint = new PowerPointOptions()
                {
                    EnableVerboseLogging = true
                }
            });

            // Load PowerPoint template using the NEW system
            Console.WriteLine("Loading template with NEW PowerPoint system...");
            var recipe = chef.LoadPowerPointTemplate(templatePath);

            // Add basic variables
            Console.WriteLine("Adding variables...");
            recipe.AddVariable("Title", "DocuChef Test - NEW SYSTEM");
            recipe.AddVariable("Subtitle", "Multi-Slide and Data Binding Test with PowerPoint Functions");
            recipe.AddVariable("Date", DateTime.Now);
            recipe.AddVariable("LogoPath", logoPath);
            recipe.AddVariable("CompanyName", "DocuChef Technology Lab");

            // Create Items array
            var items = new List<Item>();
            for (int i = 1; i <= 7; i++)
            {
                items.Add(new Item
                {
                    Id = i,
                    Name = $"Product {i}",
                    Description = $"Description for Product {i}.",
                    Price = 10000 * i,
                    // ImageUrl = logoPath  // Using logo as placeholder for product images
                    ImageUrl = $"https://placehold.co/60x60?text=Item{i}"
                });
            }

            // Add Items to recipe
            recipe.AddVariable("Items", items);            // Cook the recipe (generate the presentation)
            Console.WriteLine("Cooking recipe (generating presentation with NEW system)..."); var document = recipe.Cook(outputPath);

            Console.WriteLine($"Presentation generated successfully: {outputPath}");            // Verify the data binding worked correctly
            Console.WriteLine("\n" + new string('=', 60));
            VerifyOutput.CheckDataBinding(outputPath);
            Console.WriteLine(new string('=', 60) + "\n");

            Console.WriteLine($"✅ Test completed successfully! Output: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"❌ Error occurred: {ex.Message}");
            Console.WriteLine($"Stack trace: {ex.StackTrace}");

            // Print inner exception details if available
            var innerEx = ex.InnerException;
            while (innerEx != null)
            {
                Console.WriteLine($"Inner exception: {innerEx.Message}");
                innerEx = innerEx.InnerException;
            }
        }

        // Open the generated PowerPoint file
        try
        {
            var psi = new ProcessStartInfo
            {
                FileName = outputPath,
                UseShellExecute = true
            };
            Process.Start(psi);
            Console.WriteLine("Opened the generated PowerPoint file.");
        }
        catch (Exception openEx)
        {
            Console.WriteLine($"Could not open the file automatically: {openEx.Message}");
        }

    }

    // Product item class
    public class Item
    {
        public int Id { get; set; }
        public required string Name { get; set; }
        public required string Description { get; set; }
        public decimal Price { get; set; }
        public required string ImageUrl { get; set; }
    }
}
