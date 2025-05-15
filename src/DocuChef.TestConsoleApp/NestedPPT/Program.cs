using DocuChef.PowerPoint;
using System.Diagnostics;

/*
# Template Structure: nested_template.pptx

## Template Slide 1: Title Slide
- Top Center: ${ppt.Image(LogoPath)} - Company logo image
- Center: ${PresentationTitle} - Main presentation title
- Below Title: ${PresentationSubtitle} - Presentation subtitle
- Below Subtitle: Created At: ${Date:yyyy년 MM월 dd일} - Formatted date
- Below Date: ${CompanyName} Company - Company name

## Template Slide 2: Category Detail Slide
- Top: ${Categories[0].Name} - Category name (will be repeated for each category)
- Right: ${ppt.Image(Categories[0].ImageUrl)} - Category image
- Center: ${Categories[0].Description} - Category description
- Bottom Right: ${ppt.Image(LogoPath)} - Company logo
- Slide Notes: #foreach: Categories - This directive is used for automatic slide duplication (Should work without this option.)

## Template Slide 3: Products List Slide
- Top: ${Categories_Name} - Current category name
- Item 1:
  * Left: ${ppt.Image(Categories_Products[0].ImageUrl)} - Product image
  * Center: ${Categories_Products[0].Name}${Categories_Products[0].Description} - Product name and description
  * Right: ${Categories_Products[0].Price:C2} - Formatted product price
- Item 2: 
  * Left: ${ppt.Image(Categories_Products[1].ImageUrl)} - Product image
  * Center: ${Categories_Products[1].Name}${Categories_Products[1].Description} - Product name and description
  * Right: ${Categories_Products[1].Price:C2} - Formatted product price
- Item 3:
  * Left: ${ppt.Image(Categories_Products[2].ImageUrl)} - Product image
  * Center: ${Categories_Products[2].Name}${Categories_Products[2].Description} - Product name and description
  * Right: ${Categories_Products[2].Price:C2} - Formatted product price
- Slide Notes: 
  #foreach: Categories_Products, max: 3  (Should work without this option.)

# Expected Results

Slide 1: Title Slide
- Top Center: Company logo (from ${ppt.Image(LogoPath)})
- Center: "Product Catalog 2025" (from ${PresentationTitle})
- Below Title: "Cutting-Edge Technology for Modern Life" (from ${PresentationSubtitle})
- Below Subtitle: "Created At: 2025년 05월 15일" (from ${Date:yyyy년 MM월 dd일})
- Below Date: "Global Electronics Inc. Company" (from ${CompanyName} Company)

Slide 2: Category Detail Slide (Smartphones)
- Top: "Smartphones" (from ${Categories[0].Name})
- Right: Smartphone image (from ${ppt.Image(Categories[0].ImageUrl)})
- Center: "Latest mobile devices with cutting-edge technology" (from ${Categories[0].Description})
- Bottom Right: Company logo (from ${ppt.Image(LogoPath)})

Slide 3: Products List Slide (Smartphones Products 1-3)
- Top: "Smartphones" (from ${Categories_Name})
- Item 1 Left: S25 image (from ${ppt.Image(Categories_Products[0].ImageUrl)})
- Item 1 Center: "Ultra Galaxy S25" + "Flagship smartphone with 8K video and AI assistant" (from ${Categories_Products[0].Name}${Categories_Products[0].Description})
- Item 1 Right: "$1,299.99" (from ${Categories_Products[0].Price:C2})
- Item 2 Left: iPhone16 image (from ${ppt.Image(Categories_Products[1].ImageUrl)})
- Item 2 Center: "iPhone 16 Pro" + "Premium smartphone with advanced camera system" (from ${Categories_Products[1].Name}${Categories_Products[1].Description})
- Item 2 Right: "$1,399.99" (from ${Categories_Products[1].Price:C2})
- Item 3 Left: Pixel9 image (from ${ppt.Image(Categories_Products[2].ImageUrl)})
- Item 3 Center: "Pixel 9" + "Pure Android experience with exceptional photography" (from ${Categories_Products[2].Name}${Categories_Products[2].Description})
- Item 3 Right: "$999.99" (from ${Categories_Products[2].Price:C2})

Slide 4: Products List Slide (Smartphones Product 4)
- Top: "Smartphones" (from ${Categories_Name})
- Item 1 Left: FlipZ5 image (from ${ppt.Image(Categories_Products[0].ImageUrl)}) - note index has auto-adjusted to Categories_Products[3]
- Item 1 Center: "Flip Z5" + "Foldable smartphone with flexible display" (from ${Categories_Products[0].Name}${Categories_Products[0].Description})
- Item 1 Right: "$1,099.99" (from ${Categories_Products[0].Price:C2})
- Item 2: Hidden (not enough products in this category)
- Item 3: Hidden (not enough products in this category)

Slide 5: Category Detail Slide (Laptops)
- Top: "Laptops" (from ${Categories[0].Name}) - note index has auto-adjusted to Categories[1]
- Right: Laptops image (from ${ppt.Image(Categories[0].ImageUrl)})
- Center: "Powerful computing devices for work and play" (from ${Categories[0].Description})
- Bottom Right: Company logo (from ${ppt.Image(LogoPath)})

Slide 6: Products List Slide (Laptops Products 1-3)
- Top: "Laptops" (from ${Categories_Name})
- All 3 product items visible showing the 3 laptop products
- Item 1: UltraBook Pro - $1,799.99
- Item 2: MacBook Air M4 - $1,499.99
- Item 3: Gaming Titan X - $2,499.99

Slide 7: Category Detail Slide (Smart Home)
- Top: "Smart Home" (from ${Categories[0].Name}) - index auto-adjusted to Categories[2]
- Right: Smart Home image
- Center: "Connected devices for the modern home"
- Bottom Right: Company logo

Slide 8: Products List Slide (Smart Home Products 1-3)
- Top: "Smart Home"
- All 3 product items showing the first 3 smart home products
- Item 1: Smart Hub Pro - $249.99
- Item 2: AI Security Camera - $199.99
- Item 3: Smart Thermostat - $129.99

Slide 9: Products List Slide (Smart Home Products 4-5)
- Top: "Smart Home"
- Item 1: Voice Assistant Speaker - $89.99
- Item 2: Smart Lighting Kit - $149.99
- Item 3: Hidden (not enough products in this category)

Slide 10: Category Detail Slide (Wearables)
- Top: "Wearables" (from ${Categories[0].Name}) - index auto-adjusted to Categories[3]
- Right: Wearables image
- Center: "Wearable technology for fitness and productivity"
- Bottom Right: Company logo

Slide 11: Products List Slide (Wearables Products 1-3)
- Top: "Wearables"
- All 3 product items showing all wearable products
- Item 1: Fitness Watch Pro - $299.99
- Item 2: Smart Glasses - $499.99
- Item 3: Health Monitor Band - $179.99
*/

namespace DocuChef.TestConsoleApp.NestedPPT
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("DocuChef PowerPoint Template Test - Nested Data Structure");
            Console.WriteLine("=======================================================");

            // File path setup
            string basePath = AppDomain.CurrentDomain.BaseDirectory;
            string templatePath = Path.Combine(basePath, "files", "ppt", "nested_template.pptx");
            string logoPath = Path.Combine(basePath, "files", "logo.png");
            string outputPath = Path.Combine(basePath, "output_nested_test.pptx");

            // Check if template file exists
            if (!File.Exists(templatePath))
            {
                Console.WriteLine($"Template file not found: {templatePath}");
                Console.WriteLine($"Expected path: {templatePath}");
                Console.WriteLine("Please ensure 'nested_template.pptx' exists in the 'files/ppt' directory.");
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

            try
            {
                // Create Chef instance with verbose logging
                using var chef = new Chef(new RecipeOptions()
                {
                    EnableVerboseLogging = true,
                    PowerPoint = new PowerPointOptions()
                    {
                        // Enable automatic slide creation for array items
                        CreateNewSlidesWhenNeeded = true,
                        MaxSlidesFromTemplate = 50 // Allow up to 50 slides to be generated
                    }
                });

                // Load PowerPoint template
                Console.WriteLine("Loading template...");
                var recipe = chef.LoadPowerPointTemplate(templatePath);

                // Add basic variables
                Console.WriteLine("Adding main variables...");
                recipe.AddVariable("CompanyName", "Global Electronics Inc.");
                recipe.AddVariable("PresentationTitle", "Product Catalog 2025");
                recipe.AddVariable("PresentationSubtitle", "Cutting-Edge Technology for Modern Life");
                recipe.AddVariable("Date", DateTime.Now);
                recipe.AddVariable("LogoPath", logoPath);

                // Create nested data structure - Categories with Products
                var categories = CreateSampleCategoryData();
                recipe.AddVariable("Categories", categories);
                Console.WriteLine($"Added {categories.Count} categories with their products.");

                // Generate document
                Console.WriteLine("Generating document...");
                var document = recipe.Generate();

                // Save document
                Console.WriteLine($"Saving document: {outputPath}");
                document.SaveAs(outputPath);
                Console.WriteLine("Document generation completed!");

                // Automatically open the generated document
                Console.WriteLine("Opening the generated document...");
                Process.Start(new ProcessStartInfo
                {
                    FileName = outputPath,
                    UseShellExecute = true
                });
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error occurred: {ex.Message}");
                Console.WriteLine(ex.StackTrace);
            }

            Console.WriteLine("Program completed. Press any key to exit...");
            Console.ReadKey();
        }

        /// <summary>
        /// Creates sample data with nested structure (Categories -> Products)
        /// </summary>
        static List<Category> CreateSampleCategoryData()
        {
            return new List<Category>
            {
                new Category
                {
                    Id = 1,
                    Name = "Smartphones",
                    Description = "Latest mobile devices with cutting-edge technology",
                    ImageUrl = "https://placehold.co/100x100?text=Smartphones",
                    Products = new List<Product>
                    {
                        new Product { Id = 101, Name = "Ultra Galaxy S25", Price = 1299.99m, Description = "Flagship smartphone with 8K video and AI assistant", ImageUrl = "https://placehold.co/60x60?text=S25" },
                        new Product { Id = 102, Name = "iPhone 16 Pro", Price = 1399.99m, Description = "Premium smartphone with advanced camera system", ImageUrl = "https://placehold.co/60x60?text=iPhone16" },
                        new Product { Id = 103, Name = "Pixel 9", Price = 999.99m, Description = "Pure Android experience with exceptional photography", ImageUrl = "https://placehold.co/60x60?text=Pixel9" },
                        new Product { Id = 104, Name = "Flip Z5", Price = 1099.99m, Description = "Foldable smartphone with flexible display", ImageUrl = "https://placehold.co/60x60?text=FlipZ5" }
                    }
                },
                new Category
                {
                    Id = 2,
                    Name = "Laptops",
                    Description = "Powerful computing devices for work and play",
                    ImageUrl = "https://placehold.co/100x100?text=Laptops",
                    Products = new List<Product>
                    {
                        new Product { Id = 201, Name = "UltraBook Pro", Price = 1799.99m, Description = "Ultra-thin laptop with 24-hour battery life", ImageUrl = "https://placehold.co/60x60?text=UltraBook" },
                        new Product { Id = 202, Name = "MacBook Air M4", Price = 1499.99m, Description = "Lightweight laptop with powerful performance", ImageUrl = "https://placehold.co/60x60?text=MacBookAir" },
                        new Product { Id = 203, Name = "Gaming Titan X", Price = 2499.99m, Description = "High-performance gaming laptop with RTX 5080", ImageUrl = "https://placehold.co/60x60?text=GamingTitan" }
                    }
                },
                new Category
                {
                    Id = 3,
                    Name = "Smart Home",
                    Description = "Connected devices for the modern home",
                    ImageUrl = "https://placehold.co/100x100?text=SmartHome",
                    Products = new List<Product>
                    {
                        new Product { Id = 301, Name = "Smart Hub Pro", Price = 249.99m, Description = "Central control system for all smart home devices", ImageUrl = "https://placehold.co/60x60?text=SmartHub" },
                        new Product { Id = 302, Name = "AI Security Camera", Price = 199.99m, Description = "Intelligent security camera with facial recognition", ImageUrl = "https://placehold.co/60x60?text=AICam" },
                        new Product { Id = 303, Name = "Smart Thermostat", Price = 129.99m, Description = "Energy-saving temperature control with learning capability", ImageUrl = "https://placehold.co/60x60?text=Thermostat" },
                        new Product { Id = 304, Name = "Voice Assistant Speaker", Price = 89.99m, Description = "Premium speaker with integrated voice assistant", ImageUrl = "https://placehold.co/60x60?text=Speaker" },
                        new Product { Id = 305, Name = "Smart Lighting Kit", Price = 149.99m, Description = "Customizable lighting system with app control", ImageUrl = "https://placehold.co/60x60?text=Lighting" }
                    }
                },
                new Category
                {
                    Id = 4,
                    Name = "Wearables",
                    Description = "Wearable technology for fitness and productivity",
                    ImageUrl = "https://placehold.co/100x100?text=Wearables",
                    Products = new List<Product>
                    {
                        new Product { Id = 401, Name = "Fitness Watch Pro", Price = 299.99m, Description = "Advanced fitness tracking with health monitoring", ImageUrl = "https://placehold.co/60x60?text=FitnessWatch" },
                        new Product { Id = 402, Name = "Smart Glasses", Price = 499.99m, Description = "AR-enabled glasses with voice control", ImageUrl = "https://placehold.co/60x60?text=SmartGlasses" },
                        new Product { Id = 403, Name = "Health Monitor Band", Price = 179.99m, Description = "24/7 health monitoring wristband with ECG", ImageUrl = "https://placehold.co/60x60?text=HealthBand" }
                    }
                }
            };
        }
    }

    // Data models for nested structure
    public class Category
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public string Description { get; set; }
        public string ImageUrl { get; set; }
        public List<Product> Products { get; set; } = new List<Product>();
    }

    public class Product
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public string Description { get; set; }
        public decimal Price { get; set; }
        public string ImageUrl { get; set; }
    }
}