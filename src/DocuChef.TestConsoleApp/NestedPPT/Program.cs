using DocuChef.PowerPoint;
using System.Diagnostics;

/*
# nested_template.pptx 설명
## 슬라이드 1
- 상단 중앙: ${ppt.Image(LogoPath)}
- 중앙: ${PresentationTitle}
- Title 아래: ${PresentationSubtitle}
- Subtitle 아래: Created At: ${Date:yyyy년 MM월 dd일}
- Date 아래: ${CompanyName} Company

## 슬라이드 2
- 상단: ${Categories[0].Name}
- 우측: ${ppt.Image(Categories[0].ImageUrl)}
- 중앙: ${Categories[0].Description}
- 우하단: ${ppt.Image(LogoPath)}

## 슬라이드 3
- 상단: ${Categories_Name}
- 목록1 좌측: ${ppt.Image(Categories_Products[0].ImageUrl)}
- 목록1 중앙: ${Categories_Products[0].Name}${Categories_Products[0].Description}
- 목록1 우측: ${Categories_Products[0].Price:C2}
- 목록2 좌측: ${ppt.Image(Categories_Products[1].ImageUrl)}
- 목록2 중앙: ${Categories_Products[1].Name}${Categories_Products[1].Description}
- 목록2 우측: ${Categories_Products[1].Price:C2}
- 목록3 좌측: ${ppt.Image(Categories_Products[2].ImageUrl)}
- 목록3 중앙: ${Categories_Products[2].Name}${Categories_Products[2].Description}
- 목록3 우측: ${Categories_Products[2].Price:C2}

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