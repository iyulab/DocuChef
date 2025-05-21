/*
# 템플릿 엔진 동작 설명

## 템플릿 구조
```
[template.pptx]
* 슬라이드 1: #foreach: Categories, max: 1
  - ${Categories[0].Name}
* 슬라이드 2: #foreach: Categories>Products, max: 2
  - ${Categories>Products[0].Name}
  - ${Categories>Products[1].Name}
```

## 데이터 소스
```
Categories = [
  { "Food", Products: [ "Bibimbap", "Bulgogi" ] }
  { "Drinks", Products: [ "Americano", "Green Tea", "Milk" ] }
  { "Desserts", Products: [ "Tiramisu", "Chocolate Cake", "Fruit Salad", "Macaron", "Pudding" ] }
]
```

## 예상되는 슬라이드 생성 결과와 바인딩

### 슬라이드 1: 카테고리 "Food" (Categories, offset: 0)
- **바인딩 출력**: 
  - `${Categories[0].Name}` → "Food"

### 슬라이드 2: Food의 제품들 (Categories>Products, offset: 0)
- **바인딩 출력**: 
  - `${Categories>Products[0].Name}` → "Bibimbap"
  - `${Categories>Products[1].Name}` → "Bulgogi"

### 슬라이드 3: 카테고리 "Drinks" (Categories, offset: 1)
- **바인딩 출력**: 
  - `${Categories[0].Name}` → "Drinks"

### 슬라이드 4: Drinks의 제품들 (Categories>Products, offset: 0)
- **바인딩 출력**: 
  - `${Categories>Products[0].Name}` → "Americano"
  - `${Categories>Products[1].Name}` → "Green Tea"

### 슬라이드 5: Drinks의 제품들(추가) (Categories>Products, offset: 2)
- **바인딩 출력**: 
  - `${Categories>Products[0].Name}` → "Milk"
  - `${Categories>Products[1].Name}` → "" (비어있음 - 세 번째 항목만 존재)

### 슬라이드 6: 카테고리 "Desserts" (Categories, offset: 2)
- **바인딩 출력**: 
  - `${Categories[0].Name}` → "Desserts"

### 슬라이드 7: Desserts의 제품들 (Categories>Products, offset: 0)
- **바인딩 출력**: 
  - `${Categories>Products[0].Name}` → "Tiramisu"
  - `${Categories>Products[1].Name}` → "Chocolate Cake"

### 슬라이드 8: Desserts의 제품들(추가) (Categories>Products, offset: 2)
- **바인딩 출력**: 
  - `${Categories>Products[0].Name}` → "Fruit Salad"
  - `${Categories>Products[1].Name}` → "Macaron"

### 슬라이드 9: Desserts의 제품들(추가) (Categories>Products, offset: 4)
- **바인딩 출력**: 
  - `${Categories>Products[0].Name}` → "Pudding"
  - `${Categories>Products[1].Name}` → "" (비어있음 - 다섯 번째 항목만 존재)

## 동작 설명

1. **부모-자식 관계 처리**: 
   - 시스템은 각 카테고리와 그에 속한 제품들 사이의 부모-자식 관계를 인식합니다.
   - 카테고리 슬라이드 다음에 해당 카테고리의 제품 슬라이드들이 배치됩니다.

2. **바인딩 표현식 해석**:
   - `${Categories[0].Name}`는 현재 카테고리 컨텍스트에서 이름을 가져옵니다.
   - `${Categories>Products[0].Name}`, `${Categories>Products[1].Name}`는 현재 제품 슬라이드에서 표시될 제품들의 이름을 가져옵니다.

3. **그룹화 처리**:
   - `max: 2` 설정으로 인해 각 제품 슬라이드는 최대 2개의 제품만 표시합니다.
   - 두 개 이상의 제품이 있는 경우(Drinks, Desserts) 추가 슬라이드가 생성됩니다.
   - 제품 수가 홀수인 경우(Drinks: 3개, Desserts: 5개) 마지막 슬라이드에서 두 번째 바인딩(`${Categories>Products[1].Name}`)은 빈 값이 됩니다.

4. **컨텍스트 관리**:
   - 각 슬라이드는 자신의 특정 컨텍스트(카테고리 또는 제품 그룹)를 가지고 있습니다.
   - 중첩된 컬렉션(Categories>Products)에서는 부모 컨텍스트(현재 카테고리)도 유지되어 슬라이드 간의 관계가 보존됩니다.
 */

using System.Diagnostics;

namespace DocuChef.TestConsoleApp.NestedPPT
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("DocuChef Presentation Generator Test App");

            // Create sample data
            var data = CreateSampleData();

            // Print data
            PrintSampleData(data);

            // Set PowerPoint presentation generation parameters
            string basePath = AppDomain.CurrentDomain.BaseDirectory;
            string templatePath = Path.Combine(basePath, "files", "ppt", "nested_template.pptx");
            string logoPath = Path.Combine(basePath, "files", "logo.png");
            string outputPath = Path.Combine(basePath, "output_nested_template.pptx");

            Console.WriteLine($"\nTemplate file: {templatePath}");
            Console.WriteLine($"Output file: {outputPath}");

            try
            {
                // Initialize DocuChef engine
                var chef = new Chef(new RecipeOptions() { EnableVerboseLogging = true });
                var recipe = chef.LoadRecipe(templatePath);
                recipe.AddVariable(data);
                recipe.Cook(outputPath);

                Console.WriteLine("\nPowerPoint presentation generated successfully!");

                //Open the generated presentation
                OpenPresentationFile(outputPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"\nError: {ex.Message}");
                Console.WriteLine(ex.StackTrace);
            }

            Console.WriteLine("\nPress any key to exit...");
            Console.ReadKey();
        }

        /// <summary>
        /// Creates sample data for testing
        /// </summary>
        private static DataModel CreateSampleData()
        {
            // Create categories
            var categories = new List<Category>
            {
                new Category(1, "Food"),
                new Category(2, "Drinks"),
                new Category(3, "Desserts")
            };

            // Add products to categories
            categories[0].AddProduct(new Product(1, "Bibimbap", 8000, "Traditional Korean dish"));
            categories[0].AddProduct(new Product(2, "Bulgogi", 12000, "Korean BBQ"));

            categories[1].AddProduct(new Product(3, "Americano", 4500, "Espresso with water"));
            categories[1].AddProduct(new Product(4, "Green Tea", 3500, "Hot green tea"));
            categories[1].AddProduct(new Product(5, "Milk", 2000, "Fresh milk"));

            categories[2].AddProduct(new Product(6, "Tiramisu", 6000, "Italian dessert"));
            categories[2].AddProduct(new Product(7, "Chocolate Cake", 5500, "Sweet chocolate cake"));
            categories[2].AddProduct(new Product(8, "Fruit Salad", 5000, "Fresh fruit mix"));
            categories[2].AddProduct(new Product(9, "Macaron", 3000, "French dessert"));
            categories[2].AddProduct(new Product(10, "Pudding", 4000, "Soft pudding"));

            return new DataModel
            {
                Title = "Delicious Menu Introduction",
                Date = DateTime.Now.ToString("yyyy-MM-dd"),
                Author = "John Doe",
                Categories = categories,
                CompanyInfo = new CompanyInfo
                {
                    Name = "Delicious Restaurant",
                    Address = "123-45 Gangnam, Seoul",
                    Phone = "02-123-4567",
                    Website = "www.delicious-restaurant.com"
                }
            };
        }

        /// <summary>
        /// Prints sample data to console
        /// </summary>
        private static void PrintSampleData(DataModel data)
        {
            Console.WriteLine("\n===== Sample Data =====");
            Console.WriteLine($"Title: {data.Title}");
            Console.WriteLine($"Date: {data.Date}");
            Console.WriteLine($"Author: {data.Author}");

            Console.WriteLine("\n----- Categories and Products -----");
            foreach (var category in data.Categories)
            {
                Console.WriteLine($"Category: {category.Name} (ID: {category.Id})");
                foreach (var product in category.Products)
                {
                    Console.WriteLine($"  - {product.Name}: {product.Price:C} - {product.Description}");
                }
            }

            Console.WriteLine("\n----- Company Info -----");
            Console.WriteLine($"Company: {data.CompanyInfo.Name}");
            Console.WriteLine($"Address: {data.CompanyInfo.Address}");
            Console.WriteLine($"Phone: {data.CompanyInfo.Phone}");
            Console.WriteLine($"Website: {data.CompanyInfo.Website}");
        }

        /// <summary>
        /// Opens the generated presentation file
        /// </summary>
        private static void OpenPresentationFile(string filePath)
        {
            try
            {
                var processStartInfo = new ProcessStartInfo
                {
                    FileName = filePath,
                    UseShellExecute = true
                };

                Process.Start(processStartInfo);
                Console.WriteLine("Presentation opened.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Cannot open presentation: {ex.Message}");
            }
        }
    }

    /// <summary>
    /// Main data model class
    /// </summary>
    public class DataModel
    {
        public string Title { get; set; }
        public string Date { get; set; }
        public string Author { get; set; }
        public List<Category> Categories { get; set; }
        public CompanyInfo CompanyInfo { get; set; }
    }

    /// <summary>
    /// Company information class
    /// </summary>
    public class CompanyInfo
    {
        public string Name { get; set; }
        public string Address { get; set; }
        public string Phone { get; set; }
        public string Website { get; set; }
    }

    /// <summary>
    /// Category class
    /// </summary>  
    public class Category
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public List<Product> Products { get; set; }

        public Category(int id, string name)
        {
            Id = id;
            Name = name;
            Products = new List<Product>();
        }

        public void AddProduct(Product product)
        {
            Products.Add(product);
        }

        public override string ToString()
        {
            return $"{Name} (ID: {Id})";
        }
    }

    /// <summary>
    /// Product class
    /// </summary>
    public class Product
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public decimal Price { get; set; }
        public string Description { get; set; }
        public string Url { get; set; }

        public Product(int id, string name, decimal price, string description)
        {
            Id = id;
            Name = name;
            Price = price;
            Description = description;
            Url = $"https://placehold.co/60x60?text={name}";
        }

        public override string ToString()
        {
            return $"{Name}: {Price:C} - {Description}";
        }
    }
}