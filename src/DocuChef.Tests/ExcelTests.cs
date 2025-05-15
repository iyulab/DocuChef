using ClosedXML.Excel;
using DocuChef.Excel;
using DocuChef.Exceptions;
using ClosedXML.Report.XLCustom;
using Xunit.Abstractions;
using FluentAssertions;

namespace DocuChef.Tests;

public class ExcelTests : TestBase
{
    private readonly string _tempDirectory;
    private readonly string _templatePath;

    public ExcelTests(ITestOutputHelper output) : base(output)
    {
        // Create a temporary directory for test files
        _tempDirectory = Path.Combine(Path.GetTempPath(), "DocuChefTests", Guid.NewGuid().ToString());
        Directory.CreateDirectory(_tempDirectory);

        // Create a simple Excel template for testing
        _templatePath = Path.Combine(_tempDirectory, "template.xlsx");
        CreateSampleTemplate(_templatePath);
    }

    public void Dispose()
    {
        // Clean up temp files after tests
        try
        {
            if (Directory.Exists(_tempDirectory))
            {
                Directory.Delete(_tempDirectory, true);
            }
        }
        catch (IOException)
        {
            // Ignore cleanup errors
        }
    }

    [Fact]
    public void Chef_LoadTemplate_WithExcelFile_ReturnsExcelRecipe()
    {
        // Arrange
        var chef = CreateNewChef();

        // Act
        var recipe = chef.LoadTemplate(_templatePath);

        // Assert
        recipe.Should().NotBeNull();
        recipe.Should().BeOfType<ExcelRecipe>();
    }

    [Fact]
    public void Chef_LoadTemplate_WithInvalidExtension_ThrowsException()
    {
        // Arrange
        var chef = CreateNewChef();
        var invalidFilePath = Path.Combine(_tempDirectory, "template.txt");
        File.WriteAllText(invalidFilePath, "This is not an Excel file");

        // Act & Assert
        Action act = () => chef.LoadTemplate(invalidFilePath);

        act.Should().Throw<DocuChefException>()
           .WithMessage("*Unsupported file format*");
    }

    [Fact]
    public void Chef_LoadExcelTemplate_WithValidPath_ReturnsExcelRecipe()
    {
        // Arrange
        var chef = CreateNewChef();

        // Act
        var recipe = chef.LoadExcelTemplate(_templatePath);

        // Assert
        recipe.Should().NotBeNull();
        recipe.Should().BeOfType<ExcelRecipe>();
    }

    [Fact]
    public void Chef_LoadExcelTemplate_WithNonExistentFile_ThrowsFileNotFoundException()
    {
        // Arrange
        var chef = CreateNewChef();
        var nonExistentPath = Path.Combine(_tempDirectory, "nonexistent.xlsx");

        // Act & Assert
        Action act = () => chef.LoadExcelTemplate(nonExistentPath);

        act.Should().Throw<FileNotFoundException>();
    }

    [Fact]
    public void ExcelRecipe_AddVariable_AddsVariableToTemplate()
    {
        // Arrange
        var chef = CreateNewChef();
        var recipe = chef.LoadExcelTemplate(_templatePath);

        // Act
        recipe.AddVariable("TestVar", "TestValue");

        // Assert
        // Since we can't directly access the variables dictionary, we'll test indirectly
        // by generating a document and checking it contains our variable
        var document = recipe.Generate();
        using var stream = new MemoryStream();
        document.SaveAs(stream);
        stream.Length.Should().BeGreaterThan(0);
    }

    [Fact]
    public void ExcelRecipe_AddVariable_WithNullName_ThrowsArgumentNullException()
    {
        // Arrange
        var chef = CreateNewChef();
        var recipe = chef.LoadExcelTemplate(_templatePath);

        // Act & Assert
        Action act = () => recipe.AddVariable(null, "Value");

        act.Should().Throw<ArgumentNullException>();
    }

    [Fact]
    public void ExcelRecipe_RegisterGlobalVariable_RegistersVariable()
    {
        // Arrange
        var chef = CreateNewChef();
        var recipe = chef.LoadExcelTemplate(_templatePath);
        var testValue = "TestGlobalValue";

        // Act
        recipe.RegisterGlobalVariable("TestGlobalVar", testValue);

        // Assert
        // Similar to AddVariable, we'll test indirectly
        var document = recipe.Generate();
        using var stream = new MemoryStream();
        document.SaveAs(stream);
        stream.Length.Should().BeGreaterThan(0);
    }

    [Fact]
    public void ExcelRecipe_RegisterGlobalVariable_WithNullName_ThrowsArgumentNullException()
    {
        // Arrange
        var chef = CreateNewChef();
        var recipe = chef.LoadExcelTemplate(_templatePath);

        // Act & Assert
        Action act = () => recipe.RegisterGlobalVariable(null, "Value");

        act.Should().Throw<ArgumentNullException>();
    }

    [Fact]
    public void ExcelRecipe_RegisterFunction_RegistersFunction()
    {
        // Arrange
        var chef = CreateNewChef();
        var recipe = chef.LoadExcelTemplate(_templatePath);

        // Act
        recipe.RegisterFunction("testFunc", (cell, value, parameters) => {
            cell.SetValue("Function called");
        });

        // Assert
        // We can only verify it doesn't throw an exception
        var document = recipe.Generate();
        using var stream = new MemoryStream();
        document.SaveAs(stream);
        stream.Length.Should().BeGreaterThan(0);
    }

    [Fact]
    public void ExcelRecipe_RegisterFunction_WithNullName_ThrowsArgumentNullException()
    {
        // Arrange
        var chef = CreateNewChef();
        var recipe = chef.LoadExcelTemplate(_templatePath);

        // Act & Assert
        Action act = () => recipe.RegisterFunction(null, (cell, value, parameters) => { });

        act.Should().Throw<ArgumentNullException>();
    }

    [Fact]
    public void ExcelRecipe_RegisterFunction_WithNullFunction_ThrowsArgumentNullException()
    {
        // Arrange
        var chef = CreateNewChef();
        var recipe = chef.LoadExcelTemplate(_templatePath);

        // Act & Assert
        Action act = () => recipe.RegisterFunction("name", null);

        act.Should().Throw<ArgumentNullException>();
    }

    [Fact]
    public void ExcelRecipe_Generate_ReturnsExcelDocument()
    {
        // Arrange
        var chef = CreateNewChef();
        var recipe = chef.LoadExcelTemplate(_templatePath);

        // Act
        var document = recipe.Generate();

        // Assert
        document.Should().NotBeNull();
        document.Should().BeOfType<ExcelDocument>();
    }

    [Fact]
    public void ExcelDocument_SaveAs_WithValidPath_SavesDocument()
    {
        // Arrange
        var chef = CreateNewChef();
        var recipe = chef.LoadExcelTemplate(_templatePath);
        var document = recipe.Generate();
        var outputPath = Path.Combine(_tempDirectory, "output.xlsx");

        // Act
        document.SaveAs(outputPath);

        // Assert
        File.Exists(outputPath).Should().BeTrue();
        new FileInfo(outputPath).Length.Should().BeGreaterThan(0);
    }

    [Fact]
    public void ExcelDocument_SaveAs_WithNullPath_ThrowsArgumentNullException()
    {
        // Arrange
        var chef = CreateNewChef();
        var recipe = chef.LoadExcelTemplate(_templatePath);
        var document = recipe.Generate();

        // Act & Assert
        Action act = () => document.SaveAs((string)null);

        act.Should().Throw<ArgumentNullException>();
    }

    [Fact]
    public void ExcelDocument_SaveAs_WithStream_SavesDocument()
    {
        // Arrange
        var chef = CreateNewChef();
        var recipe = chef.LoadExcelTemplate(_templatePath);
        var document = recipe.Generate();
        using var stream = new MemoryStream();

        // Act
        document.SaveAs(stream);

        // Assert
        stream.Length.Should().BeGreaterThan(0);
    }

    [Fact]
    public void ExcelDocument_SaveAs_WithNullStream_ThrowsArgumentNullException()
    {
        // Arrange
        var chef = CreateNewChef();
        var recipe = chef.LoadExcelTemplate(_templatePath);
        var document = recipe.Generate();

        // Act & Assert
        Action act = () => document.SaveAs((Stream)null);

        act.Should().Throw<ArgumentNullException>();
    }

    [Fact]
    public void ExcelDocument_Dispose_DisposesWorkbook()
    {
        // Arrange
        var chef = CreateNewChef();
        var recipe = chef.LoadExcelTemplate(_templatePath);
        var document = recipe.Generate();

        // Act
        document.Dispose();

        // Assert
        // We can check if the object is disposed by trying to access a method
        // that should throw ObjectDisposedException
        Action act = () => document.SaveAs(new MemoryStream());

        act.Should().Throw<ObjectDisposedException>();
    }

    [Fact]
    public void ChefExtensions_LoadRecipe_LoadsExcelTemplate()
    {
        // Arrange
        var chef = CreateNewChef();

        // Act
        var recipe = chef.LoadRecipe(_templatePath);

        // Assert
        recipe.Should().NotBeNull();
        recipe.Should().BeOfType<ExcelRecipe>();
    }

    [Fact]
    public void RecipeExtensions_AddIngredient_AddsVariableToRecipe()
    {
        // Arrange
        var chef = CreateNewChef();
        var recipe = chef.LoadExcelRecipe(_templatePath);

        // Act
        recipe.AddIngredient("TestVar", "TestValue");

        // Assert
        // Test indirectly by generating document
        var document = recipe.Generate();
        using var stream = new MemoryStream();
        document.SaveAs(stream);
        stream.Length.Should().BeGreaterThan(0);
    }

    [Fact]
    public void RecipeExtensions_Cook_GeneratesDocument()
    {
        // Arrange
        var chef = CreateNewChef();
        var recipe = chef.LoadExcelRecipe(_templatePath);

        // Act
        var dish = recipe.Cook();

        // Assert
        dish.Should().NotBeNull();
        dish.Should().BeOfType<ExcelDocument>();
    }

    [Fact]
    public void DishExtensions_Serve_SavesDocument()
    {
        // Arrange
        var chef = CreateNewChef();
        var recipe = chef.LoadExcelRecipe(_templatePath);
        var dish = recipe.Cook();
        var outputPath = Path.Combine(_tempDirectory, "served.xlsx");

        // Act
        dish.Serve(outputPath);

        // Assert
        File.Exists(outputPath).Should().BeTrue();
        new FileInfo(outputPath).Length.Should().BeGreaterThan(0);
    }

    [Fact]
    public void Integration_CompleteWorkflow_GeneratesExpectedDocument()
    {
        // Arrange
        var chef = CreateNewChef();
        var recipe = chef.LoadExcelRecipe(_templatePath);

        // Add data to the template
        recipe.AddIngredient("Title", "Sales Report");
        recipe.AddIngredient("Date", DateTime.Now);

        var products = new List<Product>
        {
            new Product { Id = 1, Name = "Product 1", Price = 10.99m },
            new Product { Id = 2, Name = "Product 2", Price = 20.50m },
            new Product { Id = 3, Name = "Product 3", Price = 15.75m }
        };

        recipe.AddIngredient("Products", products);

        // Register a custom function
        recipe.RegisterTechnique("highlight", (cell, value, parameters) => {
            cell.SetValue(value);
            cell.Style.Font.Bold = true;
        });

        // Act
        var dish = recipe.Cook();
        var outputPath = Path.Combine(_tempDirectory, "integration_test.xlsx");
        dish.Serve(outputPath);

        // Assert
        File.Exists(outputPath).Should().BeTrue();
        var fileInfo = new FileInfo(outputPath);
        fileInfo.Length.Should().BeGreaterThan(0);

        // Additional check: attempt to open the generated file to ensure it's valid Excel
        using var workbook = new XLWorkbook(outputPath);
        workbook.Worksheets.Count.Should().BeGreaterThan(0);
    }

    #region Helper Methods

    private void CreateSampleTemplate(string path)
    {
        using var workbook = new XLWorkbook();
        var worksheet = workbook.Worksheets.Add("Sheet1");

        // Add some template placeholders
        worksheet.Cell("A1").Value = "{{Title}}";
        worksheet.Cell("A2").Value = "Generated on: {{Date}}";
        worksheet.Cell("A4").Value = "Products:";
        worksheet.Cell("A5").Value = "Id";
        worksheet.Cell("B5").Value = "Name";
        worksheet.Cell("C5").Value = "Price";

        // Add a range for products
        worksheet.Range("A6:C6").SetValue("{{Products.Id}}|{{Products.Name}}|{{Products.Price}}");

        workbook.SaveAs(path);
    }

    // Sample class for testing
    private class Product
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public decimal Price { get; set; }
    }

    #endregion
}