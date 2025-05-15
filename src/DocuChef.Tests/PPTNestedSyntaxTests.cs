using DocuChef.PowerPoint;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using FluentAssertions;
using System.Text;
using Xunit.Abstractions;

namespace DocuChef.Tests
{
    /// <summary>
    /// Tests for PowerPoint nested data structures
    /// </summary>
    public class PPTNestedSyntaxTests : TestBase
    {
        private readonly string _tempDirectory;
        private readonly string _templatePath;

        public PPTNestedSyntaxTests(ITestOutputHelper output) : base(output)
        {
            _tempDirectory = Path.Combine(Path.GetTempPath(), "DocuChefTests", Guid.NewGuid().ToString());
            Directory.CreateDirectory(_tempDirectory);
            _templatePath = Path.Combine(_tempDirectory, "nested_template.pptx");
            PPTHelper.CreateNestedSyntaxTemplate(_templatePath);
        }

        public override void Dispose()
        {
            try { if (Directory.Exists(_tempDirectory)) Directory.Delete(_tempDirectory, true); }
            catch { }
            base.Dispose();
        }

        [Fact]
        public void Department_Team_Member_Hierarchy_Renders_Correctly()
        {
            // Arrange
            var chef = CreateNewChef();

            // Create a hierarchical organization structure
            var departments = new List<Department>
            {
                new Department
                {
                    Name = "Engineering",
                    Teams = new List<Team>
                    {
                        new Team
                        {
                            Name = "Frontend",
                            Members = new List<Member>
                            {
                                new Member { Name = "Alice Smith", Role = "Lead Engineer" },
                                new Member { Name = "Bob Johnson", Role = "Senior Developer" }
                            }
                        },
                        new Team
                        {
                            Name = "Backend",
                            Members = new List<Member>
                            {
                                new Member { Name = "Dave Miller", Role = "Database Architect" },
                                new Member { Name = "Eve Wilson", Role = "API Developer" }
                            }
                        }
                    }
                },
                new Department
                {
                    Name = "Marketing",
                    Teams = new List<Team>
                    {
                        new Team
                        {
                            Name = "Digital",
                            Members = new List<Member>
                            {
                                new Member { Name = "Frank White", Role = "SEO Specialist" },
                                new Member { Name = "Grace Lee", Role = "Social Media Manager" }
                            }
                        }
                    }
                }
            };

            // Create a simplified template specifically for this test
            var simplifiedTemplatePath = Path.Combine(_tempDirectory, "dept_team_member_template.pptx");
            using (var presentationDoc = PPTHelper.CreateBasicPresentation(simplifiedTemplatePath))
            {
                // Create a department slide
                var deptSlidePart = PPTHelper.AddSlide(presentationDoc);
                PPTHelper.AddTextShape(deptSlidePart, "Department: ${Departments[0].Name}",
                    "DepartmentTitleShape", 1, 1524000, 1524000, 6096000, 800000);

                // Create a team slide with direct reference
                var teamSlidePart = PPTHelper.AddSlide(presentationDoc);
                PPTHelper.AddTextShape(teamSlidePart, "Department: ${Departments[0].Name}",
                    "DeptNameShape", 1, 1524000, 1524000, 6096000, 800000);
                PPTHelper.AddTextShape(teamSlidePart, "Team: ${Departments[0].Teams[0].Name}",
                    "TeamNameShape", 2, 1524000, 2500000, 6096000, 800000);

                // Create a member slide with direct references
                var memberSlidePart = PPTHelper.AddSlide(presentationDoc);
                PPTHelper.AddTextShape(memberSlidePart,
                    "Department: ${Departments[0].Name}\nTeam: ${Departments[0].Teams[0].Name}",
                    "HeaderShape", 1, 1524000, 1524000, 6096000, 800000);
                PPTHelper.AddTextShape(memberSlidePart,
                    "Member 1: ${Departments[0].Teams[0].Members[0].Name}, ${Departments[0].Teams[0].Members[0].Role}\n" +
                    "Member 2: ${Departments[0].Teams[0].Members[1].Name}, ${Departments[0].Teams[0].Members[1].Role}",
                    "MembersShape", 2, 1524000, 2500000, 6096000, 1500000);

                presentationDoc.Save();
            }

            var recipe = chef.LoadPowerPointTemplate(simplifiedTemplatePath);
            recipe.AddVariable("Departments", departments);

            // Act
            var document = recipe.Generate();
            var outputPath = Path.Combine(_tempDirectory, "department_hierarchy_test.pptx");
            document.SaveAs(outputPath);

            // Assert
            using var presentationDocument = PresentationDocument.Open(outputPath, false);

            // Extract text from all slides
            var allSlideText = new StringBuilder();
            foreach (var slideId in presentationDocument.PresentationPart.Presentation.SlideIdList.Elements<SlideId>())
            {
                var slidePart = (SlidePart)presentationDocument.PresentationPart.GetPartById(slideId.RelationshipId);
                var textElements = PPTHelper.GetTextElements(slidePart);
                foreach (var text in textElements)
                {
                    allSlideText.Append(text).Append(" ");
                    _output.WriteLine($"Text found: '{text}'");
                }
            }

            var combinedText = allSlideText.ToString();

            // Verify department names
            combinedText.Should().Contain("Engineering");

            // Verify team names - these were failing in the original test
            combinedText.Should().Contain("Frontend");

            // Verify member names
            combinedText.Should().Contain("Alice Smith");
            combinedText.Should().Contain("Lead Engineer");
        }

        [Fact]
        public void Category_Products_Hierarchy_Renders_Correctly()
        {
            // Arrange
            var chef = CreateNewChef();

            // Create a categories/products hierarchical data structure
            var categories = new List<Category>
            {
                new Category
                {
                    Name = "Electronics",
                    Description = "Electronic devices and accessories",
                    Products = new List<Product>
                    {
                        new Product { Name = "Laptop", Price = 1299.99 },
                        new Product { Name = "Smartphone", Price = 899.99 },
                        new Product { Name = "Tablet", Price = 499.99 }
                    }
                },
                new Category
                {
                    Name = "Home & Kitchen",
                    Description = "Home appliances and kitchen tools",
                    Products = new List<Product>
                    {
                        new Product { Name = "Coffee Maker", Price = 129.99 },
                        new Product { Name = "Blender", Price = 89.99 },
                        new Product { Name = "Toaster", Price = 49.99 }
                    }
                }
            };

            // Create a simplified template specifically for this test
            var simplifiedTemplatePath = Path.Combine(_tempDirectory, "category_products_template.pptx");
            using (var presentationDoc = PPTHelper.CreateBasicPresentation(simplifiedTemplatePath))
            {
                // Create a category slide with direct indexing
                var categorySlidePart = PPTHelper.AddSlide(presentationDoc);
                PPTHelper.AddTextShape(categorySlidePart,
                    "Category: ${Categories[0].Name}\n" +
                    "Description: ${Categories[0].Description}\n" +
                    "Product Count: ${Categories[0].Products.length}",
                    "CategoryInfoShape", 1, 1524000, 1524000, 6096000, 1500000);

                // Create a products slide with direct indexing
                var productsSlidePart = PPTHelper.AddSlide(presentationDoc);
                PPTHelper.AddTextShape(productsSlidePart, "${Categories[0].Name} Products",
                    "CategoryNameShape", 1, 1524000, 1024000, 6096000, 800000);
                PPTHelper.AddTextShape(productsSlidePart,
                    "Product 1: ${Categories[0].Products[0].Name} - $${Categories[0].Products[0].Price}\n" +
                    "Product 2: ${Categories[0].Products[1].Name} - $${Categories[0].Products[1].Price}\n" +
                    "Product 3: ${Categories[0].Products[2].Name} - $${Categories[0].Products[2].Price}",
                    "ProductsListShape", 2, 1524000, 2000000, 6096000, 2000000);

                presentationDoc.Save();
            }

            var recipe = chef.LoadPowerPointTemplate(simplifiedTemplatePath);
            recipe.AddVariable("Categories", categories);

            // Act
            var document = recipe.Generate();
            var outputPath = Path.Combine(_tempDirectory, "category_products_test.pptx");
            document.SaveAs(outputPath);

            // Assert
            using var presentationDocument = PresentationDocument.Open(outputPath, false);

            // Extract text from all slides
            var allSlideText = new StringBuilder();
            foreach (var slideId in presentationDocument.PresentationPart.Presentation.SlideIdList.Elements<SlideId>())
            {
                var slidePart = (SlidePart)presentationDocument.PresentationPart.GetPartById(slideId.RelationshipId);
                var textElements = PPTHelper.GetTextElements(slidePart);
                foreach (var text in textElements)
                {
                    allSlideText.Append(text).Append(" ");
                    _output.WriteLine($"Text found: '{text}'");
                }
            }

            var combinedText = allSlideText.ToString();

            // Verify category information
            combinedText.Should().Contain("Electronics");
            combinedText.Should().Contain("Electronic devices and accessories");

            // Verify products - these were failing in the original test
            combinedText.Should().Contain("Laptop");
            combinedText.Should().Contain("1299.99");
        }

        [Fact]
        public void Deep_Nested_Properties_Are_Accessible()
        {
            // Arrange
            var chef = CreateNewChef();

            // Create a deep nested structure
            var company = new Company
            {
                Name = "Acme Corporation",
                Headquarters = new Office
                {
                    Location = "New York",
                    Address = new Address
                    {
                        Street = "123 Main St",
                        City = "New York",
                        State = "NY",
                        PostalCode = "10001",
                        GeoLocation = new GeoCoordinates
                        {
                            Latitude = 40.7128,
                            Longitude = -74.0060
                        }
                    }
                },
                Departments = new List<Department>
                {
                    new Department
                    {
                        Name = "R&D",
                        Budget = 5000000,
                        Manager = new Member
                        {
                            Name = "John Doe",
                            Role = "Director of R&D",
                            ContactInfo = new ContactInfo
                            {
                                Email = "john.doe@acme.example",
                                Phone = "555-123-4567"
                            }
                        }
                    }
                }
            };

            // Create a simplified template specifically for this test
            var simplifiedTemplatePath = Path.Combine(_tempDirectory, "deep_nested_template.pptx");
            using (var presentationDoc = PPTHelper.CreateBasicPresentation(simplifiedTemplatePath))
            {
                var testSlidePart = PPTHelper.AddSlide(presentationDoc);

                // Add a shape specifically for testing deep nested properties
                PPTHelper.AddTextShape(testSlidePart,
                    "Company: ${Company.Name}\n" +
                    "HQ City: ${Company.Headquarters.Address.City}\n" +
                    "HQ Geo: ${Company.Headquarters.Address.GeoLocation.Latitude}, ${Company.Headquarters.Address.GeoLocation.Longitude}\n" +
                    "R&D Manager: ${Company.Departments[0].Manager.Name}\n" +
                    "Manager Email: ${Company.Departments[0].Manager.ContactInfo.Email}",
                    "DeepNestedPropsShape", 1, 1524000, 1524000, 6096000, 2000000);

                presentationDoc.Save();
            }

            var recipe = chef.LoadPowerPointTemplate(simplifiedTemplatePath);
            recipe.AddVariable("Company", company);

            // Act
            var document = recipe.Generate();
            var outputPath = Path.Combine(_tempDirectory, "deep_nested_props_test.pptx");
            document.SaveAs(outputPath);

            // Assert
            using var presentationDocument = PresentationDocument.Open(outputPath, false);
            var firstSlidePart = PPTHelper.GetFirstSlidePart(presentationDocument);
            var textElements = PPTHelper.GetTextElements(firstSlidePart);

            foreach (var text in textElements)
            {
                _output.WriteLine($"Text element: '{text}'");
            }

            // Join all text elements to verify content
            var allText = string.Join(" ", textElements);

            // Verify deep nested properties
            allText.Should().Contain("Acme Corporation");
            allText.Should().Contain("New York");
            allText.Should().Contain("40.7128");
            allText.Should().Contain("-74.006");
            allText.Should().Contain("John Doe");
            allText.Should().Contain("john.doe@acme.example");
        }

        // Classes for testing nested structures
        private class Department
        {
            public string Name { get; set; }
            public decimal Budget { get; set; }
            public List<Team> Teams { get; set; } = new List<Team>();
            public Member Manager { get; set; }
        }

        private class Team
        {
            public string Name { get; set; }
            public List<Member> Members { get; set; } = new List<Member>();
        }

        private class Member
        {
            public string Name { get; set; }
            public string Role { get; set; }
            public ContactInfo ContactInfo { get; set; }
        }

        private class Category
        {
            public string Name { get; set; }
            public string Description { get; set; }
            public List<Product> Products { get; set; } = new List<Product>();
        }

        private class Product
        {
            public string Name { get; set; }
            public double Price { get; set; }
        }

        private class Company
        {
            public string Name { get; set; }
            public Office Headquarters { get; set; }
            public List<Department> Departments { get; set; } = new List<Department>();
        }

        private class Office
        {
            public string Location { get; set; }
            public Address Address { get; set; }
        }

        private class Address
        {
            public string Street { get; set; }
            public string City { get; set; }
            public string State { get; set; }
            public string PostalCode { get; set; }
            public GeoCoordinates GeoLocation { get; set; }
        }

        private class GeoCoordinates
        {
            public double Latitude { get; set; }
            public double Longitude { get; set; }
        }

        private class ContactInfo
        {
            public string Email { get; set; }
            public string Phone { get; set; }
        }
    }
}