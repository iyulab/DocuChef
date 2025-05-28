using DocuChef.Presentation.Processors;
using System;
using System.Collections.Generic;

namespace DocuChef.TestConsoleApp
{
    public class DebugContextOperator
    {
        public static void TestContextOperator()
        {
            Console.WriteLine("=== Debug Context Operator ===");
            
            var dataBinder = new DataBinder();
            
            // Simple test case: Categories>Items[0].Name
            var categories = new[]
            {
                new 
                { 
                    Name = "Electronics",
                    Items = new[]
                    {
                        new { Name = "Smartphone", Price = 999 },
                        new { Name = "Laptop", Price = 1299 }
                    }
                }
            };
            var data = new Dictionary<string, object> { ["Categories"] = categories };
            
            Console.WriteLine("Data structure:");
            Console.WriteLine($"Categories[0].Name = {categories[0].Name}");
            Console.WriteLine($"Categories[0].Items[0].Name = {categories[0].Items[0].Name}");
            
            var expression = "Categories>Items[0].Name";
            Console.WriteLine($"\nTesting expression: {expression}");
            
            var result = dataBinder.ResolveExpression(expression, data);
            Console.WriteLine($"Result: '{result}'");
            
            // Test multiple nesting levels
            Console.WriteLine("\n=== Multiple Nesting Levels ===");
            
            var companies = new[]
            {
                new 
                { 
                    Name = "TechCorp",
                    Departments = new[]
                    {
                        new 
                        {
                            Name = "Engineering",
                            Teams = new[]
                            {
                                new { Name = "Frontend Team", Size = 5 },
                                new { Name = "Backend Team", Size = 8 }
                            }
                        }
                    }
                }
            };
            var companyData = new Dictionary<string, object> { ["Companies"] = companies };
            
            Console.WriteLine("Data structure:");
            Console.WriteLine($"Companies[0].Name = {companies[0].Name}");
            Console.WriteLine($"Companies[0].Departments[0].Teams[0].Name = {companies[0].Departments[0].Teams[0].Name}");
            
            var complexExpression = "Companies>Departments>Teams[0].Name";
            Console.WriteLine($"\nTesting expression: {complexExpression}");
            
            var complexResult = dataBinder.ResolveExpression(complexExpression, companyData);
            Console.WriteLine($"Result: '{complexResult}'");
        }
    }
}
