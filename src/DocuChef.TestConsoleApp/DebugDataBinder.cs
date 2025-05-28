using DocuChef.Presentation.Processors;

namespace DocuChef.TestConsoleApp;

public class DebugDataBinder
{
    public static void Test()
    {
        var dataBinder = new DataBinder();
        
        // Test 1: Simple case
        Console.WriteLine("=== Test 1: Simple Expression ===");
        var data1 = new Dictionary<string, object> { ["Name"] = "Test Product" };
        var result1 = dataBinder.ResolveExpression("Name", data1);
        Console.WriteLine($"Input: 'Name', Expected: 'Test Product', Actual: '{result1}'");
        
        // Test 2: Context operator case
        Console.WriteLine("\n=== Test 2: Context Operator Expression ===");
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
        var data2 = new Dictionary<string, object> { ["Categories"] = categories };
        var result2 = dataBinder.ResolveExpression("Categories>Items[0].Name", data2);
        Console.WriteLine($"Input: 'Categories>Items[0].Name', Expected: 'Smartphone', Actual: '{result2}'");
        
        // Test 3: Let's try a simpler context operator case
        Console.WriteLine("\n=== Test 3: Simpler Context Operator ===");
        var result3 = dataBinder.ResolveExpression("Categories>Items", data2);
        Console.WriteLine($"Input: 'Categories>Items', Expected: Array, Actual: '{result3}'");
    }
}
