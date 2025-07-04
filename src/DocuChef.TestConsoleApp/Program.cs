using DocuChef.TestConsoleApp;

//TestTemplate.Run("template_3.pptx");
//Test1.Run("TEST.pptx");
// Test1.Run("TEST2.pptx");

//DeepTest.Run();
//AliasTest.Run();

// 진행률 기능 테스트
Console.WriteLine("Choose test to run:");
Console.WriteLine("1. Progress Demo");
Console.WriteLine("2. Detailed Progress Demo");
Console.WriteLine("3. Alias Test (existing)");
Console.WriteLine("4. Exit");
Console.Write("Enter choice (1-4): ");

var choice = Console.ReadLine();
switch (choice)
{
    case "1":
        ProgressTest.RunProgressDemo();
        break;
    case "2":
        ProgressTest.RunDetailedProgressDemo();
        break;
    case "3":
        // AliasTest.Run();
        Console.WriteLine("AliasTest is not implemented yet.");
        break;
    case "4":
        Console.WriteLine("Exiting...");
        break;
    default:
        Console.WriteLine("Invalid choice. Running Progress Demo...");
        ProgressTest.RunProgressDemo();
        break;
}