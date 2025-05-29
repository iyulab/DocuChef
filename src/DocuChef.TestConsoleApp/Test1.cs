using System.Diagnostics;
using System.Text.Json;

namespace DocuChef.TestConsoleApp;

/* 
===============================================================================
TEST.pptx 템플릿 분석 및 중첩 foreach 구현 테스트
===============================================================================

📋 템플릿 구조:
┌─────────────┬─────────────────────────────────────────────────────────────┐
│ 슬라이드 1  │ ${Today:yyyy-MM-dd}                           │ 정적 (제목)  │
│ 슬라이드 2  │ ${Products[0].Key}                           │ 그룹 헤더    │
│ 슬라이드 3  │ ${Products>Items[0]["Item_캐릭터_세분류_명"]} │ 아이템 목록  │
│             │ ${Products>Items[1]["Item_캐릭터_세분류_명"]} │ (2개씩 표시) │
│ 슬라이드 4  │ "END"                                        │ 정적 (종료)  │
└─────────────┴─────────────────────────────────────────────────────────────┘

📊 데이터 구조:
groupedItems = [
  { Key: "MININI", Items: ["BROWN", "CONY", "BROWN", "FRIENDS"] },    // 4개
  { Key: "BT21",   Items: ["TATA", "CHIMMY", "COOKY", "SHOOKY"] }     // 4개
]

🎯 예상 결과 (중첩 foreach 적용):
1. 제목: "2025-05-28"
2. 그룹1 헤더: "MININI" 
3. 그룹1 목록1: "BROWN, CONY" (Items 1~2)
4. 그룹1 목록2: "BROWN, FRIENDS" (Items 3~4)
5. 그룹2 헤더: "BT21"
6. 그룹2 목록1: "TATA, CHIMMY" (Items 1~2)  
7. 그룹2 목록2: "COOKY, SHOOKY" (Items 3~4)
8. 종료: "END"

🔧 구현 미션:
1. 슬라이드2: #foreach: Products (그룹별 반복)
2. 슬라이드3: #foreach: Products>Items, max: 2 (그룹 내 아이템 2개씩 반복)
3. 중첩된 foreach 처리로 총 8개 슬라이드 생성
4. Context Operator (>) 올바른 해석
5. 한국어 키("Item_캐릭터_세분류_명") 정상 바인딩

�📖 참고: SYNTAX_OF_PPT.md - "Nested Context Example", "Array Data Processing"
===============================================================================
*/

internal class Test1
{
  public static void Run(string fileName)
  {
    var basePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "files", "ppt");
    var templatePath = Path.Combine(basePath, fileName);
    var outputPath = Path.Combine(basePath, "test-output.pptx");

    Console.WriteLine("DocuChef Test1 - PowerPoint Template Processing");
    Console.WriteLine("===============================================\n");

    // Step 1: Inspect template before processing
    Console.WriteLine("Step 1: Template Inspection");
    Console.WriteLine("============================");
    TemplateInspector.InspectTemplate(templatePath);
    Console.WriteLine();

    // Step 2: Prepare test data
    Console.WriteLine("Step 2: Preparing Test Data");
    Console.WriteLine("============================");
        //var dataPath = Path.Combine(basePath, "test-data.json");
        //var json = File.ReadAllText(dataPath);
    var json = @"[
    { ""Item_IP_대"": ""MININI"", ""Item_캐릭터_세분류_명"": ""BROWN"", ""Hello"": ""World"" },
    { ""Item_IP_대"": ""MININI"", ""Item_캐릭터_세분류_명"": ""CONY"", ""Hello"": ""World"" },
    { ""Item_IP_대"": ""BT21"", ""Item_캐릭터_세분류_명"": ""TATA"", ""Hello"": ""World"" },
    { ""Item_IP_대"": ""BT21"", ""Item_캐릭터_세분류_명"": ""CHIMMY"", ""Hello"": ""World"" },
    { ""Item_IP_대"": ""BT21"", ""Item_캐릭터_세분류_명"": ""COOKY"", ""Hello"": ""World"" },
    { ""Item_IP_대"": ""BT21"", ""Item_캐릭터_세분류_명"": ""SHOOKY"", ""Hello"": ""World"" },
    { ""Item_IP_대"": ""MININI"", ""Item_캐릭터_세분류_명"": ""BROWN"", ""Hello"": ""World"" },
    { ""Item_IP_대"": ""MININI"", ""Item_캐릭터_세분류_명"": ""FRIENDS"", ""Hello"": ""World"" },
    { ""Item_IP_대"": ""MININI"", ""Item_캐릭터_세분류_명"": ""HELLO"", ""Hello"": ""World"" }
    ]";
    var items = System.Text.Json.JsonSerializer.Deserialize<List<Dictionary<string, JsonElement>>>(json);

    var groupedItems = items
       .GroupBy(item => item.ContainsKey("Item_IP_대") ? item["Item_IP_대"].ToString() : "null")
       .Select(p =>
       {
         return new
         {
           p.Key,
           Items = p.ToArray()
         };
       });

    Console.WriteLine($"✅ Prepared {items.Count} data items, grouped into {groupedItems.Count()} categories");
    foreach (var group in groupedItems)
    {
      Console.WriteLine($"   - {group.Key}: {group.Items.Length} items");
    }
    Console.WriteLine();

    // Step 3: Process template with DocuChef
    Console.WriteLine("Step 3: Processing Template with DocuChef");
    Console.WriteLine("==========================================");
    var chef = new Chef(new RecipeOptions() { EnableVerboseLogging = true });
    var recipe = chef.LoadRecipe(templatePath);
    recipe.AddVariable("Products", groupedItems);
    recipe.AddVariable("Today", DateTime.Now);

    Console.WriteLine("🍳 Cooking template...");
    recipe.Cook(outputPath);
    Console.WriteLine($"✅ Template processed successfully! Output: {outputPath}\n");    // Step 4: Verify output
    Console.WriteLine("Step 4: Output Verification");
    Console.WriteLine("============================");
    VerifyOutput.CheckDataBinding(outputPath);
    Console.WriteLine();

    // Step 4.5: Inspect generated file structure
    Console.WriteLine("Step 4.5: Generated File Analysis");
    Console.WriteLine("==================================");
    TemplateInspector.InspectGeneratedFile(outputPath);
    Console.WriteLine();

    // Step 5: Open output file
    Console.WriteLine("Step 5: Opening Output File");
    Console.WriteLine("============================");
    Console.WriteLine("📂 Opening generated PowerPoint file...");
    Process.Start(new ProcessStartInfo
    {
      FileName = "explorer",
      Arguments = outputPath,
      UseShellExecute = true,
      CreateNoWindow = true
    });

    Console.WriteLine("✅ Test1 completed successfully!");
  }
}