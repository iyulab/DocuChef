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

� 구현 현황 (2025-01-11):
┌─────────────────────────┬─────────┬────────────────────────────────────┐
│ 구현 영역               │ 상태    │ 비고                               │
├─────────────────────────┼─────────┼────────────────────────────────────┤
│ 1. Template Analysis    │ ✅ 완료 │ 표현식/디렉티브 올바르게 식별      │
│ 2. Slide Plan Generation│ ✅ 완료 │ 8개 슬라이드 구조 정확히 생성      │
│ 3. Nested foreach Logic │ ✅ 완료 │ Products>Items 중첩 처리 구현      │
│ 4. Slide Structure Gen  │ ✅ 완료 │ 올바른 RelationshipId로 슬라이드 복제│
│ 5. Context Operator     │ ✅ 완료 │ "Products>Items" 문법 해석         │
│ 6. Index Offset Calc    │ ✅ 완료 │ 중첩 아이템 인덱스 (0,2 오프셋)   │
│ 7. Data Binding         │ ❌ 대기 │ 표현식 바인딩 미완료 (주석 처리됨) │
│ 8. DataBinder 통합      │ ❌ 대기 │ 생성된 슬라이드에 데이터 바인딩    │
└─────────────────────────┴─────────┴────────────────────────────────────┘

🏗️ 핵심 구현 방법론:
• SlidePlanGenerator.ProcessNestedContextSlide(): 중첩 foreach 로직 구현
  - 부모 컬렉션(Products) 반복 + 자식 컬렉션(Items) 페이징 처리
  - IndexOffset 계산으로 중첩 아이템 올바른 인덱싱 (0, 2, 4...)
• ContextBasedPowerPointProcessor: 4단계 파이프라인 실행
  - Analysis → Planning → Generation → Binding (Binding 단계 주석 처리됨)
• SlideGenerator: RelationshipId 기반 슬라이드 복제 및 구조 생성

🔍 다음 단계:
1. DataBinder 통합 활성화 (ContextBasedPowerPointProcessor L111 주석 해제)
2. 표현식 바인딩 검증 (${Today:yyyy-MM-dd}, ${Products[0].Key} 등)
3. 중첩 컨텍스트 바인딩 확인 (${Products>Items[0]["Item_캐릭터_세분류_명"]})

�📖 참고: SYNTAX_OF_PPT.md - "Nested Context Example", "Array Data Processing"
===============================================================================
*/

internal class Test1
{
    public static void Run()
    {
        var basePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "files", "ppt");
        var templatePath = Path.Combine(basePath, "TEST.pptx");
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
        // var dataPath = Path.Combine(basePath, "test-data.json");
        // var json = File.ReadAllText(dataPath);
        var json = @"[
        { ""Item_IP_대"": ""MININI"", ""Item_캐릭터_세분류_명"": ""BROWN"" },
        { ""Item_IP_대"": ""MININI"", ""Item_캐릭터_세분류_명"": ""CONY"" },
        { ""Item_IP_대"": ""BT21"", ""Item_캐릭터_세분류_명"": ""TATA"" },
        { ""Item_IP_대"": ""BT21"", ""Item_캐릭터_세분류_명"": ""CHIMMY"" },
        { ""Item_IP_대"": ""BT21"", ""Item_캐릭터_세분류_명"": ""COOKY"" },
        { ""Item_IP_대"": ""BT21"", ""Item_캐릭터_세분류_명"": ""SHOOKY"" },
        { ""Item_IP_대"": ""MININI"", ""Item_캐릭터_세분류_명"": ""BROWN"" },
        { ""Item_IP_대"": ""MININI"", ""Item_캐릭터_세분류_명"": ""FRIENDS"" }
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

        Console.WriteLine("🍳 Cooking template...");
        recipe.Cook(outputPath);
        Console.WriteLine($"✅ Template processed successfully! Output: {outputPath}\n");

        // Step 4: Verify output
        Console.WriteLine("Step 4: Output Verification");
        Console.WriteLine("============================");
        VerifyOutput.CheckDataBinding(outputPath);
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