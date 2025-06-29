using System.Text.Json;

namespace DocuChef.TestConsoleApp;

/*
# 슬라이드1: 제목슬라이드 (바인딩 없음)

# 슬라이드2: 
${Brands[0].Key}

# 슬라이드3:
제목:
${Brands>Types[0].Key}

본문
${Brands>Types>Items[0]["Item_캐릭터_세분류_명"]}
${Brands>Types>Items[1]["Item_캐릭터_세분류_명"]}
${Brands>Types>Items[2]["Item_캐릭터_세분류_명"]}

# 슬라이드4: EOD

---

# 내가 원하는 결과

슬라이드1: 제목슬라이드
슬라이드2: Brands[0] 반복 구간시작
슬라이드3: Brands[0].Types[0] 반복 구간 시작
슬라이드4: Brands[0].Types[1] (생성)
슬라이드5: Brands[0].Types[2] (생성)
슬라이드6: Brands[1] (생성)
슬라이드7: Brands[1].Types[0] (생성)
슬라이드8: EOD
*/

internal class DeepTest
{
    internal static void Run()
    {

        var basePath = AppDomain.CurrentDomain.BaseDirectory;
        var templatePath = Path.Combine(basePath, "files", "deep", "test-template.pptx");
        var outputPath = Path.Combine(basePath, "files", "deep", "test-output.pptx");
        var dataPath = Path.Combine(basePath, "files", "deep", "test-data.json");

        var chef = new Chef(new RecipeOptions() { EnableVerboseLogging = true });
        var recipe = chef.LoadRecipe(templatePath);

        var json = File.ReadAllText(dataPath);
        var items = System.Text.Json.JsonSerializer.Deserialize<List<Dictionary<string, JsonElement>>>(json);

        var brands = items
           .GroupBy(item => item.TryGetValue("Item_IP_대", out JsonElement value) ? value.ToString() : "null")
           .Select(types =>
           {
               return new
               {
                   types.Key,
                   Types = types
                    .GroupBy(type => type.TryGetValue("Stock_자재구분값_영어", out JsonElement value) ? value.ToString() : "null")
                    .Select(items =>
                    {
                        return new
                        {
                            items.Key,
                            Items = items
                        };
                    })
               };
           });

        recipe.AddVariable("Today", DateTime.Now);
        recipe.AddVariable("Brands", brands);
        recipe.Cook(outputPath);

    }
}
