using System.Diagnostics;
using System.Text.Json;

namespace DocuChef.TestConsoleApp;

/*
# 현재 템플릿 구조
슬라이드1: 제목슬라이드 (바인딩 없음)
슬라이드2: ${Brands[0].Key} - Brands 반복
슬라이드3: ${Brands>Types[0].Key} & Items - Brands>Types 중첩 반복 
슬라이드4: EOD

# 문제점:
슬라이드3에서 Types[0]만 참조하고 있어서 Types 반복이 안됨.
현재는 각 Brand마다 하나의 Types 슬라이드만 생성됨.

# 원하는 결과:
슬라이드1: 제목슬라이드
슬라이드2: Brands[0] (B&F)
슬라이드3: B&F>Types[0] (KEYRING)
슬라이드4: B&F>Types[1] (FIGURINE)  
슬라이드5: B&F>Types[2] (TOY_SET)
슬라이드6: Brands[1] (BT21)
슬라이드7: BT21>Types[0] (PLUSH)
슬라이드8: EOD

# 해결방법:
슬라이드3을 Types 컬렉션을 반복하도록 수정하고
각 Type마다 별도 슬라이드가 생성되도록 해야 함.
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

        //var brands = items
        //   .GroupBy(item => item.TryGetValue("Item_IP_대", out JsonElement value) ? value.ToString() : "null")
        //   .Select(types =>
        //   {
        //       return new
        //       {
        //           types.Key,
        //           Types = types
        //            .GroupBy(type => type.TryGetValue("Stock_자재구분값_영어", out JsonElement value) ? value.ToString() : "null")
        //            .Select(items =>
        //            {
        //                return new
        //                {
        //                    items.Key,
        //                    Items = items
        //                };
        //            })
        //       };
        //   });

        // 테스트를 위해 더 다양한 브랜드와 타입을 생성
        var brands = new[]
        {
            new
            {
                Key = "B&F",
                Types = new[]
                {
                    new
                    {
                        Key = "KEYRING",
                        Items = new[]
                        {
                            new Dictionary<string, object> { ["Item_캐릭터_세분류_명"] = "BROWN" },
                            new Dictionary<string, object> { ["Item_캐릭터_세분류_명"] = "SALLY" }
                        }
                    },
                    new
                    {
                        Key = "FIGURINE",
                        Items = new[]
                        {
                            new Dictionary<string, object> { ["Item_캐릭터_세분류_명"] = "BROWN" },
                            new Dictionary<string, object> { ["Item_캐릭터_세분류_명"] = "CONY" }
                        }
                    },
                    new
                    {
                        Key = "TOY_SET",
                        Items = new[]
                        {
                            new Dictionary<string, object> { ["Item_캐릭터_세분류_명"] = "FRIENDS" }
                        }
                    }
                }
            },
            new
            {
                Key = "BT21",
                Types = new[]
                {
                    new
                    {
                        Key = "PLUSH",
                        Items = new[]
                        {
                            new Dictionary<string, object> { ["Item_캐릭터_세분류_명"] = "TATA" },
                            new Dictionary<string, object> { ["Item_캐릭터_세분류_명"] = "COOKY" }
                        }
                    }
                }
            }
        };

        recipe.AddVariable("Today", DateTime.Now);
        recipe.AddVariable("Brands", brands);
        recipe.Cook(outputPath);

        Process.Start(new ProcessStartInfo
        {
            FileName = "explorer",
            Arguments = outputPath,
            UseShellExecute = true,
            CreateNoWindow = true
        });
    }
}
