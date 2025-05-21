//using DocuChef;
//using System.Diagnostics;
//using System.Text.Json;

//var basePath = @"D:\OneDrive_iyulab\iyulab\projects - 문서\Projects\IPX\Data";
//var templatePath = Path.Combine(basePath, "test-template.pptx");
////var templatePath = Path.Combine(basePath, "TEST.pptx");
//var outputPath = Path.Combine(basePath, "test-output.pptx");

//var dataPath = Path.Combine(basePath, "test-data.json");
//var json = File.ReadAllText(dataPath);
////var json = @"[
////{ ""Item_IP_대"": ""MININI"", ""Item_캐릭터_세분류_명"": ""BROWN"" },
////{ ""Item_IP_대"": ""MININI"", ""Item_캐릭터_세분류_명"": ""CONY"" },
////{ ""Item_IP_대"": ""BT21"", ""Item_캐릭터_세분류_명"": ""TATA"" },
////{ ""Item_IP_대"": ""BT21"", ""Item_캐릭터_세분류_명"": ""CHIMMY"" },
////{ ""Item_IP_대"": ""BT21"", ""Item_캐릭터_세분류_명"": ""COOKY"" },
////{ ""Item_IP_대"": ""BT21"", ""Item_캐릭터_세분류_명"": ""SHOOKY"" },
////{ ""Item_IP_대"": ""MININI"", ""Item_캐릭터_세분류_명"": ""BROWN"" },
////{ ""Item_IP_대"": ""MININI"", ""Item_캐릭터_세분류_명"": ""FRIENDS"" }
////]";
//var items = System.Text.Json.JsonSerializer.Deserialize<List<Dictionary<string, JsonElement>>>(json);

//var groupedItems = items
//    .GroupBy(item => item.ContainsKey("Item_IP_대") ? item["Item_IP_대"].ToString() : "null")
//    .Select(p =>
//    {
//        return new
//        {
//            p.Key,
//            Items = p.ToArray()
//        };
//    });

//var chef = new Chef(new RecipeOptions() { EnableVerboseLogging = true });
//var recipe = chef.LoadRecipe(templatePath);
//recipe.AddVariable("Products", groupedItems);
//recipe.Cook(outputPath);

//Process.Start(new ProcessStartInfo
//{
//    FileName = "explorer",
//    Arguments = outputPath,
//    UseShellExecute = true,
//    CreateNoWindow = true
//});