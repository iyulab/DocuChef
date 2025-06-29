# template_1.pptx
## 슬라이드1:
- ${ppt.Image(LogoPath)}
- ${Title}
- BOLD ${Subtitle} Italic
- Created By: ${Date:yyyy-MM-dd}
*/
// TestTemplate.Run("template_1.pptx");

/* 
# template_2.pptx
## 슬라이드1:
- ${ppt.Image(LogoPath)}
- ${Title}
- BOLD ${Subtitle} Italic
- Created By: ${Date:yyyy-MM-dd}

## 슬라이드2:
- ${ppt.Image(LogoPath)}
- ${CompanyName}
- ${Items[0].Id}. ${Items[0].Name} - ${Items[0].Description}
Price: ${Items[0].Price:C0} USD
- ${Items[1].Id}. ${Items[1].Name} - ${Items[1].Description}
Price: ${Items[1].Price:C0} USD
- ${Items[2].Id}. ${Items[2].Name} - ${Items[2].Description}
Price: ${Items[2].Price:C0} USD
- ${Items[3].Id}. ${Items[3].Name} - ${Items[3].Description}
Price: ${Items[3].Price:C0} USD
- ${Items[4].Id}. ${Items[4].Name} - ${Items[4].Description}
Price: ${Items[4].Price:C0} USD
*/
// TestTemplate.Run("template_2.pptx");

/* 
# template_3.pptx
## 슬라이드1:
- ${ppt.Image(LogoPath)}
- ${Title}
- BOLD ${Subtitle} Italic
- Created By: ${Date:yyyy-MM-dd}

## 슬라이드2:
- ${ppt.Image(LogoPath)}
- ${CompanyName}
- ${Items[0].Id}. ${Items[0].Name} - ${Items[0].Description}
Price: ${Items[0].Price:C0} USD
- ${Items[1].Id}. ${Items[1].Name} - ${Items[1].Description}
Price: ${Items[1].Price:C0} USD
- ${Items[2].Id}. ${Items[2].Name} - ${Items[2].Description}
Price: ${Items[2].Price:C0} USD
- ${Items[3].Id}. ${Items[3].Name} - ${Items[3].Description}
Price: ${Items[3].Price:C0} USD
- ${Items[4].Id}. ${Items[4].Name} - ${Items[4].Description}
Price: ${Items[4].Price:C0} USD
- ${ppt.Image(Items[0].ImageUrl)}
- ${ppt.Image(Items[1].ImageUrl)}
- ${ppt.Image(Items[2].ImageUrl)}
- ${ppt.Image(Items[3].ImageUrl)}
- ${ppt.Image(Items[4].ImageUrl)}

## 슬라이드3: 바인딩 없음
- END