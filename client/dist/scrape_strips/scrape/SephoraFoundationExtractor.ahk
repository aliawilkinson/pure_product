#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.



inputbox, OverrideVar, Input Override, Number of Cycles , , , , , , , , 1

If !OverrideVar
{
	OverrideVar := 1
}


ExcelInterface := ComObjCreate("Excel.Application")

;FILEPATH := "C:\Users\Christian\Desktop\FoundationResults.XLS"

FILEPATH := "C:\Users\Christian\Desktop\imagetest.XLSX"

;WORKBOOK := ExcelInterface.WORKBOOKS.OPEN(FILEPATH,, readonly := true)

;wb:=ExcelInterface.Workbooks.Add(),wb.SaveAs(FILEPATH)

wb:=ExcelInterface.Workbooks.Open(FILEPATH)

ExcelInterface.Visible := True

row := 1

col := 1


wb.SHEETS("Sheet1").CELLS(row,col).VALUE := "Brand"

col += 1

wb.SHEETS("Sheet1").CELLS(row,col).VALUE := "Name"

col += 1

wb.SHEETS("Sheet1").CELLS(row,col).VALUE := "Price"

col += 1

wb.SHEETS("Sheet1").CELLS(row,col).VALUE := "Size"

col += 1

wb.SHEETS("Sheet1").CELLS(row,col).VALUE := "Rating"

col += 1

wb.SHEETS("Sheet1").CELLS(row,col).VALUE := "Details"

col += 1

wb.SHEETS("Sheet1").CELLS(row,col).VALUE := "Ingredients"




col += 1

wb.SHEETS("Sheet1").CELLS(row,col).VALUE := "Image URL"



col := 1

row := 2





PriceArray := Object()

IEInterface := ComObjCreate("InternetExplorer.Application")

IEInterface.Visible := True


URL := "https://www.sephora.com/shop/foundation-makeup?pageSize=300"

BaseUrl := "https://www.sephora.com"

IEInterface.Navigate(URL)

IELOAD(IEInterface)



sleep, 3000

LinkDeposit := Object()

Div := 915

;DivIterator := 17

DivIterator := 1

PriceDiv := 14

PriceDivIterator := 3

Loop, 
{
	
	try
		textVar := IEInterface.Document.GetElementsByTagName("DIV")[Div].OuterHTML
	Catch
	{
		break
	}
	
	
	
	Clipboard := textVar
	
	;msgbox %textVar%
	
	
	
	RegExMatch(textVar, "href=.+?(:product)", FinalText)
	
	StringTrimLeft, FinalText, FinalText, 6
	
	If !FinalText
	{
		Div += 1
		continue
	}
	
	
		;msgbox %FinalText%
	
	
	If (PrevText = FinalText)
	{
		Div += 1
		continue
	}
	else
	{
		LinkDeposit.push(FinalText)
	}
	
	
	PrevText := FinalText
	
	Div += DivIterator
	
	PriceText := IEInterface.Document.GetElementsByTagName("SPAN")[PriceDiv].InnerText
	
		;msgbox %PriceText%
	
	PriceArray.Push(PriceText)
	
	PriceDiv += PriceDivIterator
	
}


;TestLength := LinkDeposit.length()

;msgbox, 0, Hi, % LinkDeposit.length()

;msgbox, 0, Hi, % LinkDeposit[TestLength]





	
	BrandArray := Object()
	NameArray := Object()
	
	SizeArray := Object()
	RatingArray := Object()
	DetailsArray := Object()
	IngredientsArray := Object()
	
	
	
	Loop, % LinkDeposit.length()
	{
		
		if (a_index < OverrideVar)
		{
			row += 1
			continue
		}
		
		CompURL := BaseUrl . LinkDeposit[a_index]
		
		;CompURL := "https://www.sephora.com/product/double-wear-stay-in-place-makeup-P378284?icid2=:p378284:product"
		
		IEInterface.Navigate(CompURL)
		
		IELOAD(IEInterface)
		
		sleep, 4000
		
		Try
			BrandText := IEInterface.Document.GetElementsByTagName("SPAN")[10].InnerText
		
		Catch
		{
			BrandText := "Error"
		}
		;msgbox %BrandText%
		
		BrandArray.Push(BrandText)
		
		try
			NameText := IEInterface.Document.GetElementsByTagName("SPAN")[11].InnerText
		
		Catch
		{
			NameText := "Error"
		}
		
		;msgbox %NameText%
		
		NameArray.Push(NameText)
		
		try
			SizeText := IEInterface.Document.GetElementsByTagName("SPAN")[12].InnerText
		
		Catch
		{
			SizeText := "Error"
		}
		
		;msgbox %SizeText%
		
		IfNotInString, SizeText, •
		{
			SizeText := "N/A"
		}
		else
		{
			StringTrimRight, SizeText, SizeText, 1
		}
		
		SizeArray.Push(SizeText)
		
		Loop
		{
			if (a_index < 900)
				continue
			
			Try
				RatingText := IEInterface.Document.GetElementsByTagName("DIV")[a_index].InnerText
			
			Catch
			{
				RatingText := "Error"
				Break
			}
			
			IfInString, RatingText, stars
			{
				If (StrLen(RatingText) > 20)
				{
					continue
				}
				
				;msgbox %RatingText%
				
				RatingArray.Push(RatingText)
				
				break
			}
			else
			{
				sleep, 50
			}
			
		}
		
		
		
		
		Loop, 2000
		{
			if (a_index < 600)
				continue
			
			
			try
				DetailsText := IEInterface.Document.GetElementsByTagName("DIV")[a_index].OuterHTML
			Catch
			{
				DetailsText := "Error"
				IngredientsText := "Error"
				break
			}
			
			IfInString, DetailsText,  css-8l83pg
			{
				
			;msgbox %a_index%
				
			;msgbox %DetailsText%
				
				IfInString, DetailsText, GridCell Box
				{
					continue
				}
				
				Loop
				{
					if (a_index < 15)
						continue
					
					Breaker := 0
					
					ButtonText := IEInterface.Document.GetElementsByTagName("Button")[a_index].OuterHTML
					
					IfInString, ButtonText, Details
					{
						BaseNumber := a_index
					}
					
					IfInString, ButtonText, Ingredients
					{
						EndNumber := a_index
						
						break
					}
					
					If a_index = 100
					{
						breaker := 1
						break
					}
					
				}
				
				
				If Breaker = 0
				{
				;msgbox %shifter%
					
					DetailsText := IEInterface.Document.GetElementsByTagName("DIV")[a_index + 2].InnerText
					
					
					Shifter := (2 * (EndNumber - BaseNumber)) + 2
					
					
					IngredientsText := IEInterface.Document.GetElementsByTagName("DIV")[a_index + Shifter].InnerText
				}
				else
				{
					DetailsText := IEInterface.Document.GetElementsByTagName("DIV")[a_index + 2].InnerText
					
					IngredientsText := "N/A"
				}
				
				
				
				
				
				
				
				If (StrLen(DetailsText) > 20000 || StrLen(DetailsText) < 50 )
				{
					continue
				}
				
				break
				
				
				
				
			}
			else
			{
				sleep, 50
			}
			
		}
		
		
		
		wb.SHEETS("Sheet1").CELLS(row,col).VALUE := BrandText
		col += 1
		
		wb.SHEETS("Sheet1").CELLS(row,col).VALUE := NameText
		col += 1
		
		wb.SHEETS("Sheet1").CELLS(row,col).VALUE := PriceArray[a_index]
		col += 1
		
		wb.SHEETS("Sheet1").CELLS(row,col).VALUE := SizeText
		col += 1
		
		wb.SHEETS("Sheet1").CELLS(row,col).VALUE := RatingText
		col += 1
		
		wb.SHEETS("Sheet1").CELLS(row,col).VALUE := DetailsText
		col += 1
		
		wb.SHEETS("Sheet1").CELLS(row,col).VALUE := IngredientsText
		col += 1
		
		col := 1
		
		row += 1
		
	}
	
	




Loop, % LinkDeposit.length()
{
	
	if (a_index < OverrideVar)
	{
		row += 1
		continue
	}
	
	CompURL := BaseUrl . LinkDeposit[a_index]
	
	IEInterface.Navigate(CompURL)
	
	IELOAD(IEInterface)
	
	sleep, 1000
	
	imageDiv := 400
	
	
	try
		NameText := IEInterface.Document.GetElementsByTagName("SPAN")[11].InnerText
	
	Catch
	{
		NameText := "Error"
	}
	
	col := 2
	
	wb.SHEETS("Sheet1").CELLS(row,col).VALUE := NameText
	
	col := 8
	
	Loop, 300
	{
		
		ImageInfo := IEInterface.Document.GetElementsByTagName("DIV")[imageDiv].OuterHTML
		
		IfInString, ImageInfo, svg xmlns
		{
			IfInString, ImageInfo, http://www.w3.org/2000/svg
			{
				IfInString, ImageInfo, href=
				{
					IfInString, ImageInfo, src=
					{
						IfInString, ImageInfo, jpg
						{
							IfNotInString, ImageInfo, makeup-cosmetics
							{
								IfInString, ImageInfo, main-Lhero
								{
									
									IfNotInString, ImageInfo, skincare
									{
										
										RegExMatch(ImageInfo, "href=.+?\s", FinalImageInfo)
										
										StringTrimLeft, FinalImageInfo, FinalImageInfo, 6
										StringTrimRight, FinalImageInfo, FinalImageInfo, 2
										
										if !FinalImageInfo
										{
											imageDiv += 1
											continue
										}
										
										CompURL := BaseUrl . FinalImageInfo
										
										wb.SHEETS("Sheet1").CELLS(row,col).VALUE := CompURL
										
										break
									}
									
								}
							}
						}
					}
				}
			}
		}
		
		
		
		
		imageDiv += 1
		
	}
	
	row += 1
}







Wb.Saveas(FILEPATH)

Wb.quit()

ExcelInterface := ""

IEInterface.Quit()

IEInterface := ""





exitapp






IELoad(IEInstance)    ;Wait till website is loaded - from AHK forums
{
	If !IEInstance    
		Return False
	Loop    
		Sleep,100
	Until (IEInstance.busy)
	Loop    
		Sleep,100
	Until (!IEInstance.busy)
	Loop    
		Sleep,100
	Until (IEInstance.Document.Readystate = "Complete")
	Return True
}
















F12::

IEInterface.quit()

IEInterface := ""

Wb.Saveas(FILEPATH)

Wb.quit()

ExcelInterface := ""


ExitApp

