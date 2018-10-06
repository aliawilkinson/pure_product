#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.


ExcelInterface := ComObjCreate("Excel.Application")

;FILEPATH := "C:\Users\Christian\Desktop\FoundationResults.XLS"

FILEPATH := "C:\Users\Christian\Desktop\Ingredienttest.XLSX"

;WORKBOOK := ExcelInterface.WORKBOOKS.OPEN(FILEPATH,, readonly := true)

;wb:=ExcelInterface.Workbooks.Add(),wb.SaveAs(FILEPATH)

wb:=ExcelInterface.Workbooks.Open(FILEPATH)

ExcelInterface.Visible := True

row := 2

col := 1


IngredientNameTotalArray := Object()

Loop
{
	
	Var := wb.SHEETS("Sheet1").CELLS(row,col).VALUE
	
	If !Var
		Break
	
	IngredientNameTotalArray.Push(var)
	
	row += 1
	
	
}

row := 2

Col := 2

IEInterface := ComObjCreate("InternetExplorer.Application")

IEInterface.Visible := True


URL := "https://www.ewg.org/skindeep/"

BaseURL := "https://www.ewg.org"

IEInterface.Navigate(URL)

IELOAD(IEINTERFACE)
;sleep, 4000

;sleep, 2500

Loop % IngredientNameTotalArray.Length()
{
	
	
	
	
	
	ScoreData := "None"
	AvailabilityData := "None"
	HighConcern := "None"
	MidConcern := "None"
	LowConcern := "None"
	AboutData := "None"
	FunctionData := "None"
	SynonymData := "None"
	CompURL := "None"
	Citation := "None"
	
	ScoreData0 := ""
	ScoreData1 := ""
	ScoreData2 := ""
	ScoreData3 := ""
	ScoreData4 := ""
	
	
	Col := 2
	
	
	
	if (wb.SHEETS("Sheet1").CELLS(row,col).VALUE)
	{
		
		row += 1
		
		continue
	}
	
	
	IEInterface.Document.GetElementByID("s").Value := IngredientNameTotalArray[a_index]
	
	IEInterface.Document.GetElementByID("gobtn").click()
	
	IELOAD(IEINTERFACE)
	
	;sleep, 4000
	;sleep, 2500
	
	Try
		FoundCheck := IEInterface.Document.GetElementsByTagName("P")[1].InnerText
	
	Ifinstring, FoundCheck, that match your request
	{
		
		
		
		ScoreData := "Not Found!"
		AvailabilityData := "Not Found!"
		HighConcern := "Not Found!"
		MidConcern := "Not Found!"
		LowConcern := "Not Found!"
		AboutData := "Not Found!"
		FunctionData := "Not Found!"
		SynonymData := "Not Found!"
		CompURL := "Not Found!"
		Citation := "Not Found!"
		
		
		wb.SHEETS("Sheet1").CELLS(row,col).VALUE := ScoreData
		col += 1
		wb.SHEETS("Sheet1").CELLS(row,col).VALUE := AvailabilityData
		col += 1
		wb.SHEETS("Sheet1").CELLS(row,col).VALUE := HighConcern
		col += 1
		wb.SHEETS("Sheet1").CELLS(row,col).VALUE := MidConcern
		col += 1
		wb.SHEETS("Sheet1").CELLS(row,col).VALUE := LowConcern
		col += 1
		wb.SHEETS("Sheet1").CELLS(row,col).VALUE := AboutData
		col += 1
		wb.SHEETS("Sheet1").CELLS(row,col).VALUE := FunctionData
		col += 1
		wb.SHEETS("Sheet1").CELLS(row,col).VALUE := SynonymData
		col += 1
		wb.SHEETS("Sheet1").CELLS(row,col).VALUE := CompURL
		col += 1
		wb.SHEETS("Sheet1").CELLS(row,col).VALUE := Citation
		
		IEInterface.Navigate(URL)
		
		IELOAD(IEINTERFACE)
		;sleep, 4000
		sleep, 2500
		row += 1
		
		FoundCheck := ""
		
		continue
	}
	
	
	Data := IEInterface.Document.GetElementsByTagName("TD")[11].OuterHTML
	
	IfNotInString, Data, Ingredient:
	{
		ScoreData := "Not Found!"
		AvailabilityData := "Not Found!"
		HighConcern := "Not Found!"
		MidConcern := "Not Found!"
		LowConcern := "Not Found!"
		AboutData := "Not Found!"
		FunctionData := "Not Found!"
		SynonymData := "Not Found!"
		CompURL := "Not Found!"
		Citation := "Not Found!"
		
		
		wb.SHEETS("Sheet1").CELLS(row,col).VALUE := ScoreData
		col += 1
		wb.SHEETS("Sheet1").CELLS(row,col).VALUE := AvailabilityData
		col += 1
		wb.SHEETS("Sheet1").CELLS(row,col).VALUE := HighConcern
		col += 1
		wb.SHEETS("Sheet1").CELLS(row,col).VALUE := MidConcern
		col += 1
		wb.SHEETS("Sheet1").CELLS(row,col).VALUE := LowConcern
		col += 1
		wb.SHEETS("Sheet1").CELLS(row,col).VALUE := AboutData
		col += 1
		wb.SHEETS("Sheet1").CELLS(row,col).VALUE := FunctionData
		col += 1
		wb.SHEETS("Sheet1").CELLS(row,col).VALUE := SynonymData
		col += 1
		wb.SHEETS("Sheet1").CELLS(row,col).VALUE := CompURL
		col += 1
		wb.SHEETS("Sheet1").CELLS(row,col).VALUE := Citation
		
		IEInterface.Navigate(URL)
		
		IELOAD(IEINTERFACE)
		;sleep, 4000
		sleep, 2500
		row += 1
		
		FoundCheck := ""
		
		continue
	}
	
	Var := RegexMatch(Data, "href.+?>", FormData)
	
	StringTrimLeft, FormData, FormData, 6
	
	StringTrimRight, FormData, FormData, 2
	
	;Msgbox %FormData%
	
	CompURL := BaseURL . FormData
	
	IEInterface.Navigate(CompURL)
	
	
	IELOAD(IEINTERFACE)
	;sleep, 2500
	
	ScoreData := ""
	
	ScoreData := IEInterface.Document.GetElementsByTagName("IMG")[3].OuterHTML
	
	Var := RegexMatch(ScoreData, "\d+_\d+_\d", ScoreData)
	
	
	
	StringSplit, ScoreData, ScoreData, _
	
	
	
	ScoreData1 += 0
	ScoreData2 += 0
	
	If (ScoreData1 = ScoreData2)
	{
		;msgbox Only One
		ScoreData := ScoreData1
	}
	else
	{
		
		;msgbox %ScoreData1%,%ScoreData2%
		;msgbox Range!
		ScoreData := ScoreData2 . "-" . ScoreData1
	}
	
	
	wb.SHEETS("Sheet1").CELLS(row,col).VALUE := ScoreData
	
	
	col += 1
	
	Try
		AvailabilityData := IEInterface.Document.GetElementsByTagName("SPAN")[3].InnerText
	Catch
		AvailabilityData := "Error!"
	
	wb.SHEETS("Sheet1").CELLS(row,col).VALUE := AvailabilityData
	
	
	col += 1
	
	
	;Concerns Section-------------------------
	
	ConcernDataWhole := IEInterface.Document.GetElementsByTagName("DIV")[103].InnerText
	
	ifinstring, ConcernDataWhole, concerns:
	{
		
		;msgbox %ConcernDataWhole%
		
		StringSplit, ConcernData, ConcernDataWhole, `;
		
		HighConcern := "None"
		MidConcern := "None"
		LowConcern := "None"
		
		IfInString, ConcernData1, Other HIGH concerns:
		{
			HighConcern := ConcernData1
		}
		
		IfInString, ConcernData2, Other HIGH concerns:
		{
			HighConcern := ConcernData2
		}
		
		IfInString, ConcernData3, Other HIGH concerns:
		{
			HighConcern := ConcernData3
		}
		
		IfInString, ConcernData1, Other Moderate concerns:
		{
			MidConcern := ConcernData1
		}
		
		IfInString, ConcernData2, Other Moderate concerns:
		{
			MidConcern := ConcernData2
		}
		
		IfInString, ConcernData3, Other Moderate concerns:
		{
			MidConcern := ConcernData3
		}
		
		IfInString, ConcernData1, Other Low concerns:
		{
			LowConcern := ConcernData1
		}
		
		IfInString, ConcernData2, Other Low concerns:
		{
			LowConcern := ConcernData2
		}
		
		IfInString, ConcernData3, Other Low concerns:
		{
			LowConcern := ConcernData3
		}
		
		HighConcern := RegexReplace(HighConcern, "Other HIGH concerns: ", "")
		
		MidConcern := RegexReplace(MidConcern, "Other MODERATE concerns: ", "")
		
		LowConcern := RegexReplace(LowConcern, "Other LOW concerns: ", "")
		
		wb.SHEETS("Sheet1").CELLS(row,col).VALUE := HighConcern
		
		col += 1
		
		wb.SHEETS("Sheet1").CELLS(row,col).VALUE := MidConcern
		
		col += 1
		
		wb.SHEETS("Sheet1").CELLS(row,col).VALUE := LowConcern
		
		
		
		
		
		col += 1
		
	}
	else
	{
		wb.SHEETS("Sheet1").CELLS(row,col).VALUE := HighConcern
		
		col += 1
		
		wb.SHEETS("Sheet1").CELLS(row,col).VALUE := MidConcern
		
		col += 1
		
		wb.SHEETS("Sheet1").CELLS(row,col).VALUE := LowConcern
		
		
		
		
		
		col += 1
	}
	
	;------------------------------------------
	
	
	AboutData := "None"
	FunctionData := "None"
	SynonymData := "None"
	
	
	Try
	{
		Data := IEInterface.Document.GetElementsByTagName("P")[1].InnerText
		
		;msgbox %Data% 
		
		IfInString, Data, About
		{
			AboutData := Data
		}
		else
		{
			ifinstring, Data, Function(s):
			{
				FunctionData := Data
			}
			else
			{
				ifinstring, Data, Synonym(s):
				{
					SynonymData := Data
				}
			}
		}
	}
	
	Try
	{
		Data := IEInterface.Document.GetElementsByTagName("P")[2].InnerText
		;msgbox %Data% 
		IfInString, Data, About
		{
			AboutData := Data
		}
		else
		{
			ifinstring, Data, Function(s):
			{
				FunctionData := Data
			}
			else
			{
				ifinstring, Data, Synonym(s):
				{
					SynonymData := Data
				}
			}
		}
	}
	
	Try
	{
		Data := IEInterface.Document.GetElementsByTagName("P")[3].InnerText
		;msgbox %Data% 
		IfInString, Data, About
		{
			AboutData := Data
		}
		else
		{
			ifinstring, Data, Function(s):
			{
				FunctionData := Data
			}
			else
			{
				ifinstring, Data, Synonym(s):
				{
					SynonymData := Data
				}
			}
		}
	}
	
	Try
	{
		Data := IEInterface.Document.GetElementsByTagName("P")[4].InnerText
		;msgbox %Data% 
		IfInString, Data, About
		{
			AboutData := Data
		}
		else
		{
			ifinstring, Data, Function(s):
			{
				FunctionData := Data
			}
			else
			{
				ifinstring, Data, Synonym(s):
				{
					SynonymData := Data
				}
			}
		}
	}
	
	Try
	{
		Data := IEInterface.Document.GetElementsByTagName("P")[5].InnerText
		;msgbox %Data% 
		IfInString, Data, About
		{
			AboutData := Data
		}
		else
		{
			ifinstring, Data, Function(s):
			{
				FunctionData := Data
			}
			else
			{
				ifinstring, Data, Synonym(s):
				{
					SynonymData := Data
				}
			}
		}
	}
	
	
	Try
	{
		Data := IEInterface.Document.GetElementsByTagName("P")[6].InnerText
		;msgbox %Data% 
		IfInString, Data, About
		{
			AboutData := Data
		}
		else
		{
			ifinstring, Data, Function(s):
			{
				FunctionData := Data
			}
			else
			{
				ifinstring, Data, Synonym(s):
				{
					SynonymData := Data
				}
			}
		}
	}
	
	
	;AboutData := IEInterface.Document.GetElementsByTagName("P")[1].InnerText
	
	AboutData := RegexReplace(AboutData, "About.+?:\s", "")
	
	wb.SHEETS("Sheet1").CELLS(row,col).VALUE := AboutData
	
	col += 1
	
	
	;FunctionData := IEInterface.Document.GetElementsByTagName("P")[2].InnerText
	
	FunctionData := RegexReplace(FunctionData, "Function\(s\):\s", "")
	
	wb.SHEETS("Sheet1").CELLS(row,col).VALUE := FunctionData
	
	col += 1
	
	
	;SynonymData := IEInterface.Document.GetElementsByTagName("P")[3].InnerText
	
	SynonymData := RegexReplace(SynonymData, "Synonym\(s\):\s", "")
	
	wb.SHEETS("Sheet1").CELLS(row,col).VALUE := SynonymData
	
	col += 1
	
	wb.SHEETS("Sheet1").CELLS(row,col).VALUE := CompUrl
	
	col += 1
	
	
	Title := IEInterface.Document.GetElementsByTagName("H1")[0].InnerText
	
	Citation := "'" . Title . "' " . "EWG.Accessed " . A_MMMM . A_DD . ", " . A_YYYY . ". " . CompURL
	
	wb.SHEETS("Sheet1").CELLS(row,col).VALUE := Citation
	
	
	
	
	
	row += 1
	
	IEInterface.Navigate(URL)
	
	
	IELOAD(IEINTERFACE)
	;sleep, 4000
	sleep, 2500
}













IELoad(IEInstance)    ;Wait till website is loaded - from AHK forums
{
	
	URL := "https://www.ewg.org/skindeep/"
	
	If !IEInstance    
		Return False
	Loop    
	{
		Sleep,100
		
		if a_index > 50
		{
			
			break
			
			;IEInstance.Navigate(URL)
			
			;sleep, 3000
			
			;return
		}
	}
	Until (IEInstance.busy)
	Loop    
	{
		Sleep,100
		
		if a_index > 50
		{
			
			break
			
			;IEInstance.Navigate(URL)
			
			;sleep, 3000
			
			;return
		}
	}
	Until (!IEInstance.busy)
	Loop    
	{
		Sleep,100
		
		if a_index > 50
		{
			
			break
			
			;IEInstance.Navigate(URL)
			
			;sleep, 3000
			
			;return
		}
		
	}
	Until (IEInstance.Document.Readystate = "Complete")
	Return True
}




F12::


ExitApp
