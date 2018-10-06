#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.




ExcelInterface := ComObjCreate("Excel.Application")

FILEPATH := "C:\Users\Christian\Desktop\Results.XLS"

;WORKBOOK := ExcelInterface.WORKBOOKS.OPEN(FILEPATH,, readonly := true)

wb:=ExcelInterface.Workbooks.Add(),wb.SaveAs(FILEPATH)

ExcelInterface.Visible := True

row := 1

col := 1


wb.SHEETS("Sheet1").CELLS(row,col).VALUE := "Ingredients"

col += 1

wb.SHEETS("Sheet1").CELLS(row,col).VALUE := "Description"

col += 1

wb.SHEETS("Sheet1").CELLS(row,col).VALUE := "Categories"

col += 1
S
wb.SHEETS("Sheet1").CELLS(row,col).VALUE := "Fragrances"

col += 1

wb.SHEETS("Sheet1").CELLS(row,col).VALUE := "Rating"

col := 1

row := 2

IEInterface := ComObjCreate("InternetExplorer.Application")

URL := "https://www.paulaschoice.com/ingredient-dictionary"

IEInterface.Navigate(URL)

IELOAD(IEInterface)

IEInterface.Visible := True

sleep, 1000


ratingIt := 0

ingredientTitleIt := 150

descIt := 5

categoryIt := 49

TDIt := 1

/*
	ratingIt := 2532
	
	ingredientTitleIt := 3731
	
	descIt := 974
	
	categoryIt := 1315
	
	var := ""
	
	totalMessage := ""
*/

Loop,
{
	var := IEInterface.Document.GetElementsByTagName("A")[ingredientTitleIt].Innertext
	
	blankTester := ""
	
	If !Var
	{
		break
	}
	
	totalMessage .= "Title: " var "`r`n"
	
	wb.SHEETS("Sheet1").CELLS(row,col).VALUE := var
	
	col += 1
	
	blankTester := IEInterface.Document.GetElementsByTagName("TD")[TDIt].Innertext
	
	;If TDIt = 99
	;{
	;	MSGBOX %blankTester%
	;	Clipboard := blankTester
	;}
	
	;MSGBOX %blankTester%
	
	
	If InStr(blankTester, "…") or InStr(blankTester, ".") or InStr(blankTester, "quercetin") or InStr(blankTester, "hibiscus") or InStr(blankTester, "demonstrated") or InStr(blankTester, "acid attached to a polymer structure") or InStr(blankTester, "showing vitamin B12")
	{
		
		var := IEInterface.Document.GetElementsByTagName("P")[descIt].Innertext
		
		Loop,
		{
			If !var 
			{
				descIt += 1
				var := IEInterface.Document.GetElementsByTagName("P")[descIt].Innertext
				
			;msgbox %var%
			}
			else
			{
				break
			}
			
		}
		
		descIt += 1
		
	}
	else
	{
		;If TDIt = 99
		;{
		;	msgbox, 0, Test!, No period or ellipse detected
		;}
		
		var := "None"	
	}
	
	
	
	totalMessage .= "Desc: " var "`r`n"
	
	wb.SHEETS("Sheet1").CELLS(row,col).VALUE := var
	
	col += 1
	
	var := IEInterface.Document.GetElementsByTagName("DIV")[categoryIt].Innertext
	
	StringReplace, var, var, Fragrance:%a_space% , ~
	
	StringSplit, FragrancePiece, var, ~
	
	;Msgbox %FragrancePiece0%
	
	StringTrimLeft, FragrancePiece1,FragrancePiece1, 12
	;StringTrimLeft, FragrancePiece2,FragrancePiece2, 11
	
	If FragrancePiece0 = 2
	{
		StringTrimRight, FragrancePiece1, FragrancePiece1, 2
		
		totalMessage .= "Categories: " FragrancePiece1 "`r`n"
		totalMessage .= "Fragrance: " FragrancePiece2 "`r`n"
		
		
	}
	else
	{
		totalMessage .= "Categories: " FragrancePiece1 "`r`n"
		totalMessage .= "Fragrance: N/A`r`n"
	}
	
	Stringsplit, CategoryPiece, FragrancePiece1, `,
	
	wb.SHEETS("Sheet1").CELLS(row,col).VALUE := FragrancePiece1
	
	col += 1
	
	If FragrancePiece0 = 2
	{
		wb.SHEETS("Sheet1").CELLS(row,col).VALUE := FragrancePiece2
	}
	else
	{
		wb.SHEETS("Sheet1").CELLS(row,col).VALUE := "N/A"
	}
	
	col += 1
	
	var := IEInterface.Document.GetElementsByTagName("TD")[ratingIt].Innertext
	
	totalMessage .= "Rating: " var
	
	wb.SHEETS("Sheet1").CELLS(row,col).VALUE := var
	
	col := 1
	
	;Msgbox, 0, Test!, %totalMessage%
	
	If FragrancePiece0 = 2
	{
		
		StringSplit, FragSplit, FragrancePiece2, `,
		
	}
	else
	{
		FragSplit0 := 0
	}
	
	
	
	;Msgbox, %CategoryPiece0% , %FragSplit0%
	
	ratingIt += 2
	
	categoryIt += 1
	
	
	
	ingredientTitleIt += 1 + CategoryPiece0 + FragSplit0
	
	TDIt += 2
	
	totalMessage := ""
	
	FragSplit0 := ""
	CategoryPiece0 := ""
	
	row += 1
	
	
}



Wb.Saveas(FILEPATH)

Wb.quit()

ExcelInterface := ""


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