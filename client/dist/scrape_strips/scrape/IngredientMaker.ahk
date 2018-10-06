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

Loop
{
	
	RawData := ""
	
;InputBox, RawData, Input String
	
	Gui, Input: New, , Test
	
	Gui, Input: Add, Edit, r9 vRawData Multi
	
	Gui, Input: Add, Button, Default gButton, OK
	
	Gui, Input: Show
	
	winwaitclose, Test
	
	Result := object()
	
	Gui, Input: Destroy
	
	
	
	
	MetaData := RegexReplace(RawData, "`n",",")
	;msgbox %metadata%
	
	MetaData := RegexReplace(MetaData, "\\",",")
	;msgbox %metadata%
	
	MetaData := RegexReplace(MetaData, ";","")
	;msgbox %metadata%
	
	MetaData := RegexReplace(MetaData, "(\s\(.+?\))","")
	;msgbox %metadata%
	
	MetaData := RegexReplace(MetaData, "\.","")
	;msgbox %metadata%
	
	MetaData := RegexReplace(MetaData, "/",",")
	;msgbox %metadata%
	
	Result := StrSplit(MetaData, ",")
	
	;msgbox % result.length()
	
	Loop, % result.length()
	{
		OverIt := a_index
		
		matchfound := 0
		
		var := Result[OverIt]
		
		Loop
		{
			StringRight, var1, var , 1
			
			if (Var1 = A_space)
			{
				StringTrimRight, var,var, 1
			}
			else
			{
				break
			}
		}
		
		Loop
		{
			StringLeft, var1, var , 1
			
			;msgbox %var1%
			
			if (Var1 = A_space)
			{
				StringTrimLeft, var,var, 1
			}
			else
			{
				break
			}
		}
		
		;msgbox % IngredientNameTotalArray.Length()
		
		Loop, % IngredientNameTotalArray.Length()
		{
			If (var = IngredientNameTotalArray[a_index])
			{
				matchfound := 1
				
				;Msgbox, 0, Duplicate!, Found Duplicate %var%, 0.5
				
				break
			}
		}
		
		
		if (matchfound = 1)
			continue
		
		wb.SHEETS("Sheet1").CELLS(row,col).VALUE := var
		
		IngredientNameTotalArray.Push(var)
		row += 1
	}
	
	
}


exitapp




Button:


gui Input: Submit



f12::

exitapp


