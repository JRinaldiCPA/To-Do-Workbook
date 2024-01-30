Attribute VB_Name = "myUtilityMacros_Keyboard"
Option Explicit
Sub u_Paste_Special_Values()
Attribute u_Paste_Special_Values.VB_ProcData.VB_Invoke_Func = "V\n14"

' Purpose: To copy > paste values only, or paste from the clipboard, in one step.
' Trigger: Keyboard Shortcut - Ctrl + Shift + V
' Updated: 6/24/2021

' Change Log:
'       9/16/2019:  Initial Creation
'       6/24/2021:  Added the code to parse the text and replace the tabs (Chr 9) with spaces.
'                   This allowed the text from NotePad++ to be exported to Excel without breaking across multiple lines.
'       7/1/2021:   Removed the code related to the Selection and the Formula

' ***********************************************************************************************************************************

On Error GoTo ErrorHandler

' -----------------
' Declare Variables
' -----------------

    Dim objDataObj As New MSForms.DataObject
    
    Dim strCopied As String
    
' ---------------------------------------------------------------------------------------------
' If the selection has a value C/P as values only, otherwise paste the value from the clipboard
' ---------------------------------------------------------------------------------------------
    
    If Application.CutCopyMode = xlCopy Then
        Selection.PasteSpecial xlPasteValues
    Else
        objDataObj.GetFromClipboard
        strCopied = Trim(objDataObj.GetText(1))
        strCopied = Replace(strCopied, Chr(9), "   ")
        
        objDataObj.SetText strCopied
            objDataObj.PutInClipboard
            
        ActiveSheet.PasteSpecial Format:="Text"
            
    End If
    
ErrorHandler:
    
End Sub
Sub u_Email_To_Personal()
Attribute u_Email_To_Personal.VB_Description = "Used to ber Ctrl + Shift + J, disabled on 6/20/23"
Attribute u_Email_To_Personal.VB_ProcData.VB_Invoke_Func = " \n14"

' Purpose: To quickly be able to send a message home.
' Trigger: Keyboard Shortcut - Ctrl + Shift + J
' Updated: 9/19/2019

' Change Log:
'       9/19/2019:  Initial Creation
'       7/12/2021:  Updated to punt notes to Body

' ***********************************************************************************************************************************

Call myPrivateMacros.DisableForEfficiency

' -----------------
' Declare Variables
' -----------------

    Dim OutApp As Object
        Set OutApp = CreateObject("Outlook.Application")
    
    Dim OutMail As Object
        Set OutMail = OutApp.CreateItem(olMailItem)

    Dim strSelect As String
        If Selection.Rows.count > 1 Then
            strSelect = myFunctions_ToDo.fx_Copy_to_Clipboard
        ElseIf Selection.Value <> "" Then
            strSelect = Selection.Value
        Else
            strSelect = myFunctions_ToDo.fx_Copy_to_Clipboard
        End If
    
    Dim strSubject As String
        strSubject = Format(Date, "m.d.yy") & ": " & "Notes from Work"
    
    Dim strBody As String
        If strSelect <> "" Then
            strBody = Replace(strSelect, Chr(10) & Chr(10), Chr(10))
        Else
            strBody = Format(Date, "m.d.yy") & ": " & Replace(InputBox("What do you want to send yourself", "Email Subject", strSelect), Chr(10) & Chr(10), " / ")
        End If
    
    If strBody = "" Then GoTo Cancel ' Abort if there is no content passed
    
' ------------------------
' Send the email to myself
' ------------------------
    
    With OutMail
        .To = "JRinaldi925@gmail.com"
        .Subject = strSubject
        .Body = strBody
        .Display
        Application.SendKeys "%s"
    End With
    
    Selection.Value = vbNullString

' -----------------------------
' Release the Outlook variables
' -----------------------------

Cancel:

    Set OutMail = Nothing
    Set OutApp = Nothing

Call myPrivateMacros.DisableForEfficiencyOff

End Sub
Sub u_Delete_Cur_Row()
Attribute u_Delete_Cur_Row.VB_ProcData.VB_Invoke_Func = "Y\n14"

' Purpose: To delete the current row with a similar shortcut as VBA.
' Trigger: Keyboard Shortcut - Ctrl + Shift + Y
' Updated: 5/3/2022
'
' Change Log:
'       5/4/2020:   Initial Creation
'       5/3/2022:   Removed the 'ActiveCell.Select'...JK
'
' ***********************************************************************************************************************************

    Selection.EntireRow.Delete
        ActiveCell.Select   ' Want to keep this in if you are deleting multiple cells

End Sub
Sub u_Insert_Row()
Attribute u_Insert_Row.VB_ProcData.VB_Invoke_Func = "I\n14"

' Purpose: To insert a row in the current location.
' Trigger: Keyboard Shortcut - Ctrl + Shift + I
' Updated: 6/24/2021

' Change Log:
'       6/24/2021:  Initial Creation
'       6/24/2021:  Updated to handle if I select the row to insert
'       6/24/2021:  Updated the strCurValue to loop through and aggregate the data.
'       6/24/2021:  Updated to remove formatting from the new line
'       6/25/2021:  Added code for my Daily ws to be smart enough to JUST pull the bullet if I am inserting on a bulleted row

' ***********************************************************************************************************************************

myPrivateMacros.DisableForEfficiency

' -----------------
' Declare Variables
' -----------------

    Dim rngSelection As Range
    Set rngSelection = Selection

    Dim cell As Range

    Dim strCurValue As String
        If rngSelection.Columns.count = 1 And rngSelection.Rows.count < 50 Then
            For Each cell In rngSelection
                If strCurValue <> "" Then
                    strCurValue = strCurValue & Chr(10) & cell.Value2
                Else
                    strCurValue = cell.Value2
                End If
            Next cell
        End If

' -----------------------------------------
' Just copy the bullet if ws_Daily is active
' -----------------------------------------

    If ActiveSheet.Name = "Daily" Or ActiveSheet.Name = "Temp" Then
        If Left(strCurValue, 2) = "• " Or Left(strCurValue, 2) = "¤ " Then
            strCurValue = Left(strCurValue, 2)
        ElseIf Left(strCurValue, 3) = "(?)" Then
            strCurValue = Left(strCurValue, 3)
        End If
    End If

' --------------
' Insert the row
' --------------

    rngSelection.Resize(1, 1).EntireRow.Insert
    
    ' Mirror the values if the selected row is not blank

    If strCurValue <> "" Then
        rngSelection.Resize(1, 1).Offset(-1, 0).Value2 = strCurValue
    End If

    ' Select the newly inserted row
    rngSelection.Resize(1, 1).Offset(-1, 0).Select
        Selection.Font.Bold = False
        Selection.Font.Underline = False
        Selection.Font.Italic = False

myPrivateMacros.DisableForEfficiencyOff

End Sub
Sub u_Steal_Formatting_Row_Above()

' Purpose: To take the formatting from the row directly above the selected row.
' Trigger: Called: o_71_Dynamic_Macro_Splitter
' Updated: 11/5/2020

' Change Log:
'          01/09/2020:  Added the code to prevent overwtitting existing data
'          01/16/2020:  Moved the code to prevent overwrite to the splitter
'          09/18/2020:  Added in the code to apply the row height
'          11/05/2020:  Updated to pull the validations from the row above

' ***********************************************************************************************************************************
  
' -----------------
' Declare Variables
' -----------------

    Dim curRow As Long
        curRow = ActiveCell.Row
        
    Dim rowTarget As Long
        rowTarget = Cells(curRow, ActiveCell.Column).End(xlUp).Row
        If rowTarget = 1 Then rowTarget = curRow - 1

    Dim FirstCol As Long
        FirstCol = Range("1:1").Find("*").Column
        
    Dim LastCol As Long
        LastCol = Cells(1, ActiveSheet.Columns.count).End(xlToLeft).Column

    Dim rngFormat As Range
        Set rngFormat = Range(Cells(rowTarget, FirstCol), Cells(rowTarget, LastCol))
    
    Dim rngTarget As Range
        Set rngTarget = Range(Cells(curRow, FirstCol), Cells(curRow, LastCol))
    
' ---------------------------------------
' Copy the formatting to the target range
' ---------------------------------------
    
    rngFormat.Copy: rngTarget.PasteSpecial xlPasteFormats
    rngFormat.Copy: rngTarget.PasteSpecial xlPasteValidation
    
    rngTarget.RowHeight = rngFormat.RowHeight
    
    Application.CutCopyMode = False

End Sub
Sub u_Insert_•_Bullet()

' Purpose: To instert a • as a bullet to begin a string of text.
' Trigger: Keyboard Shortcut - Ctrl + =
' Updated: 11/18/2023

' Change Log:
'       11/18/2023: Initial Creation

' ***********************************************************************************************************************************

' -----------------
' Declare Variables
' -----------------

    Dim strBullet As String
    
    
    
' ---------------------------------------------------------------------------------------------
' If the selection has a value C/P as values only, otherwise paste the value from the clipboard
' ---------------------------------------------------------------------------------------------
    
    If Application.CutCopyMode = xlCopy Then
        Selection.PasteSpecial xlPasteValues
    Else
        objDataObj.GetFromClipboard
        strCopied = Trim(objDataObj.GetText(1))
        strCopied = Replace(strCopied, Chr(9), "   ")
        
        objDataObj.SetText strCopied
            objDataObj.PutInClipboard
            
        ActiveSheet.PasteSpecial Format:="Text"
            
    End If
    
ErrorHandler:
    
End Sub
