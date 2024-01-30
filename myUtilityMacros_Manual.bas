Attribute VB_Name = "myUtilityMacros_Manual"
Option Explicit
Sub u_Hide_All_Worksheets_Except_ActiveSheet()

' Purpose: To hide all of the sheets in a workbook.
' Trigger: Manual
' Updated: 11/14/2019

' ***********************************************************************************************************************************

Dim ws As Worksheet

For Each ws In ActiveWorkbook.Worksheets
    If ws.Name <> ActiveSheet.Name Then ws.Visible = xlSheetHidden
Next ws

End Sub
Sub u_Unhide_All_Columns_n_Rows_in_Workbook()

' Purpose: To unhide all of the rows / columns in each sheet in a workbook.
' Trigger: Manual
' Updated: 11/14/2019

' ***********************************************************************************************************************************

Dim ws As Worksheet

For Each ws In ActiveWorkbook.Worksheets
    ws.Columns.EntireColumn.Hidden = False
    ws.Rows.EntireRow.Hidden = False
Next ws

End Sub
Sub u_Resize_Visible_Columns()
    
' Purpose: To autosize all of the columns in the ActiveSheet.
' Trigger: Manual
' Updated: 9/16/2019

' ***********************************************************************************************************************************
    
    ActiveWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeVisible).Columns.AutoFit
    
End Sub
Sub u_Hide_Standard_Columns_With_Grey_Headers()
    
' Purpose: To hide all of the columns that have the standard grey fill.
' Trigger: Manual
' Updated: 7/7/2022

' Change Log:
'       7/6/2022:   Initial Creation
'       7/7/2022:   Added the ability to toggle between visible and not visible

' ***********************************************************************************************************************************
    
Application.ScreenUpdating = False
    
' -----------------
' Declare Variables
' -----------------
    
    Dim wsTarget As Worksheet
    Set wsTarget = ActiveWorkbook.ActiveSheet
    
    Dim clrStandardGrey As Long
        clrStandardGrey = RGB(230, 230, 230)
        
    Dim clrStandardWhite As Long
        clrStandardWhite = RGB(255, 255, 255)
        
    Dim i As Integer
        i = 1
    
' ----------------
' Hide the Columns
' ----------------
    
    Do Until IsEmpty(wsTarget.Cells(1, i).Value2)
        
        If wsTarget.Cells(1, i).Interior.Color = clrStandardGrey Or wsTarget.Cells(1, i).Interior.Color = clrStandardWhite Then
            If wsTarget.Columns(i).Hidden = False Then
                wsTarget.Columns(i).Hidden = True
            Else
                wsTarget.Columns(i).Hidden = False
            End If
        End If
        
        i = i + 1
    Loop
       
Application.ScreenUpdating = True
       
End Sub
Sub u_VBA_Bible_Table_Formatting()

' Purpose: To format the table the way I like in the VBA Bible.
' Trigger: Manual
' Updated: 1/7/2021

' Change Log:
'       9/16/2019: Initial Creation
'       1/7/2021: Added the AutoFit code

' ***********************************************************************************************************************************

Call myPrivateMacros.DisableForEfficiency

' -----------------
' Declare Variables
' -----------------
    
    Dim Table_rg As Range
        Set Table_rg = Selection
    
    Dim LastRow As Long
        LastRow = ActiveSheet.Cells(Rows.count, Table_rg.Column).End(xlUp).Row
    
    Dim LastCol As Long
        LastCol = ActiveSheet.Cells(LastRow, ActiveSheet.Columns.count).End(xlToLeft).Column
    
    Dim Rows_cnt
        Rows_cnt = Table_rg.Rows.count
    
    Dim Columns_cnt
        Columns_cnt = Table_rg.Columns.count
    
    Dim Title_row As Range
        Set Title_row = Range(Cells(Table_rg.Row, Table_rg.Column), Cells(Table_rg.Row, Table_rg.Column + Columns_cnt - 1))
        
    Dim Header_row As Range
        Set Header_row = Title_row.Offset(1, 0)
        
    Dim x As Long

' -----------
' Apply the formatting
' -----------
    
    With Title_row
        .Merge

        .Font.Bold = True
        .Font.Size = 14
                   
    End With

    With Header_row
        .Font.Bold = True
        .Font.Size = 12

        .Borders(xlEdgeBottom).Color = RGB(190, 190, 190)
        
        .Interior.Color = RGB(242, 242, 242)
            
    End With
    
    For x = Header_row.Row + 1 To LastRow
        With Range(Cells(x, Table_rg.Column), Cells(x, LastCol))
            If x Mod 2 = 0 Then .Interior.Color = RGB(240, 240, 240)
            .Borders(xlEdgeBottom).Color = RGB(190, 190, 190)
            .Borders(xlEdgeTop).Color = RGB(190, 190, 190)
        End With
    Next x

    With Table_rg
        .Font.Name = "Cambria"
        
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).Color = vbBlack ' Make this black instead of grey
    
    End With

    ActiveSheet.Cells.EntireColumn.AutoFit

Call myPrivateMacros.DisableForEfficiencyOff

End Sub
Sub u_Add_OLD_All_Files()

' Purpose: To add the word [OLD] to each file in a folder.
' Trigger: Manual
' Updated: 2/6/2020

' ***********************************************************************************************************************************

' -----------------
' Declare Variables
' -----------------

    Dim objFSO As Object
        Set objFSO = CreateObject("Scripting.FileSystemObject")
        
    Dim strDir As String
        strDir = objFSO.GetParentFolderName(Application.GetOpenFilename(Title:="Select an example file to set the folder")) & "\"

    Dim strFileExt As String
                
    Dim strCurFile As String
        strCurFile = Dir(strDir & "*.*")

' -----------
' Find the latest itteration
' -----------
    
    Do While strCurFile <> ""
        
        strFileExt = Mid(strCurFile, InStrRev(strCurFile, "."))
        
        Debug.Print objFSO.GetFile(strCurFile).Name & " - " & Format(objFSO.GetFile(strCurFile).DateLastModified, "MM/DD/YYYY")
               
        objFSO.GetFile(strDir & strCurFile).Name = Left(strCurFile, Len(strCurFile) - Len(strFileExt)) & " [OLD]" & strFileExt
        
        strCurFile = Dir()

    Loop

End Sub
Sub u_Update_Hyperlinks()
    
' Purpose: To update all of the hyperlinks on the Active Sheet.
' Trigger: Manual
' Updated: 2/21/2020

' ***********************************************************************************************************************************
    
' -----------------
' Declare Variables
' -----------------
    
    Dim ws As Worksheet
        Set ws = ActiveWorkbook.ActiveSheet
    
    Dim hl As Hyperlink
    
    Dim strOldLink As String
        strOldLink = InputBox("What is the part of the hyperlink you want to update?")
    
    Dim strNewLink As String
        strNewLink = InputBox("What should that part of the hyperlink now be?")

' -----------
' Fix the Hyperlinks
' -----------

    For Each hl In ws.Hyperlinks
        
        Debug.Print "Old Address: "; hl.Address
        
        hl.Address = Replace(hl.Address, strOldLink, strNewLink)
        
        Debug.Print "New Address: "; hl.Address
        
    Next hl
    
End Sub
Sub u_Remove_Custom_Styles()

' Purpose: To remove all custom styles from the active workbook to fix the "To Many Formats" error.
' Trigger: Manual
' Updated: 9/17/2019
' Note: This one can take a few minutes

' ***********************************************************************************************************************************
 
Call myPrivateMacros.DisableForEfficiency
 
On Error GoTo ErrorHandler
 
' -----------------
' Declare Variables
' -----------------
 
    Dim tmpSt As Style
    
    Dim wb As Workbook
    
    Dim wkb As Workbook
    
    Set wkb = ActiveWorkbook
        
' -----------
' Run your code
' -----------
  
    For Each tmpSt In wkb.Styles
        With tmpSt
            If .BuiltIn = False Then
                .Locked = False
                .Delete
            End If
        End With
    Next tmpSt
 
ErrorHandler:
    Set tmpSt = Nothing
    Set wkb = Nothing
 
Call myPrivateMacros.DisableForEfficiencyOff
 
End Sub
