Attribute VB_Name = "myUtilityMacros_Ribbon"
Option Explicit
Sub u_Delete_Test_Project()

' Purpose: To allow me to quickly delete a project and the associated folder, for testing my code.
' Trigger: Ribbon > GTD Macros > Support > Delete Test Record
' Updated: 5/3/2022

' Change Log:
'       9/16/2019:  Initial Creation
'       5/3/2022:   Removed the reference to the DARequestID database that Axcel created

' ***********************************************************************************************************************************

' -----------------
' Declare Variables
' -----------------
        
    Dim objFSO As Object
        Set objFSO = VBA.CreateObject("Scripting.FileSystemObject")
    
    Dim rg
        Set rg = Range(Cells(Selection.Row, "A"), Cells(Selection.Row, "Z"))

' -----------
' Run your code
' -----------
    
    'If ActiveWorkbook.ActiveSheet.Name = "DA Requests" Then Call DatabaseConnection.DeleteDARequestID((Range("A" & rg.Row).Value))
    
On Error Resume Next
    If rg.Hyperlinks.count > 0 Then objFSO.DeleteFolder rg.Hyperlinks(1).Address
        ActiveWorkbook.ActiveSheet.Rows(rg.Row).Delete

End Sub
Sub u_CloseAllExcept()

' Purpose: To close all of the open Workbooks (including Personal Macro) except for my To Do.
' Trigger: Ribbon > GTD Macros > Reset > Close All wbs Except To Do
' Updated: 9/16/2019

' ***********************************************************************************************************************************

Call myPrivateMacros.DisableForEfficiency

    Dim wb As Workbook
    
        For Each wb In Workbooks
            If wb.Name <> "To Do.xlsm" Then
               wb.Close SaveChanges:=True
            End If
        Next wb

Call myPrivateMacros.DisableForEfficiencyOff

End Sub
Sub u_Apply_Table_Formatting_to_Selection()

' Purpose: To format the table the way I like, including a blue header, banded grey rows, and light grey borders.
' Trigger: Ribbon > Personal Macros > Data Formatting > Table Formatting
' Updated: 11/10/2023

' Change Log:
'       9/16/2019:  Initial Creation
'       6/24/2021:  Updated to make the Blue my Blue2, and moved to a variable
'       5/3/2022:   Combined with the code from 'u_Report_Manipulation'
'       11/10/2023: Updated so if only one cell is selected it uses the .CurrentRegion

' ***********************************************************************************************************************************

Call myPrivateMacros.DisableForEfficiency

' -----------------
' Declare Variables
' -----------------
    
    Dim Table_rg As Range
        Set Table_rg = Selection
    
        If Table_rg.Cells.count = 1 Then
            Set Table_rg = Table_rg.CurrentRegion
        End If
    
    Dim LastRow As Long
        LastRow = Table_rg.Rows.count + Table_rg.Row - 1
    
    Dim LastCol As Long
        LastCol = Table_rg.Columns.count + Table_rg.Column - 1
    
    Dim Rows_cnt
        Rows_cnt = Table_rg.Rows.count
    
    Dim Columns_cnt
        Columns_cnt = Table_rg.Columns.count
    
    Dim TitleRow As Long
        TitleRow = Table_rg.Row
        
    Dim x As Long
    
    Dim y As Long
        y = 1
    
    Dim clrGrey1 As Long
        clrGrey1 = RGB(230, 230, 230)
    
' --------------------
' Apply the formatting
' --------------------
    
    With Range(Cells(Table_rg.Row, Table_rg.Column), Cells(Table_rg.Row, LastCol))
        .Font.Bold = True
        .Borders(xlEdgeBottom).Color = RGB(226, 234, 246)
        .Interior.Color = clrGrey1
        .EntireColumn.AutoFit
    End With
    
    For x = TitleRow + 1 To LastRow
        With Range(Cells(x, Table_rg.Column), Cells(x, LastCol))
            If y Mod 2 = 0 Then .Interior.Color = RGB(240, 240, 240)
            If y Mod 2 <> 0 Then .Interior.Color = RGB(255, 255, 255)
            .Borders(xlEdgeBottom).Color = RGB(190, 190, 190)
            .Borders(xlEdgeTop).Color = RGB(190, 190, 190)
        End With
        
        y = y + 1
    
    Next x

    With Table_rg
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).Color = vbBlack
    End With

' -----------------------------
' Freeze panes on the title row
' -----------------------------

    If TitleRow = 1 Then
        Application.GoTo Range("A1"), True
        ActiveSheet.Range("A2").Select
            ActiveWindow.FreezePanes = True
    Else
    
    Dim bolFreezeandFilter As Long
        bolFreezeandFilter = MsgBox(Prompt:="Would you like to freeze panes?", Buttons:=vbYesNo + vbQuestion, Title:="Freeze Panes?")
    
        If bolFreezeandFilter = vbYes Then
            Application.GoTo Range("A" & TitleRow), True
            Range("A" & TitleRow + 1).Select
                ActiveWindow.FreezePanes = True
        End If

    End If
    
' -----------------------------
' Filter and auto size the data
' -----------------------------
    
    Range(TitleRow & ":" & TitleRow).AutoFilter

    ActiveSheet.Cells.EntireColumn.AutoFit
    ActiveSheet.Cells.EntireColumn.AutoFit
    
    ActiveWindow.DisplayGridlines = False

Call myPrivateMacros.DisableForEfficiencyOff

End Sub
Sub u_Banded_Rows()

' Purpose: To create a conditional formatting for banded rows, and apply it to the selected cells.
' Trigger: Ribbon > Personal Macros > Data Formatting > Apply Dynamic Banded Rows
' Updated: 3/29/2020

' CHange Log:
'       3/29/2020: Updated to do odd rows with color fill and lightened banded color

' ***********************************************************************************************************************************
    
    Dim Cond_fx_Banded_Rows As FormatCondition
        Set Cond_fx_Banded_Rows = Selection.FormatConditions.Add(xlExpression, Formula1:="=Isodd(Row())")
     
    With Cond_fx_Banded_Rows
        .Interior.Color = RGB(248, 248, 248)
        .Borders(xlEdgeBottom).Color = RGB(190, 190, 190)
        .Borders(xlEdgeTop).Color = RGB(190, 190, 190)
    End With
    
End Sub
Sub u_Cycle_Thru_Tab_or_Cell_Colors()
    
' Purpose: To cycle through the tab colors for the ActiveSheet.
' Trigger: Ribbon > Personal Macros > Workbook Formatting > Change Cell/Tab Color
' Updated: 6/1/2021

' Change Log:
'       6/1/2021: Intial Creation

' ***********************************************************************************************************************************

    uf_Color_Selector.Show vbModeless

End Sub
Sub u_Create_Dynamic_Reference_Number()

' Purpose: To create an auto incremented reference number based on the cell above's value.
' Trigger: Ribbon > Personal Macros > Functions > Create Ref. Number
' Updated: 1/9/2020

' Change Log:
'       12/6/2019:  Added the max code for rngOldRef
'       1/9/2020:   Added the code to prevent overwtitting existing data
'       1/16/2020:  Moved the code to prevent overwrite to the splitter

' ***********************************************************************************************************************************

Call myPrivateMacros.DisableForEfficiency

' -----------------
' Declare Variables
' -----------------

    Dim ws As Worksheet
        Set ws = ActiveWorkbook.ActiveSheet
    
    Dim rngOldRef As Range
        Set rngOldRef = ws.Range(Cells(2, Selection.Column), Cells(Selection.Row, Selection.Column))
        If rngOldRef.Row = 1 Then Set rngOldRef = Selection.Offset(-1, 0)
    
    Dim OldRef As String
        OldRef = WorksheetFunction.Max(rngOldRef)
    
    Dim NewRef As String
        Select Case Asc(Right(OldRef, 1))
            Case 65 To 90, 97 To 122
                NewRef = Left(OldRef, Len(OldRef) - 1) & Chr(Asc(Right(OldRef, 1)) + 1)
            Case Else
                NewRef = Left(OldRef, InStrRev(OldRef, ".")) & Right(OldRef, (Len(OldRef) - InStrRev(OldRef, "."))) + 1
            End Select
    
' ------------------------------------------------
' ------------------------------------------------
' -----------

    Selection.Value = NewRef
    
    Selection.Offset(-1, 0).Copy: Selection.PasteSpecial xlPasteFormats

Call myPrivateMacros.DisableForEfficiencyOff

End Sub
Sub u_Delete_Hidden_Sheets()

' Purpose: To delete all of the hidden Worksheets in a Workbook.
' Trigger: Ribbon > Personal Macros > Functions > Delete Hidden Worksheets
' Updated: 1/14/2020

' ***********************************************************************************************************************************

    Dim ws As Worksheet
    
Application.DisplayAlerts = False
    
    For Each ws In ActiveWorkbook.Worksheets
        If ws.Visible = xlSheetHidden Then ws.Delete
    Next ws
    
Application.DisplayAlerts = True

End Sub
Sub u_Unhide_All_Worksheets()

' Purpose: To unhide all of the sheets in a workbook.
' Trigger: Ribbon > Personal Macros > Functions > Unhide All Worksheets
' Updated: 11/14/2019

' ***********************************************************************************************************************************

    Dim ws As Worksheet

    For Each ws In ActiveWorkbook.Worksheets
        ws.Visible = xlSheetVisible
    Next ws

End Sub
Sub u_Unhide_All_Columns_n_Rows_in_ActiveSheet()

' Purpose: To unhide all of the rows and columns in the ActiveSheet.
' Trigger: Ribbon > Personal Macros > Functions > Unhide All Worksheets
' Updated: 7/31/2022

' ***********************************************************************************************************************************

    With ActiveWorkbook.ActiveSheet
        If .AutoFilterMode = True Then .AutoFilter.ShowAllData
        .Columns.EntireColumn.Hidden = False
        .Rows.EntireRow.Hidden = False
    End With

End Sub
Sub u_Delete_Hidden_Columns()

' Purpose: To delete all of the hidden columns in a Worksheet.
' Trigger: Ribbon > Personal Macros > Functions > Delete Hidden Columns
' Updated: 11/29/2021

' ***********************************************************************************************************************************

' -----------------
' Declare Variables
' -----------------

    Dim wsActive As Worksheet
    Set wsActive = ActiveWorkbook.ActiveSheet

    Dim int_LastCol As Long
        int_LastCol = wsActive.Cells(1, 999).End(xlToLeft).Column

    Dim i As Long

' -------------------------
' Delete the Hidden Columns
' -------------------------
    
Application.DisplayAlerts = False
    
    For i = int_LastCol To 1 Step -1
        If wsActive.Columns(i).Hidden = True Then
            wsActive.Columns(i).Delete
        End If
    Next i
    
Application.DisplayAlerts = True

End Sub
Sub u_Toggle_Commitment_vs_Outstanding_in_PivotTable()

' Purpose: To swtich between the Commitment and Outstanding fields in a pivot table, and applicable % of Total.
' Trigger: Ribbon > Personal Macros > Functions > PT Toggle Commit v OS
' Updated: 4/15/2021

' Change Log:
'           4/14/2022:  Intial Creation
'           4/15/2022:  Updated to use VisibleFields to remove the Outstanding
'                       Added the ability to flip between Commitment and Outstanding

' ***********************************************************************************************************************************

' -----------------
' Declare Variables
' -----------------

On Error GoTo ErrorHandler

    ' Declare "Booleans"
    
    Dim bolCurrentCommitment As Boolean
    
    Dim bolCurrentOutstanding As Boolean

    ' Declare Strings

    Dim strActivePT As String
        strActivePT = ActiveCell.PivotTable.Name
        
    Dim strFieldName As String
        
    ' Declare Integers
    
    Dim i As Integer

    ' Declare Pivot Table Variables

    Dim ptActivePT As PivotTable
    Set ptActivePT = ActiveSheet.PivotTables(strActivePT)

' ------------------------------------------------------
' Determine if I currently use Outstanding vs Commitment
' ------------------------------------------------------
    
    For i = 1 To ptActivePT.DataFields.count
        If ptActivePT.DataFields(i).SourceName = "Outstanding" Then
            bolCurrentOutstanding = True
            Exit For
        ElseIf ptActivePT.DataFields(i).SourceName = "Scorecard Exposure" Then
            bolCurrentCommitment = True
            Exit For
        ElseIf ptActivePT.DataFields(i).SourceName = "Commitment" Then
            bolCurrentCommitment = True
            Exit For
        End If
    Next i

    ' Determine if the name of the field is Scorecard Exposure or Commitment

    For i = 1 To ptActivePT.PivotFields.count
        If ptActivePT.PivotFields(i).SourceName = "Scorecard Exposure" Then
            strFieldName = "Scorecard Exposure"
            Exit For
        ElseIf ptActivePT.PivotFields(i).SourceName = "Commitment" Then
            strFieldName = "Commitment"
            Exit For
        End If
    Next i

' ---------------------------------------------------------------
' Determine if I need to add the Outstanding or Commitment fields
' ---------------------------------------------------------------

    If bolCurrentOutstanding = True Then
        GoTo RemoveOutstanding
    ElseIf bolCurrentCommitment = True Then
        GoTo RemoveCommitment
    End If

' ----------------------------
' Remove the Outstanding Field
' ----------------------------

RemoveOutstanding:

    ' Find and remove the Data Fields based on Outstanding
    For i = ptActivePT.DataFields.count To 1 Step -1
        If ptActivePT.DataFields(i).SourceName = "Outstanding" Then
            ptActivePT.DataFields(i).Orientation = xlHidden
        End If
    Next i

' -------------------------------------------
' Add the Commitment field and the % of Total
' -------------------------------------------

With ptActivePT

    .AddDataField .PivotFields(strFieldName), "Commitment ", xlSum
        .PivotFields("Commitment ").NumberFormat = "$#,##0"
    
    .AddDataField .PivotFields(strFieldName), "% of Total", xlSum
        .PivotFields("% of Total").NumberFormat = "0%"
        .PivotFields("% of Total").Calculation = xlPercentOfParentRow
        
End With

Exit Sub

'--------------------------------------------------------------------------

' ----------------------------
' Remove the Outstanding Field
' ----------------------------

RemoveCommitment:

    ' Find and remove the Data Fields based on Commitment
    For i = ptActivePT.DataFields.count To 1 Step -1
        If ptActivePT.DataFields(i).SourceName = "Commitment" Or ptActivePT.DataFields(i).SourceName = "Scorecard Exposure" Then
            ptActivePT.DataFields(i).Orientation = xlHidden
        End If
    Next i

' -------------------------------------------
' Add the Outstanding field and the % of Total
' -------------------------------------------

With ptActivePT

    .AddDataField .PivotFields("Outstanding"), "Outstanding ", xlSum
        .PivotFields("Outstanding ").NumberFormat = "$#,##0"
    
    .AddDataField .PivotFields("Outstanding"), "% of Total", xlSum
        .PivotFields("% of Total").NumberFormat = "0%"
        .PivotFields("% of Total").Calculation = xlPercentOfParentRow
        
End With

Exit Sub

'--------------------------------------------------------------------------

ErrorHandler:

End Sub
Sub u_Unhide_Change_Logs()

' Purpose: To unhide all of the change logs, and wsLists, that are in the ActiveWorkbook.
' Trigger: Ribbon > Personal Macros > Specific Functions > Unhide Change Logs
' Updated: 5/3/2022

' Change Log:
'       5/3/2022:   Updated to make the code more dynamic if the logs don't exist
'                   Added the fx_Sheet_Exists code
'                   Removed all the variables and just stuck w/ the fx_Sheet_Exists

' ***********************************************************************************************************************************

Application.ScreenUpdating = False

    'Unprotect the ActiveSheet, for when updating the CV Tracker / Sageworks Validation Dashboard
    ActiveWorkbook.ActiveSheet.Unprotect
    
    'Unhide all of the logs and wsLists
    If fx_Sheet_Exists(ActiveWorkbook.Name, "CHANGE LOG") = True Then ActiveWorkbook.Sheets("CHANGE LOG").Visible = xlSheetVisible
    If fx_Sheet_Exists(ActiveWorkbook.Name, "DATA CHANGE LOG") = True Then ActiveWorkbook.Sheets("DATA CHANGE LOG").Visible = xlSheetVisible
    If fx_Sheet_Exists(ActiveWorkbook.Name, "Faux Log") = True Then ActiveWorkbook.Sheets("Faux Log").Visible = xlSheetVisible
    If fx_Sheet_Exists(ActiveWorkbook.Name, "LISTS") = True Then ActiveWorkbook.Sheets("LISTS").Visible = xlSheetVisible
    If fx_Sheet_Exists(ActiveWorkbook.Name, "PROJECT CHANGE LOG") = True Then ActiveWorkbook.Sheets("PROJECT CHANGE LOG").Visible = xlSheetVisible

Application.ScreenUpdating = True

End Sub
Sub u_Convert_to_Values()
    
' Purpose: To convert the selected text from text to values.
' Trigger: Ribbon > Personal Macros > Specific Functions > Convert to Values
' Updated: 7/2/2023

' Change Log:
'       7/4/2021:   Intial Creation
'       7/6/2021:   Added to convert the number format to a number
'       3/8/2023:   Updated to allow each record to be individually processed
'       7/2/2023:   Added the check for 0 to ONLY handle those funky ones starting in 0, and avoiding the error in the case of special characters
'       7/3/2023:   Updated to use an array for the individual record assessment

' ***********************************************************************************************************************************

Call myPrivateMacros.DisableForEfficiency

' -----------------
' Declare Variables
' -----------------

    Dim rng_TargetData As Range
    Set rng_TargetData = Selection
    
    Dim arry_TargetData() As Variant
        arry_TargetData = WorksheetFunction.Transpose(rng_TargetData.Value)
    
    Dim i As Long
    
On Error GoTo IndividualProcess

' -----------------------------------
' Attempt entire selection conversion
' -----------------------------------

    rng_TargetData.NumberFormat = "0"
    rng_TargetData.Value = Selection.Value
    
Call myPrivateMacros.DisableForEfficiencyOff
    
Exit Sub
    
' ---------------------------------------
' Do conversion at invidiual record level
' ---------------------------------------
    
IndividualProcess:

On Error Resume Next

    Debug.Print "There was an error in the 'fx_Convert_to_Values' function at " & Now & Chr(10) _
                & "As a result the conversion had to be done at the individual record level."
    
    For i = LBound(arry_TargetData) To UBound(arry_TargetData)
        If IsNumeric(arry_TargetData(i)) = True Then
            rng_TargetData(i, 1).Value = val(arry_TargetData(i))
        End If
    Next i

Call myPrivateMacros.DisableForEfficiencyOff

End Sub
Sub u_Prep_File_For_Email()

' Purpose: To prepare the ActiveWorkbook for sending out, by saving over all formulas as values only, deleting source data, and deleting support ws.
' Trigger: Ribbon > Personal Macros > Specific Functions > Prep File for Email
' Updated: 11/18/2021

' Change Log:
'       10/9/2021: Initial Creation
'       11/18/2021: Added to abort if values = values AND it's a PIVOT
'       11/18/2021: Updated to ignore the clrPurple3
'       11/18/2021: Converted to using HasFormula and to only copy over those cells w/ a formula

' ***********************************************************************************************************************************

' -----------------------------------
' Determine if the process should run
' -----------------------------------

Dim intRunProcess As Long
    intRunProcess = MsgBox( _
        Prompt:="Would you like to prep the '" & ActiveWorkbook.Name & "' workbook for emailing out?", _
        Title:="Prep File?", _
        Buttons:=vbQuestion + vbYesNo)

        If intRunProcess = 7 Then Exit Sub 'Abort if cancel was pushed

myPrivateMacros.DisableForEfficiency

' -----------------
' Declare Variables
' -----------------

    ' Declare Worksheets
    
    Dim wbSelected As Workbook
    Set wbSelected = ActiveWorkbook
    
    Dim ws As Worksheet

    ' Declare Integers
       
    Dim i As Integer
    
    ' Declare Strings

    Dim strNewFileName As String
        strNewFileName = Replace(wbSelected.Name, Right(wbSelected.Name, 5), " .xlsx") ' Add the space at the end of the file name
    
    Dim strFilePath As String
        strFilePath = wbSelected.path
    
    Dim strNewFileFullPath As String
        strNewFileFullPath = strFilePath & "\" & strNewFileName
    
    ' Declare Colors
    
    Dim clrBlue1 As Long
        clrBlue1 = RGB(220, 230, 241)
    Dim clrBlue2 As Long
        clrBlue2 = RGB(202, 217, 235)
    Dim clrBlue3 As Long
        clrBlue3 = RGB(184, 204, 228)
    
    Dim clrPurple1 As Long
        clrPurple1 = RGB(228, 223, 236)
    Dim clrPurple2 As Long
        clrPurple2 = RGB(216, 208, 227)
    Dim clrPurple3 As Long
        clrPurple3 = RGB(204, 192, 218)
    
    Dim clrGreen1 As Long
        clrGreen1 = RGB(235, 241, 222)
    Dim clrGreen2 As Long
        clrGreen2 = RGB(226, 235, 205)
    Dim clrGreen3 As Long
        clrGreen3 = RGB(216, 228, 188)
    
    ' Declare Cells
    
    Dim cell As Range
    
' ---------------------
' Copy back values only
' ---------------------
    
    For Each ws In wbSelected.Worksheets
        If ws.AutoFilterMode = True Then ws.AutoFilter.ShowAllData
        If ws.Name <> "PIVOT" Then
            
            For Each cell In ws.UsedRange
                If cell.HasFormula = True Then
                    cell.Value2 = cell.Value2
                End If
            Next cell
            
        End If
    Next ws

' ------------------------------------
' Delete the purple / green worksheets
' ------------------------------------

Application.DisplayAlerts = False

    For Each ws In wbSelected.Worksheets
        Debug.Print ws.Name
        If ws.Tab.Color = clrPurple1 Or ws.Tab.Color = clrPurple2 Then
            ws.Delete
        ElseIf ws.Tab.Color = clrBlue1 Or ws.Tab.Color = clrBlue2 Or ws.Tab.Color = clrBlue3 Then
            If ws.Name <> "GUIDE" Then
                ws.Delete
            End If
        ElseIf ws.Tab.Color = clrGreen1 Or ws.Tab.Color = clrGreen2 Or ws.Tab.Color = clrGreen3 Then
            If ws.Name <> "PIVOT" Then
                ws.Delete
            End If
        End If
        
    Next ws

Application.DisplayAlerts = True

' ----------------------------
' Delete the hidden worksheets
' ----------------------------

Application.DisplayAlerts = False

    For Each ws In wbSelected.Worksheets
        If ws.Visible <> xlSheetVisible Then
            ws.Delete
        End If
    Next ws
    
Application.DisplayAlerts = True

' -----------------
' Save the workbook
' -----------------

Application.DisplayAlerts = False

    wbSelected.SaveAs Filename:=strNewFileFullPath, FileFormat:=xlOpenXMLWorkbook
    wbSelected.Close SaveChanges:=True

Application.DisplayAlerts = True

ThisWorkbook.FollowHyperlink (strFilePath)

myPrivateMacros.DisableForEfficiencyOff

End Sub
Sub u_Add_Change_to_ChangeLogs()
    
' Purpose: To add a new change to the Change Log, if it exists in the ActiveWorkbook.
' Trigger: Ribbon > Personal Macros > Other > Add Change to Change Log
' Updated: 12/28/2021

' Change Log:
'       10/19/2021: Intial Creation
'       12/28/2021: Replaced "CHANGELOG" with "PROJECT CHANGE LOG"

' ***********************************************************************************************************************************

    'Abort if the Project Change Log doesn't exist
    If fx_Sheet_Exists(ActiveWorkbook.Name, "PROJECT CHANGE LOG") = False Then
        MsgBox "Change Log does not exist, change was unable to be added."
        Exit Sub
    End If
    
    'Open the form and force the Change Description object to take Focus
    
    uf_Add_Change_To_ChangeLog.Show vbModeless
    
    uf_Add_Change_To_ChangeLog.txt_Change_Desc.Enabled = False
    uf_Add_Change_To_ChangeLog.txt_Change_Desc.Enabled = True
        
    uf_Add_Change_To_ChangeLog.txt_Change_Desc.SetFocus
    
End Sub
Sub u_Toggle_ActiveFields_wsTasks()

' Purpose: To hide / unhide the Active fields in the ws_Tasks.
' Trigger: Quick Access Toolbar
' Updated: 1/26/2022

' Change Log:
'       1/7/2022:   Intial Creation
'       1/26/2022:  Added the "Start" field to be toggled on / off

' ***********************************************************************************************************************************

' -----------------
' Declare Variables
' -----------------
    
    ' Declare Worksheets
    Dim ws_Tasks As Worksheet
    Set ws_Tasks = ThisWorkbook.Sheets("Tasks")
    
    ' Declare Cell References
    Dim int_LastCol As Long
        int_LastCol = ws_Tasks.Cells(1, ws_Tasks.Columns.count).End(xlToLeft).Column

    Dim arry_Header() As Variant
        arry_Header = Application.Transpose(ws_Tasks.Range(ws_Tasks.Cells(1, 1), ws_Tasks.Cells(1, int_LastCol)))
       
    Dim col_ActiveProject As Long
        col_ActiveProject = fx_Create_Headers("Active Proj.", arry_Header)
        
    Dim col_ActiveComponent As Long
        col_ActiveComponent = fx_Create_Headers("Active Comp.", arry_Header)
    
    Dim col_ActiveTask As Long
        col_ActiveTask = fx_Create_Headers("Active Task", arry_Header)

    Dim col_Start As Long
        col_Start = fx_Create_Headers("Start", arry_Header)

    ' Declare Booleans
    Dim bolHiddenFields As Boolean
        If ws_Tasks.Columns(col_ActiveTask).Hidden = True Then
            bolHiddenFields = True
        Else
            bolHiddenFields = False
        End If

' ------------------------
' Hide / Unhide the fields
' ------------------------
    
     ws_Tasks.Columns(col_ActiveProject).Hidden = Not bolHiddenFields
     ws_Tasks.Columns(col_ActiveComponent).Hidden = Not bolHiddenFields
     ws_Tasks.Columns(col_ActiveTask).Hidden = Not bolHiddenFields
     
     ws_Tasks.Columns(col_Start).Hidden = Not bolHiddenFields

End Sub
Private Sub u_Open_Personal_Macro_Workbook_ARCHIVED()

' Purpose: To allow me to quickly open and close the Personal Macro Workbook that Axcel created.
' Trigger: Ribbon > Personal Macros > Functions > Open / Close Personal Macro Workbook
' Updated: 9/16/2019

' ***********************************************************************************************************************************

' -----------------
' Declare Variables
' -----------------
        
    Dim strPMWorkbookPath As String
        strPMWorkbookPath = "C:\U Drive\Support\Other\PERSONAL.XLSB"
        
    Dim wb As Workbook
        
    Dim wbOpen As Workbook

' ------------------------------------------------------------------
' Determine if the Personal Macro wb is already open, if not open it
' ------------------------------------------------------------------
    
    For Each wb In Workbooks
        If wb.FullName = strPMWorkbookPath Then Set wbOpen = wb
    Next wb
        
    If wbOpen Is Nothing Then
        Set wbOpen = Workbooks.Open(strPMWorkbookPath)
        Application.Windows(wbOpen.Name).Visible = True 'Added on 10/21/19 to allow me to say yes to Macros
    Else
        Application.DisplayAlerts = False
            wbOpen.Close
        Application.DisplayAlerts = True
    End If

End Sub
Sub u_Modify_Entitlements_Review_Workbooks()

' Purpose: To manipulate the data in the entitlement review workbooks to ease in the completion of the review.
' Trigger: Ribbon > Specific Macros > Prep Entitlements Review Workbook
' Updated: 11/21/2023
'
' Change Log:
'       11/20/2023: Intial Creation
'       11/21/2023: Added 'MAINTENANCE' as a term to flag for Priv access

' ***********************************************************************************************************************************

' -----------------
' Declare Variables
' -----------------
    
    ' Declare Worksheets
    Dim wsEntitlements As Worksheet
    Set wsEntitlements = ActiveWorkbook.Sheets(1)
    

    ' Declare "Ranges"
    Dim intLastRow As Long
        intLastRow = fx_Find_LastRow(ws_Target:=wsEntitlements, bolIncludeUsedRange:=True, bol_MinValue2:=True)
        
    Dim intLastCol As Long
        intLastCol = fx_Find_LastColumn(ws_Target:=wsEntitlements, bolIncludeUsedRange:=True)
    
    ' Declare Columns
    Dim colAttribute As Long
        colAttribute = 2
        
    Dim colPrivileged As Long
        colPrivileged = 5
    
    Dim colJamesReview As Long
        colJamesReview = intLastCol + 1
        
    ' Declare Ranges
    Dim rngDataRange As Range
    Set rngDataRange = wsEntitlements.Range(wsEntitlements.Cells(2, 1), wsEntitlements.Cells(intLastRow, colJamesReview))
    
    ' Declare Strings
    Dim str_NewFilePath As String
        str_NewFilePath = "C:\U Drive\Projects\P.23.455 - T3 Privileged Users Review\Manasa Support\(NEW)\(REVIEWED)"
    
    ' Declare Arrays
    Dim arryData() As Variant
        arryData = wsEntitlements.Range(wsEntitlements.Cells(1, 1), wsEntitlements.Cells(intLastRow, colJamesReview))
    
    ' Declare Loop Variables
    Dim i As Long 'Used for looping through the Priv entitlement data
    
    Dim y As Long 'Used for the formatting loop
    
    Dim x As Long 'Used for the formatting loop
    
' --------------------
' Modify the worksheet
' --------------------
    
With wsEntitlements
    
    ' Insert the new Column
    .Cells(1, colJamesReview).Value2 = "James' Review"

    ' Apply the custom fill for the lookups and user generated fields
    .Range(.Cells(1, 1), .Cells(1, intLastCol)).Interior.Color = RGB(230, 230, 230) 'Gray Fill
    .Cells(1, colJamesReview).Interior.Color = RGB(211, 223, 238) 'Blue Fill
    
    ' Bold the Headers
    .Range(.Cells(1, 1), .Cells(1, colJamesReview)).Font.Bold = True
    
    ' Apply the Data Validation
    .Range(.Cells(1, colJamesReview), .Cells(intLastRow, colJamesReview)).Validation.Add Type:=xlValidateList, Formula1:="Privileged, Not Priv"

    ' Apply the line formatting
    With rngDataRange
         .Borders(xlEdgeBottom).Color = RGB(190, 190, 190)
         .Borders(xlInsideHorizontal).Color = RGB(190, 190, 190)
         .Borders(xlEdgeTop).Color = RGB(190, 190, 190)
    End With
    
    ' Apply the alternating fill
    For x = 2 To intLastRow
        With Range(Cells(x, 1), Cells(x, colJamesReview))
            If y Mod 2 = 0 Then .Interior.Color = RGB(240, 240, 240)
            If y Mod 2 <> 0 Then .Interior.Color = RGB(255, 255, 255)
        End With
        
        y = y + 1
    
    Next x
    
End With
    
' -----------------------------
' Filter and auto size the data
' -----------------------------
    
    Range("1:1").AutoFilter

    wsEntitlements.Cells.EntireColumn.AutoFit
    wsEntitlements.Cells.EntireColumn.AutoFit
    
    ActiveWindow.DisplayGridlines = False
    
    Range("2:2").Activate
    ActiveWindow.FreezePanes = True
    Application.GoTo Range("A1"), True

' --------------------------
' Complete the Priv Analysis
' --------------------------

    ' Flag anything already marked as Priv
    For i = 2 To intLastRow
        If arryData(i, colPrivileged) = True Then arryData(i, colJamesReview) = "Privileged"
    Next i
    
    ' Flag anything with the key terms in it
    For i = 2 To intLastRow
        Select Case True ' If there is a match to one of the terms mark as Privileged
            Case UCase(arryData(i, colAttribute)) Like "*ADMIN*"
                arryData(i, colJamesReview) = "Privileged"
            Case UCase(arryData(i, colAttribute)) Like "*GOD*"
                arryData(i, colJamesReview) = "Privileged"
            Case UCase(arryData(i, colAttribute)) Like "*SUPER*"
                arryData(i, colJamesReview) = "Privileged"
            Case UCase(arryData(i, colAttribute)) Like "*POWER*"
                arryData(i, colJamesReview) = "Privileged"
            Case UCase(arryData(i, colAttribute)) Like "*OWNER*"
                arryData(i, colJamesReview) = "Privileged"
            Case UCase(arryData(i, colAttribute)) Like "*DBO*"
                arryData(i, colJamesReview) = "Privileged"
            Case UCase(arryData(i, colAttribute)) Like "*MAINTENANCE*"
                arryData(i, colJamesReview) = "Privileged"
        End Select
    Next i
    
    ' Flag anything that was left out as "Not Priv"
    For i = 2 To intLastRow
        If arryData(i, colJamesReview) <> "Privileged" Then
            arryData(i, colJamesReview) = "Not Priv"
        End If
    Next i
    
    ' Apply the highlighting for anything that doesn't match to the owner's review
    For i = 2 To intLastRow
        If arryData(i, colPrivileged) <> True And arryData(i, colJamesReview) = "Privileged" Then
            wsEntitlements.Cells(i, colPrivileged).Interior.Color = RGB(251, 216, 197)
            wsEntitlements.Cells(i, colJamesReview).Interior.Color = RGB(251, 216, 197)
        End If
    Next i
    
    ' Return the data to the worksheet
    wsEntitlements.Range(wsEntitlements.Cells(1, 1), wsEntitlements.Cells(intLastRow, colJamesReview)) = arryData
        wsEntitlements.Cells(1, colJamesReview).Value2 = "James' Review"
    
' -------------
' Save the File
' -------------

    Call fx_Copy_to_Clipboard(strTextToCopy:=wsEntitlements.Parent.Name) 'Copy the filename to the clipboard

    Call fx_Close_Workbook( _
        objWorkbookToClose:=wsEntitlements.Parent, _
        bol_KeepWorkbookOpen:=True, _
        bol_SaveWorkbook:=True, _
        strAddToWorkbookName:="(JR Reviewed)", _
        bol_ConvertToXLSX:=True, _
        strNewFilePath:=str_NewFilePath)  'Save the Named Ranges
    
End Sub


