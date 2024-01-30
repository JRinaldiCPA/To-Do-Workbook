VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} uf_Add_Change_To_ChangeLog 
   Caption         =   "Project Change Log Input Form"
   ClientHeight    =   4680
   ClientLeft      =   36
   ClientTop       =   396
   ClientWidth     =   7668
   OleObjectBlob   =   "uf_Add_Change_To_ChangeLog.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "uf_Add_Change_To_ChangeLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim Workbooks
    Dim wbTarget As Workbook
    
'Dim Sheets
    Dim wsProjectChangeLog As Worksheet
    
'Dim Cell References
    Dim arry_Header() As Variant
    
    Dim col_Updated As Long
    Dim col_Version As Long
    Dim col_Details As Long
    
'Dim Strings
    Dim strWBVersion As String
    Dim strCurVersion As String
    Dim strNewVersion As String
    
'Dim Integers
    Dim int_CurRow As Long
Private Sub UserForm_Initialize()

' Purpose: To initialize the userform, including adding in the data from the arrays and pulling data from the current row if a Project is selected.
' Trigger: Event: UserForm_Initialize
' Updated: 4/19/2022

' Semantic Versioning
'   Start with Version 1.0.0 (X.Y.Z)
'   MAJOR: (ex. 1.0.0 -> 2.0.0) Major Change or new data set used
'   MINOR: (ex. 1.1.0 -> 1.2.0) Minor change, change to functionality, etc.
'   PATCH: (ex. 1.1.1 -> 1.1.2) Insignificant change, or cosmetic changes

' Change Log:
'       10/19/2021: Initial Creation
'       11/13/2021: Updated the code related to the strWBVersion
'       11/17/2021: Added code to handle an error where the strWBVersion pulls in most of the file name (ex. 'H Past Due Review v1.3.')
'       12/1/2021:  Added the strVersionType code and default the right type of update
'       12/3/2021:  Default to yes to increment the file #
'       12/28/2021: Replaced "CHANGELOG" with "PROJECT CHANGE LOG"
'                   Added the column references and Change Type field
'       3/2/2022:   Updated to capture the version of the workbook in the V_ProjectVersion named range
'       3/23/2022:  Added code to expliclty capture the ActiveWorkbook and use it throughout
'       4/19/2022:  Updated to handle a 4 digit version number

' ***********************************************************************************************************************************

If Evaluate("ISREF(" & "'PROJECT CHANGE LOG'" & "!A1)") <> True Then
    MsgBox "The Project Change Log worksheet doesn't exist in " & ActiveWorkbook.Name
    Exit Sub
End If

Call Me.o_02_Assign_Private_Variables

' -----------------
' Declare Variables
' -----------------

    ' Declare / Assign Workbook Name Variables
        
    'Dim wbTarget As Workbook
    Set wbTarget = ActiveWorkbook
    
    Dim strWBName As String
        strWBName = wbTarget.Name
        
        strWBVersion = Mid(String:=strWBName, Start:=InStr(strWBName, " (v") + 3, Length:=Len(strWBName))
        strWBVersion = Left(strWBVersion, Len(strWBVersion) - 6) ' 6 to account for the )
        
    If InStr(1, strWBVersion, ".") = 0 Or Len(strWBVersion) > 12 Then
        strWBVersion = Mid(String:=strWBName, Start:=InStrRev(strWBName, "v") + 1, Length:=Len(strWBName))
        strWBVersion = Left(strWBVersion, Len(strWBVersion) - 5)
    End If

    If Right(strWBVersion, 1) = ")" Then
        strWBVersion = Left(strWBVersion, Len(strWBVersion) - 1)
    End If

    Dim intPeriodCount As Integer
        intPeriodCount = Len(strWBVersion) - Len(Replace(strWBVersion, ".", ""))

    Dim strVersionType As String
        If intPeriodCount = 2 Or intPeriodCount = 3 Then
            strVersionType = "Patch"
        ElseIf intPeriodCount = 1 Then
            strVersionType = "Minor"
        Else
            strVersionType = "Major"
        End If

    ' Declare Zoom Variables
                   
    Dim dblZoomAdjust As Double
        dblZoomAdjust = 1.08

' ---------------------
' Initialize the values
' ---------------------

    ' Add value for the Current Version of the workbook
        txt_CurVersion.Value = strWBVersion
    
    ' Add Value for Entered TextBox
        txt_Updated.Value = Date

    ' Select the default update type
    
    If strVersionType = "Major" Then
        Me.opt_Minor = True
    ElseIf strVersionType = "Minor" Then
        Me.opt_Patch = True
    ElseIf strVersionType = "Patch" Then
        Me.opt_Patch = True
    End If
        
    ' Default to increment the file
    Me.chk_Increment_File_Version = True
        
' -------------------------------
' Adjust the size of the UserForm
' -------------------------------
        
    Me.Zoom = Me.Zoom * dblZoomAdjust
    Me.Height = Me.Height * dblZoomAdjust
    Me.Width = Me.Width * dblZoomAdjust

End Sub
Private Sub cmd_Add_Change_Click()

' Purpose: To add the change from the UserForm to the ChangeLog ws.
' Updated: 11/13/2021

' Change Log:
'       11/13/2021: Intial Creation

' ***********************************************************************************************************************************

Call myPrivateMacros.DisableForEfficiency

    ' Abort if a new change wasn't submitted
    If txt_Change_Desc = "" Then
        MsgBox "You did not submit a change"
        Exit Sub
    End If

    ' Update the Workbook Version (strNewVersion)
    Call Me.o_1_Update_Workbook_Version
    
    ' Add the new change to the ChangeLog
    Call Me.o_21_Add_New_Change
    
    ' Update the V_ProjectVersion
    Call Me.o_22_Updated_ProjectVersion_NamedRange

    ' Increment the file version if that option was selected
    If chk_Increment_File_Version = True Then
        Call Me.o_3_Save_Workbook
    End If

Call myPrivateMacros.DisableForEfficiencyOff

    Unload Me

End Sub
Private Sub cmd_Save_Workbook_Click()

' Purpose: To save the ActiveWorkbook with a new version #.
' Updated: 11/13/2021

' Change Log:
'       11/13/2021: Intial Creation

' ***********************************************************************************************************************************

Call myPrivateMacros.DisableForEfficiency


    ' Update the Workbook Version (strNewVersion)
    Call Me.o_1_Update_Workbook_Version
    
    ' Increment the file version if that option was selected
    Call Me.o_3_Save_Workbook

Call myPrivateMacros.DisableForEfficiencyOff

    Unload Me

End Sub
Private Sub cmd_Cancel_Click()

    Unload Me

End Sub
Sub o_02_Assign_Private_Variables()

' Purpose: To assign the Private Variables that were Declared above the line.
' Trigger: Event: UserForm_Initialize
' Updated: 4/18/2022

' Change Log:
'       11/13/2021: Intial Creation
'       12/28/2021: Updated the int_CurRow to use fx_Find_CurRow
'                   Added the column references and Change Type field
'       3/23/2022:  Replaced ActiveWorkbook w/ wbTarget
'       4/18/2022:  Removed the "Change Type" code

' ***********************************************************************************************************************************
    
' ----------------
' Assign Variables
' ----------------
    
    ' Assign Worksheets

    Set wsProjectChangeLog = ActiveWorkbook.Sheets("PROJECT CHANGE LOG")
    
    ' Assign Cell References
    
    arry_Header = Application.Transpose(wsProjectChangeLog.Range(wsProjectChangeLog.Cells(1, 1), wsProjectChangeLog.Cells(1, 5)))

    col_Updated = fx_Create_Headers("Updated", arry_Header)
    col_Version = fx_Create_Headers("Version", arry_Header)
    col_Details = fx_Create_Headers("Details / Notes", arry_Header)

    ' Assign Integers
    
    int_CurRow = fx_Find_CurRow(wsProjectChangeLog, "Updated", "")
        
End Sub
Sub o_1_Update_Workbook_Version()

' Purpose: To determine the new version of the workbook.
' Trigger: Called: uf_Add_Change_To_ChangeLog
' Updated: 4/19/2022

' Change Log:
'       11/13/2021: Intial Creation
'       4/19/2022:  Updated to handle a 4 digit version number

' ***********************************************************************************************************************************

' -----------------
' Declare Variables
' -----------------
    
    ' Declare Strings
    Dim strChangeType As String
    
        strCurVersion = Me.txt_CurVersion
        
    Dim strTempVersion As String
        
    ' Declare Change Type Variables
        
    If opt_Major = True Then
        strChangeType = "Major"
    ElseIf opt_Minor = True Then
        strChangeType = "Minor"
    ElseIf opt_Patch = True Then
        strChangeType = "Patch"
    End If
    
    ' Declare Integers
    Dim intPeriodCount As Integer
        intPeriodCount = Len(strCurVersion) - Len(Replace(strCurVersion, ".", ""))
    
    ' Declare Version Variables
    
    Dim strMajorVer As String
    
    If intPeriodCount = 0 Then
        strMajorVer = strCurVersion
    ElseIf intPeriodCount > 0 Then
        strMajorVer = Mid(String:=strCurVersion, Start:=1, Length:=InStr(strCurVersion, ".") - 1)
    End If
    
    Dim strMinorVer As String
    
    If intPeriodCount = 1 Then
        strMinorVer = Mid(String:=strCurVersion, Start:=InStr(strCurVersion, ".") + 1, Length:=Len(strCurVersion))
    ElseIf intPeriodCount > 1 Then
        strMinorVer = Mid(String:=strCurVersion, Start:=InStr(strCurVersion, ".") + 1, Length:=Len(strCurVersion))
        strMinorVer = Mid(String:=strMinorVer, Start:=1, Length:=InStr(strMinorVer, ".") - 1)
    End If
    
    Dim strPatchVer As String

    If intPeriodCount = 2 Then
        strPatchVer = Mid(String:=strCurVersion, Start:=InStrRev(strCurVersion, ".") + 1, Length:=Len(strCurVersion))
    ElseIf intPeriodCount > 2 Then
        strPatchVer = Mid(String:=strCurVersion, Start:=InStrRev(strCurVersion, ".", InStrRev(strCurVersion, ".") - 1) + 1, Length:=Len(strCurVersion) - InStrRev(strCurVersion, "."))
    End If
    
' -----------------------------------------------
' Increment the version, based on the Change Type
' -----------------------------------------------
    
    If strChangeType = "Major" Then
    
        If intPeriodCount = 0 Then
            strNewVersion = strMajorVer + 1
        ElseIf intPeriodCount = 1 Or intPeriodCount = 2 Then
            strNewVersion = strMajorVer + 1 & ".0"
        End If
    
    ElseIf strChangeType = "Minor" Then
    
        If intPeriodCount = 0 Then
            strNewVersion = strMajorVer & ".1"
        ElseIf intPeriodCount = 1 Or intPeriodCount = 2 Then
            strNewVersion = strMajorVer & "." & strMinorVer + 1
        End If
        
    ElseIf strChangeType = "Patch" Then
    
        If intPeriodCount = 0 Then
            strNewVersion = strMajorVer & ".0.1"
        ElseIf intPeriodCount = 1 Then
            strNewVersion = strMajorVer & "." & strMinorVer & ".1"
        ElseIf intPeriodCount >= 2 Then
            strNewVersion = strMajorVer & "." & strMinorVer & "." & strPatchVer + 1
        End If
        
    End If

'Debug.Print "New Version: " & strNewVersion

End Sub
Sub o_21_Add_New_Change()

' Purpose: To add a new change to the ChangeLog.
' Trigger: Called: uf_Add_Change_To_ChangeLog > cmd_Add_Change_Click
' Updated: 4/18/2022

' Change Log:
'       10/19/2021: Intial Creation
'       11/16/2021: Updated o include the little v for Version
'       12/28/2021: Added the column references and Change Type field
'       3/2/2022:   Updated to force the conversion of the Date Updated to be a value, and hence become a date
'       4/18/2022:  Removed the "Change Type" code

' ***********************************************************************************************************************************

' ------------------
' Add the new change
' ------------------

With wsProjectChangeLog
    .Cells(int_CurRow, col_Updated) = txt_Updated
        .Cells(int_CurRow, col_Updated).Value = .Cells(int_CurRow, col_Updated).Value
    .Cells(int_CurRow, col_Version) = "v" & strNewVersion
    .Cells(int_CurRow, col_Details) = txt_Change_Desc
End With

' ------------------------
' Apply the row formatting
' ------------------------

    With wsProjectChangeLog.Range(wsProjectChangeLog.Cells(int_CurRow, col_Updated), wsProjectChangeLog.Cells(int_CurRow, col_Details))
        If int_CurRow Mod 2 = 0 Then .Interior.Color = RGB(240, 240, 240)
        If int_CurRow Mod 2 <> 0 Then .Interior.Color = RGB(255, 255, 255)
        .Borders(xlEdgeBottom).Color = RGB(190, 190, 190)
        .Borders(xlEdgeTop).Color = RGB(190, 190, 190)
    End With
        
    wsProjectChangeLog.Cells(int_CurRow, col_Version).HorizontalAlignment = xlCenter
    
    wsProjectChangeLog.Cells(int_CurRow, col_Details).WrapText = True
                        
End Sub
Sub o_22_Updated_ProjectVersion_NamedRange()

' Purpose: To add the new version to the ProjectVersion named range.
' Trigger: Called: uf_Add_Change_To_ChangeLog > cmd_Add_Change_Click
' Updated: 12/28/2021

' Change Log:
'       10/19/2021: Intial Creation
'       11/16/2021: Updated o include the little v for Version
'       12/28/2021: Added the column references and Change Type field

' ***********************************************************************************************************************************

    On Error Resume Next

    Debug.Print Range("V_ProjectVersion").Rows.count

    ' If the Named Range doesn't exist end, otherwise update it
    If Err = 1004 Then
        GoTo ExitSub
    Else
        Range("V_ProjectVersion").Value = "v" & strNewVersion
    End If

ExitSub:
    Err.Clear
    On Error GoTo 0

End Sub
Sub o_3_Save_Workbook()

' Purpose: To save a new version of the workbook with the incremented file name.
' Trigger: Called: uf_Add_Change_To_ChangeLog > cmd_Add_Change_Click
' Updated: 3/23/2022

' Change Log:
'       11/13/2021: Intial Creation
'       3/23/2022:  Replaced ActiveWorkbook w/ wbTarget

' ***********************************************************************************************************************************

' -----------------
' Declare Variables
' -----------------
    
    ' Declare Strings
    Dim strWorkbookPath As String
        strWorkbookPath = wbTarget.FullName
        
    Dim strWorkbookNewPath As String

' ---------------------
' Save the new workbook
' ---------------------

    strWorkbookNewPath = Replace(strWorkbookPath, strCurVersion, strNewVersion)

    wbTarget.SaveAs strWorkbookNewPath

End Sub
