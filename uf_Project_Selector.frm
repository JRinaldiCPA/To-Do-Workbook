VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} uf_Project_Selector 
   Caption         =   "  --- Project Selector  --- "
   ClientHeight    =   5592
   ClientLeft      =   120
   ClientTop       =   612
   ClientWidth     =   9132.001
   OleObjectBlob   =   "uf_Project_Selector.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "uf_Project_Selector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Declare Sheets
Dim ws_Projects As Worksheet
Dim ws_Tasks As Worksheet
Dim ws_Waiting As Worksheet
Dim ws_Questions As Worksheet

'Declare Strings
Dim strProjName As String
Dim strProjArea As String

'Declare Integers
Dim intProjRow As Long

'Declare ws_Projects Cell References
Dim int_LastCol_wsProjects As Long
    
Dim arry_Header_wsProjects() As Variant

Dim col_Area_wsProjects As Integer
Dim col_Project_wsProjects As Integer
Dim col_Status_wsProjects As Integer

'Declare wsTask Cell References
Dim arry_Header_wsTasks() As Variant

Dim col_Project_wsTasks As Integer
Dim col_Completed_wsTasks As Integer

'Declare ws_Waiting Cell References
Dim arry_Header_wsWaiting() As Variant

Dim col_Project_wsWaiting As Integer
Dim col_Completed_wsWaiting As Integer

'Declare ws_Questions Cell References
Dim arry_Header_wsQuestions() As Variant

Dim col_Project_wsQuestions As Integer
Dim col_Completed_wsQuestions As Integer

'Declare Collections
Dim coll_Folders As New Collection

Option Explicit



Private Sub UserForm_Initialize()

' Purpose: To initialize the userform, including adding in the data from the arrays.
' Trigger: Keyboard Shortcut - Ctrl + Shift + O (called by o_14_Open_Project_Folder)
' Updated: 6/28/2023

' Change Log:
'       2/26/2020:  I removed the IF for the Projects Dynamic Search to include ALL Projects listed, not just those that are Active / Pending / Continuous
'       2/26/2020:  Updated all of the worksheet references to be Global
'       4/12/2020:  Removed any reference to the D/A Worksheet, removed with my move to CR&A
'       11/21/2021: Added the filtering for the ws_Questions
'       12/10/2021: Removed the code related to the P.XXX name, and bolDARequest, now that I include that in the Project Name
'       12/31/2021: Updated the formatting and rearragned some things
'       6/28/2023:  Updated to adjust the height of the Areas list for Personal vs Professional

' ***********************************************************************************************************************************

' ----------------
' Assign Variables
' ----------------
    
    Call Me.o_021_Declare_Global_Variables

' ---------------------
' Initialize the values
' ---------------------
       
    'Add the values for the Area ListBox
        
        Me.lst_Area.List = GetAreaArray

    'If Projects ws is Active pull in the data from the current row
       
    On Error Resume Next
            
        If ws_Projects.Name = ActiveWorkbook.ActiveSheet.Name Then
            lst_Area.Value = ws_Projects.Range("C" & Selection.Row).Value
            lst_Project.Value = ws_Projects.Range("D" & Selection.Row).Value
            
            Me.cmd_Open_Everything.SetFocus
            
        ElseIf ws_Tasks.Name = ActiveWorkbook.ActiveSheet.Name Then
            lst_Area.Value = ws_Tasks.Range("H" & Selection.Row).Value
            lst_Project.Value = ws_Tasks.Range("I" & Selection.Row).Value
            Me.cmd_Open_Everything.SetFocus
        Else
            Me.cmb_DynamicSearch.SetFocus
        End If

' ----------------------------------------
' Adjust the size of lst_Area for Personal
' ----------------------------------------

    #If Personal = 1 Then
        Me.lst_Area.Height = Me.lst_Area.Height + 16
        Me.chk_Continuous.Top = Me.chk_Continuous.Top + 16
        Me.chk_Pending.Top = Me.chk_Pending.Top + 16
    #End If

End Sub
Private Sub lst_Area_Click()
    
    Call Me.o_03_Create_Project_List

End Sub
Private Sub lst_Project_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    
' Purpose: To replace the values in the Projects ListBox based on what is typed.
' Updated: 10/24/2023

' Change Log:
'       2/3/2023:   Moved the code so it only applies the formatting if I am in the "Project:" cell
'                   Added the code so that if I am NOT in the "Project:" cell it opens the folder
'       6/27/2023:  Removed the code related to copying into clipboard, since it was redundant
'       10/24/2023: Updated to include both "Project: " and "Project: TBD"

' ***********************************************************************************************************************************
    
    ' Determine if I am updating my meeting notes, if not then open the project folder
    
    If Left(ActiveCell.Value, 9) = "Project: " Then
        ActiveCell.Value = Left(ActiveCell.Value, 9) & Me.lst_Project.Value
        
        ActiveCell.Font.Bold = False
        ActiveCell.Characters(1, Len("Project:")).Font.Bold = True
    
    Else
        Call Me.o_11_Open_Project_Folder
        fx_Copy_to_Clipboard (lst_Project.Value)
    End If
    
    Unload Me

End Sub
Private Sub lst_Project_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    If KeyCode = vbKeyReturn And Me.lst_Project.Value <> "" Then
        Call Me.o_12_Open_Project_Support
        Unload Me
    ElseIf KeyCode = vbKeySpace And Me.lst_Project.Value <> "" Then
        fx_Copy_to_Clipboard (lst_Project.Value)
    End If

End Sub
Private Sub cmd_Open_Project_Folder_Click()

    Call Me.o_11_Open_Project_Folder

    Unload Me
    
End Sub
Private Sub cmd_Open_Project_Support_Click()

    Call Me.o_12_Open_Project_Support
    
    Unload Me

End Sub
Private Sub cmd_Filter_Click()

Call myPrivateMacros.DisableForEfficiency
    
    Call Me.o_021_Declare_Global_Variables
    Call Me.o_022_Declare_Project_Variables
    Call Me.o_21_Filter_wsTasks
    Call Me.o_22_Filter_wsWaiting
    Call Me.o_23_Filter_wsQuestions
    
Call myPrivateMacros.DisableForEfficiencyOff
    
End Sub
Private Sub cmd_Open_Everything_Click()

Call myPrivateMacros.DisableForEfficiency
    
    Call Me.o_021_Declare_Global_Variables
    Call Me.o_022_Declare_Project_Variables
    Call Me.o_12_Open_Project_Support
    Call Me.o_21_Filter_wsTasks
    Call Me.o_22_Filter_wsWaiting
    Call Me.o_23_Filter_wsQuestions

Call myPrivateMacros.DisableForEfficiencyOff

    Unload Me

End Sub
Private Sub cmd_Dynamic_Folder_Search_Click()

    Call o_82_Open_Dynamic_Folder_Search
    
    Unload Me
    
End Sub
Private Sub cmd_Dynamic_RefX_Search_Click()

    Call o_83_Open_Dynamic_RefX_Search
    
    Unload Me

End Sub
Private Sub cmd_Dynamic_File_Search_Click()

    Call o_84_Open_Dynamic_File_Search
    
    Unload Me

End Sub
Private Sub chk_Pending_Click()

Call Me.o_03_Create_Project_List

End Sub
Private Sub chk_Continuous_Click()

Call Me.o_03_Create_Project_List

End Sub
Private Sub cmd_Open_Flowchart_Click()

    Call Me.o_31_Open_Flowchart
    
    Unload Me

End Sub
Private Sub cmd_Open_Project_Click()

    Call Me.o_32_Open_Project_Workbook
    
    Unload Me

End Sub
Private Sub cmb_DynamicSearch_Change()

' Purpose: To replace the values in the Projects ListBox based on what is typed.
' Trigger: Called by uf_Project_Selector
' Updated: 12/10/2021

' Change Log:
'       4/12/2020: Removed any reference to the D/A Worksheet, removed with my move to CR&A
'       12/10/2021: Removed the code related to the P.XXX name, now that I include that in the Project Name

' ***********************************************************************************************************************************

' -----------------
' Declare Variables
' -----------------

    Dim strProjArea As String
    
    Dim strProjName As String
    
    Dim strProjStatus As String
    
    Dim x As Long
    
    Dim y As Long
        y = 1
    
    Dim ary_Projects As Variant
        ReDim ary_Projects(1 To 999)
    
' ------------
' Run the loop
' ------------
       
    Me.lst_Project.Clear
       
    With ws_Projects
            x = 2
        Do While .Range("A" & x).Value2 <> ""
            
            strProjArea = .Range("C" & x).Value2
            strProjStatus = .Range("G" & x).Value2
            strProjName = .Range("D" & x).Value2

                If InStr(1, strProjName, Me.cmb_DynamicSearch.Value, vbTextCompare) Then
                        ary_Projects(y) = strProjName
                    y = y + 1
        
                End If
            
            x = x + 1
        Loop
    End With

    ReDim Preserve ary_Projects(1 To y)

    Me.lst_Project.List = ary_Projects

' ------------------------------
' If only one remains, select it
' ------------------------------

    If lst_Project.ListCount = 2 Then
        lst_Project.Selected(0) = True
    End If

End Sub

Private Sub cmd_Cancel_Click()
    
    Unload Me

End Sub
Sub o_021_Declare_Global_Variables()

' Purpose: To set the variables in ONE location for the rest of the subs.
' Trigger: Called: uf_Project_Selector
' Updated: 5/9/2023

' Change Log
'       12/10/2021: Removed the code related to the P.XXX name, and bolDARequest, now that I include that in the Project Name
'       5/9/2023:   Wiped all the old code and replaced with the ws_Projects references
'                   Moved all the worksheet references in

' ***********************************************************************************************************************************

' -----------------
' Declare Variables
' -----------------
    
    ' Assign Worksheets
    
    Set ws_Projects = ThisWorkbook.Sheets("Projects")
    Set ws_Tasks = ThisWorkbook.Sheets("Tasks")
    Set ws_Waiting = ThisWorkbook.Sheets("Waiting")
    Set ws_Questions = ThisWorkbook.Sheets("Questions")
    
    ' Assign ws_Projects Cell References
    
    int_LastCol_wsProjects = ws_Projects.Cells(1, ws_Projects.Columns.count).End(xlToLeft).Column
    
    arry_Header_wsProjects = Application.Transpose(ws_Projects.Range(ws_Projects.Cells(1, 1), ws_Projects.Cells(1, int_LastCol_wsProjects)))

    col_Area_wsProjects = fx_Create_Headers("Area", arry_Header_wsProjects)
    col_Project_wsProjects = fx_Create_Headers("Project", arry_Header_wsProjects)
    col_Status_wsProjects = fx_Create_Headers("Status", arry_Header_wsProjects)
    
End Sub
Sub o_022_Declare_Project_Variables()

' Purpose: To set the variables related to the selected Project.
' Trigger: Called: uf_Project_Selector
' Updated: 5/9/2023

' Change Log
'       5/9/2023:   Broke out from the o_021_Declare_Global_Variables

' ***********************************************************************************************************************************

' -----------------
' Declare Variables
' -----------------
    
    ' Assign Worksheets
    
    'Dim strProjName As String
        strProjName = lst_Project.Value
    
    'Dim strProjArea As String
        If IsNull(lst_Area.Value) Then
            strProjArea = ws_Projects.Range("C" & Application.WorksheetFunction.Match(Me.lst_Project.Value, ws_Projects.Columns(4), 0))
        Else
            strProjArea = lst_Area.Value
        End If
        
    'Dim intProjRow as Long
        intProjRow = Application.WorksheetFunction.Match(Me.lst_Project.Value, ws_Projects.Columns(4), 0)

End Sub
Sub o_23_Assign_Variables_wsTasks()

' Purpose: To assign the Variables related to ws_Tasks.
' Trigger: Called by various procedures
' Updated: 11/5/2021

' Change Log:
'       11/5/2021: Intial Creation

' ***********************************************************************************************************************************

' ----------------------
' Assign Cell References
' ----------------------

    arry_Header_wsTasks = Application.Transpose(ws_Tasks.Range(ws_Tasks.Cells(1, 1), ws_Tasks.Cells(1, 99)))
    
    col_Project_wsTasks = fx_Create_Headers("Project", arry_Header_wsTasks)
    col_Completed_wsTasks = fx_Create_Headers("Completed", arry_Header_wsTasks)

End Sub
Sub o_24_Assign_Variables_wsWaiting()

' Purpose: To assign the Variables related to ws_Waiting.
' Trigger: Called by various procedures
' Updated: 11/5/2021

' Change Log:
'       11/5/2021: Intial Creation

' ***********************************************************************************************************************************

' ----------------------
' Assign Cell References
' ----------------------

    arry_Header_wsWaiting = Application.Transpose(ws_Waiting.Range(ws_Waiting.Cells(1, 1), ws_Waiting.Cells(1, 99)))
    
    col_Project_wsWaiting = fx_Create_Headers("Project", arry_Header_wsWaiting)
    col_Completed_wsWaiting = fx_Create_Headers("Completed", arry_Header_wsWaiting)

End Sub
Sub o_25_Assign_Variables_wsQuestions()

' Purpose: To assign the Variables related to ws_Questions.
' Trigger: Called by various procedures
' Updated: 6/27/2023

' Change Log:
'       11/21/2021: Intial Creation
'       6/27/2023:  Updated to exit sub if this is my Personal To Do

' ***********************************************************************************************************************************

#If Personal = 1 Then
    Exit Sub
#End If

' ----------------------
' Assign Cell References
' ----------------------

    arry_Header_wsQuestions = Application.Transpose(ws_Questions.Range(ws_Questions.Cells(1, 1), ws_Questions.Cells(1, 99)))
    
    col_Project_wsQuestions = fx_Create_Headers("Project", arry_Header_wsQuestions)
    col_Completed_wsQuestions = fx_Create_Headers("Completed", arry_Header_wsQuestions)

End Sub
Sub o_03_Create_Project_List()
   
' Purpose: To create the list of projects based on the selected Area.
' Trigger: Click List Area
' Updated: 6/28/2023

' Change Log:
'       4/12/2020:  Removed any reference to the D/A Worksheet, removed with my move to CR&A
'       9/3/2021:   Added code to default to just Active Projects, but include Pending / Continuous if they are checked
'       9/3/2021:   Added code to select the only Project if only one is in the list
'       9/4/2021:   Added code to default the chk_Continuous = True when selecting Area = Continuous
'       12/10/2021: Removed the code related to the P.XXX name, now that I include that in the Project Name
'       6/28/2023:  Added the 'Recurring' option with 'Continuous'

' ***********************************************************************************************************************************
        
' ----------------------
' Declare your variables
' ----------------------
   
    Dim strProjArea As String
    
    Dim strProjName As String
    
    Dim strProjStatus As String
    
    'ListBox variables
    Dim strArea As String
        If IsNull(lst_Area.Value) Then
            Else: strArea = lst_Area.Value
        End If
    
    Dim x As Long
        x = 2
    
' -----------------------------------------------
' If Area = "Continuous" then chk_Continuous = True
' -----------------------------------------------
    
    If Me.lst_Area.Value = "Continuous" Then chk_Continuous = True
    
' ------------
' Run the loop
' ------------
   
    lst_Project.Clear

    With ws_Projects
        Do While .Cells(x, col_Project_wsProjects).Value2 <> ""
            
            ' Assign the Variables
            strProjArea = .Cells(x, col_Area_wsProjects).Value2
            strProjName = .Cells(x, col_Project_wsProjects).Value2
            strProjStatus = .Cells(x, col_Status_wsProjects).Value2
    
            ' Add the values to the ListBox
            If strProjArea = strArea Then
            
                If strProjStatus = "Active" Or _
                (strProjStatus = "Pending" And chk_Pending = True) Or _
                ((strProjStatus = "Continuous" Or strProjStatus = "Recurring") And chk_Continuous = True) Then
                    lst_Project.AddItem (strProjName)
                End If
            
            End If
            
            x = x + 1
        
        Loop
    End With

' ------------------------------
' If only one remains, select it
' ------------------------------

    If lst_Project.ListCount = 1 Then
        lst_Project.Selected(0) = True
    End If

End Sub
Sub o_11_Open_Project_Folder()

' Purpose: To open the project support folder if that option is selected.
' Trigger: Called: uf_Project_Selector
' Updated: 4/12/2020

' Change Log:
'          4/12/2020: Simplified to move all the variable stuff to the o_021_Declare_Global_Variables Sub
'          4/12/2020: Simplified to combine the Projects & D/A steps together

' ***********************************************************************************************************************************

' -----------------
' Declare Variables
' -----------------

    Call Me.o_022_Declare_Project_Variables
    
    Dim strProjFldrHyperLink As String

' ----------------------------------
' Open the applicable Support Folder
' ----------------------------------

    strProjFldrHyperLink = ws_Projects.Range("C" & intProjRow).Address
        ws_Projects.Range(strProjFldrHyperLink).Hyperlinks(1).Follow

End Sub
Sub o_12_Open_Project_Support()

' Purpose: To open the P.XXX project support Excel if that option is selected.
' Trigger: Called: uf_Project_Selector
' Updated: 9/28/2022

' Change Log:
'       4/12/2020:  Simplified to move all the variable stuff to the o_021_Declare_Global_Variables Sub
'       4/12/2020:  Simplified to combine the Projects & D/A steps together
'       11/18/2021: Removed the code related to going to Next Actions
'       7/19/2022:  Updated to handle the .properties files, and switch to Hyperlink.Address
'       9/28/2022:  Overhauled to open both the .xlsx and .properties files (took code from Worksheet_BeforeDoubleClick event)

' ***********************************************************************************************************************************

Call Me.o_022_Declare_Project_Variables

' -----------------
' Declare Variables
' -----------------
        
    ' Declare File Paths
    
    Dim str_P_XXX_Support_FullPath As String
        str_P_XXX_Support_FullPath = ws_Projects.Range("D" & intProjRow).Hyperlinks(1).Address
        
    Dim str_HyperlinkAddy_xlsx As String
    
    Dim str_HyperlinkAddy_txt As String
        
    ' Assign the Hyperlinks
    str_HyperlinkAddy_xlsx = Replace(Replace(str_P_XXX_Support_FullPath, "/", "\"), ".properties", ".xlsx")
    str_HyperlinkAddy_txt = Replace(Replace(str_P_XXX_Support_FullPath, "/", "\"), ".xlsx", ".properties")
    
    ' Fix the Hyperlinks
    #If Personal <> 1 Then
        If Left(str_HyperlinkAddy_xlsx, 11) <> "C:\U Drive\" Then str_HyperlinkAddy_xlsx = "C:\U Drive\" & str_HyperlinkAddy_xlsx
        If Left(str_HyperlinkAddy_txt, 11) <> "C:\U Drive\" Then str_HyperlinkAddy_txt = "C:\U Drive\" & str_HyperlinkAddy_txt
    #Else
        If Left(str_HyperlinkAddy_xlsx, 15) <> "D:\D Documents\" Then str_HyperlinkAddy_xlsx = "D:\D Documents\" & str_HyperlinkAddy_xlsx
        If Left(str_HyperlinkAddy_txt, 15) <> "D:\D Documents\" Then str_HyperlinkAddy_txt = "D:\D Documents\" & str_HyperlinkAddy_txt
    #End If
    
' -------------------------------
' Open the Project Support File/s
' -------------------------------
    
    If fx_File_Exists(str_HyperlinkAddy_xlsx) Then
        ThisWorkbook.FollowHyperlink (str_HyperlinkAddy_xlsx)
    End If
    
    If fx_File_Exists(str_HyperlinkAddy_txt) Then
        Call Shell("explorer.exe" & " " & str_HyperlinkAddy_txt)
    End If
    
End Sub
Sub o_21_Filter_wsTasks()

' Purpose: To filter ws_Tasks based on the selected project.
' Trigger: Called: uf_Project_Selector
' Updated: 11/5/2021

' Change Log:
'       4/12/2020: Simplified to move all the variable stuff to the o_021_Declare_Global_Variables Sub
'       4/12/2020: Replaced the middle filter to use ProjName instead of the ProjNum
'       11/5/2021: Added explicit cell references for ws_Tasks and ws_Waiting

' ***********************************************************************************************************************************

Call Me.o_23_Assign_Variables_wsTasks

' -------------------------------------------------
' Filter the Tasks tab based on the seleced project
' -------------------------------------------------

    With ws_Tasks

      .AutoFilter.ShowAllData
          .Range("A1").AutoFilter Field:=col_Project_wsTasks, Criteria1:=strProjName
          .Range("A1").AutoFilter Field:=col_Completed_wsTasks, Criteria1:="=", Operator:=xlFilterValues
          
    End With

End Sub
Sub o_22_Filter_wsWaiting()

' Purpose: To filter ws_Waiting based on the selected project.
' Trigger: Called: uf_Project_Selector
' Updated: 11/5/2021

' Change Log:
'       4/12/2020: Simplified to move all the variable stuff to the o_021_Declare_Global_Variables Sub
'       4/12/2020: Replaced the middle filter to use ProjName instead of the ProjNum
'       11/5/2021: Added explicit cell references for ws_Tasks and ws_Waiting

' ***********************************************************************************************************************************

Call Me.o_24_Assign_Variables_wsWaiting

' ---------------------------------------------------
' Filter the Waiting tab based on the seleced project
' ---------------------------------------------------

    With ws_Waiting
    
        .AutoFilter.ShowAllData
            .Range("A1").AutoFilter Field:=col_Project_wsWaiting, Criteria1:=strProjName
            .Range("A1").AutoFilter Field:=col_Completed_wsWaiting, Criteria1:="=", Operator:=xlFilterValues
            
    End With

End Sub
Sub o_23_Filter_wsQuestions()

' Purpose: To filter ws_Questions based on the selected project.
' Trigger: Called: uf_Project_Selector
' Updated: 6/27/2023

' Change Log:
'       11/21/2021: Initial Creation
'       6/27/2023:  Updated to exit sub if this is my Personal To Do

' ***********************************************************************************************************************************
                
#If Personal = 1 Then
    Exit Sub
#End If

Call Me.o_25_Assign_Variables_wsQuestions

' ---------------------------------------------------
' Filter the Waiting tab based on the seleced project
' ---------------------------------------------------

    With ws_Questions
    
        .AutoFilter.ShowAllData
            .Range("A1").AutoFilter Field:=col_Project_wsQuestions, Criteria1:=strProjName
            .Range("A1").AutoFilter Field:=col_Completed_wsQuestions, Criteria1:="=", Operator:=xlFilterValues
            
    End With

End Sub
Sub o_31_Open_Flowchart()

' Purpose: To open the Flowchart for the selected project.
' Trigger: Called: uf_Project_Selector
' Updated: 1/31/2022

' Change Log:
'       9/10/2021:  Initial Creation
'       1/31/2022:  Updated to get this code working

' ***********************************************************************************************************************************

    Call Me.o_022_Declare_Project_Variables

' -----------------
' Declare Variables
' -----------------
    
    Dim strProjFldrHyperLink As String
        strProjFldrHyperLink = ws_Projects.Range("C" & intProjRow).Address
        
    Dim strProjectPath As String
        strProjectPath = "C:\U Drive\" & ws_Projects.Range(strProjFldrHyperLink).Hyperlinks(1).Address & "\"
        
    Dim strFlowchartPath As String
    
' --------------------------------
' Open the applicable Support File
' --------------------------------

        strFlowchartPath = fx_Get_Most_Recent_File_From_Directory_Based_On_Modified_Date( _
                        strFolderPath:=strProjectPath, _
                        strFileExtension:=".vsd")
        
        Call Shell("explorer.exe" & " " & strFlowchartPath, vbNormalFocus)
        
End Sub
Sub o_32_Open_Project_Workbook()

' Purpose: To open the Project Workbook for the selected project.
' Trigger: Called: uf_Project_Selector
' Updated: 1/31/2022

' Change Log:
'       1/31/2022:  Initial Creation

' ***********************************************************************************************************************************

    Call Me.o_022_Declare_Project_Variables

' -----------------
' Declare Variables
' -----------------
    
    Dim strProjFldrHyperLink As String
        strProjFldrHyperLink = ws_Projects.Range("C" & intProjRow).Address
    
    Dim strProjectPath As String
        strProjectPath = "C:\U Drive\" & ws_Projects.Range(strProjFldrHyperLink).Hyperlinks(1).Address & "\"
        
    Dim strFlowchartPath As String
    
    Dim collTEMP As New Collection
    
' --------------------------------
' Open the applicable Support File
' --------------------------------
        
    Set collTEMP = fx_List_Files_In_Folder(strProjectPath)
        
    Call Shell("explorer.exe" & " " & strFlowchartPath, vbNormalFocus)
        
' ----------------------
' Declare your variables
' ----------------------
    
    'Dim Strings
    Dim strDirParent As String
        #If Personal = 0 Then
            strDirParent = "C:\U Drive\"
        #ElseIf Personal = 1 Then
            strDirParent = "D:\D Documents\"
        #End If
    
    Dim strcurFolder As String
    
    'Dim Objects
    Dim objFSO As Object
        Set objFSO = CreateObject("Scripting.FileSystemObject")
    
    Dim objDirParent As Object
        Set objDirParent = objFSO.GetFolder(strDirParent)
    
    Dim objSubFolder As Object
    
    Dim objFile As Object

' ------------------------------------
' Load the folders into the dictionary
' ------------------------------------

        Do Until strcurFolder = ""

            If (GetAttr(strDirParent & strcurFolder) And vbDirectory) = vbDirectory Then
                'dict_Folder.Add Key:=strDirParent & strcurFolder, Item:=strDirParent & strcurFolder
                coll_Folders.Add strDirParent & strcurFolder
            End If

            strcurFolder = Dir()

        Loop

    'Recur through each folder

    For Each objSubFolder In objDirParent.SubFolders

    strcurFolder = Dir(objSubFolder.path & "\", vbDirectory)

        Do Until strcurFolder = ""

                If (GetAttr(objSubFolder.path & "\" & strcurFolder) And vbDirectory) = vbDirectory Then
                    'dict_Folder.Add Key:=objSubFolder.Path & "\" & strcurFolder, Item:=objSubFolder.Path & "\" & strcurFolder
                    coll_Folders.Add objSubFolder.path & "\" & strcurFolder
                End If

            strcurFolder = Dir()
        Loop

    Next
        
End Sub

