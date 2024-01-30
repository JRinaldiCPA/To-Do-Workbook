VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} uf_Perspectives 
   Caption         =   "Perspectives"
   ClientHeight    =   2616
   ClientLeft      =   132
   ClientTop       =   756
   ClientWidth     =   8760.001
   OleObjectBlob   =   "uf_Perspectives.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "uf_Perspectives"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Dim Sheets
    Dim ws_Projects As Worksheet
    Dim ws_Tasks As Worksheet
    Dim ws_Waiting As Worksheet

'Dim Strings
    Dim strProjName As String
    Dim strProjArea As String

'Dim Integers
    Dim intProjNum As Long
    Dim intProjRow As Long



Private Sub cmd_Create_GTD_Review_Support_Click()

End Sub
Private Sub cmd_Review_Work_Completed_Click()

    Call Macros.o_55_Weekly_GTD_Review
    Unload Me

End Sub
Private Sub cmd_Mark_Work_Completed_Click()

    'Call Me.o_21_Mark_Work_Completed
    Call Me.o_22_Mark_Work_Completed_Weekly_Plan

End Sub
Private Sub cmd_Reset_GTD_Weekly_Click()

    Call Macros.o_52_Reset_To_Do_Weekly
    Unload Me

End Sub
Private Sub cmd_Create_Daily_ToDos_Click()
    
    Call Macros.o_56_Create_Daily_To_Do_txt_For_Upcoming_Week
    
End Sub
Private Sub cmd_Refresh_Weekly_Plan_Click()

'   Loop through the tasks for the selected project, wherever one is complete remove from Weekly Review


End Sub
Private Sub lst_Project_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    If KeyCode = vbKeyReturn And Me.lst_Project.Value <> "" Then
        Call Me.o_11_Open_Project_Support
        Unload Me
    End If

End Sub
Private Sub lst_Area_Click()
    
    Call Me.o_03_Create_Project_List

End Sub
Private Sub cmb_DynamicSearch_Change()

' Purpose: To replace the values in the Projects ListBox based on what is typed.
' Trigger: Called by uf_Project_Selector
' Updated: 12/10/2021

' Change Log:
'       4/12/2020: Removed any reference to the D/A Worksheet, removed with my move to CR&A
'       12/10/2021: Removed the code related to the P.XXX name, now that I include that in the Project Name

' ***********************************************************************************************************************************

' -----------
' Declare your variables
' -----------

    Dim strProjArea As String
        
    Dim strProjName As String
    
    Dim strProjStatus As String
    
    Dim x As Long
    
    Dim y As Long
        y = 1
    
    Dim ary_Projects As Variant
        ReDim ary_Projects(1 To 999)
    
' -----------
' Run the loop
' -----------
       
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

End Sub
Private Sub lst_Project_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    
    Call Me.o_11_Open_Project_Support

End Sub

Private Sub cmd_Cancel_Click()
    
    Unload Me

End Sub
Sub o_02_Declare_Global_Variables()

' Purpose: To set the variables in ONE location for the rest of the subs.
' Trigger: Called: uf_Project_Selector
' Updated: 12/10/2021

' Change Log:
'       12/10/2021: Removed the code related to the P.XXX name, and bolDARequest, now that I include that in the Project Name


' ***********************************************************************************************************************************

' -----------
' Declare your variables
' -----------
    
    'Dim strProjName As String
        strProjName = lst_Project.Value
    
    'Dim strProjArea As String
        If IsNull(lst_Area.Value) Then
            strProjArea = ws_Projects.Range("C" & Application.WorksheetFunction.Match(Me.lst_Project.Value, ws_Projects.Columns(4), 0))
        Else
            strProjArea = lst_Area.Value
        End If
    
    'Dim intProjRow as long
        intProjRow = Application.WorksheetFunction.Match(Me.lst_Project.Value, ws_Projects.Columns(4), 0)

End Sub
Sub o_03_Create_Project_List()
   
' Purpose: To create the list of projects based on the selected Area.
' Trigger: Called: uf_Project_Selector
' Updated: 12/10/2021

' Change Log:
'       9/24/2019:  Initial Creation
'       12/10/2021: Removed the code related to the P.XXX name, now that I include that in the Project Name

' ***********************************************************************************************************************************
        
Call myPrivateMacros.DisableForEfficiency

' -----------
' Declare your variables
' -----------
   
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
    
' -----------
' Run the loop
' -----------
   
    lst_Project.Clear

    With ws_Projects
        Do While .Range("A" & x).Value2 <> ""
            
            strProjArea = .Range("C" & x).Value2
            strProjName = .Range("D" & x).Value2
            strProjStatus = .Range("G" & x).Value2
    
            If (strProjStatus = "Active" Or strProjStatus = "Pending" Or strProjStatus = "Continuous" Or strProjStatus = "Recurring") _
            And strProjArea = strArea Then
                lst_Project.AddItem (strProjName)
            End If
            
            x = x + 1
        
        Loop
    End With

Call myPrivateMacros.DisableForEfficiencyOff
              
End Sub
Sub o_11_Open_Project_Support()

' Purpose: To open the P.XXX project support Excel if that option is selected.
' Trigger: Called: uf_Project_Selector
' Updated: 7/19/2021

' Change Log:
'       7/19/2021: Initial Creation
'       7/19/2021: Added code to close the workbook if it is already open
'       11/18/2021: Removed the code related to 'Next Actions'

' ***********************************************************************************************************************************

' -----------
' Declare your variables
' -----------
       
    Call Me.o_02_Declare_Global_Variables
    
    Dim strProjSupportHyperLink As String
    
    Dim bolSupportOpenAlready As Boolean
    bolSupportOpenAlready = fx_Sheet_Exists( _
        strWBName:=strProjName & ".xlsx", _
        strWsName:="Notes")
    
' -----------
' Open the Support Workbook, or close if already open
' -----------

    If bolSupportOpenAlready = False Then
        strProjSupportHyperLink = ws_Projects.Range("D" & intProjRow).Address
            ws_Projects.Range(strProjSupportHyperLink).Hyperlinks(1).Follow
    Else
        Workbooks(strProjName & ".xlsx").Close SaveChanges:=True
    End If
        
End Sub
Sub o_21_Mark_Work_Completed()

' Purpose: To mark completed Tasks / Waiting in my P.XXX / DA.XXX Support.
' Trigger: Called: cmd_Mark_Work_Completed_Click
' Updated: 7/26/2021

' Change Log:
'       7/26/2021: Initial Creation

' ***********************************************************************************************************************************

' ----------------------
' Declare your variables
' ----------------------

    Dim intSelectedRow As Long
        intSelectedRow = ActiveCell.Row

    'Dim wbSupport As Workbook
    'Set wbSupport = fx_Open_Support_Workbook
    
    Dim int_LastRow_Support As Long

    Dim i As Integer
    
    Dim dict_Support As Scripting.Dictionary
    
' --------------------------------------
' Create the Dictionary of Support Tasks
' --------------------------------------

    ' Add to the dictionary using the row and the tasks descriptions
    
'With wbSupport.Sheets("Next Actions")
'
'    int_LastRow_Support = .Cells(Rows.count, "A").End(xlUp).Row
'
'    For i = 2 To int_LastRow_Support
'
'        dict_Support.Add Key:=i, Item:=.Cells(i, "D").Value
'
'    Next i
'
'End With
    
    ' COmpare that to the list from the To Do, with the same Row and Task


Stop

' 1) Select a task from Tasks / Waiting
' 2) Hit the button to mark it as complete in:
'   1) P.XXX Support
'   2) To Do (if not already)
'   3) Weekly Plan
'   4) Current Work (?)


End Sub
Sub o_22_Mark_Work_Completed_Weekly_Plan()

' Purpose: To mark completed Tasks / Waiting in my P.XXX / DA.XXX Support.
' Trigger: Called: cmd_Mark_Work_Completed_Click
' Updated: 8/12/2021

' Change Log:
'       7/26/2021: Initial Creation
'       8/12/2021: Overhauled, moving all of the current file related work to Functions

' ***********************************************************************************************************************************

' ----------------------
' Declare your variables
' ----------------------

    ' Declare Text Variables

    Dim txtWeeklyReview As Long
        txtWeeklyReview = FreeFile

    Dim strWeeklyReviewPath As String
        strWeeklyReviewPath = fx_Get_Most_Recent_File_From_Directory_Based_On_File_Name(strFolderPath:="C:\U Drive\Support\Weekly Plan\")
        
    Dim strSelectedTask As String
        strSelectedTask = "Review the flowchart that Rachael created and update mine to reflect the steps to create the Adjusted Loan Trial"
        
    Dim strNewContent As String
        strNewContent = "#"
        
    Dim intSec1 As Long
    
    Dim intSec2 As Long
        
    ' Declare Integers
    
    Dim int_LastRow As Long
        int_LastRow = ws_Tasks.Cells(Rows.count, "A").End(xlUp).Row

    ' Declare Dictionaries
    
    Dim dict_Tasks As Scripting.Dictionary
    Set dict_Tasks = New Scripting.Dictionary

    ' Declare Loop Variables
    
    Dim i As Integer

' --------------------------------------
' Create the Dictionary of Support Tasks
' --------------------------------------

    With ws_Tasks
        
        For i = 2 To int_LastRow
            dict_Tasks.Add key:=i, Item:=.Cells(i, "J").Value
        Next i
        
    End With

' ----------------------------------------------
' Open the Weekly Review and add the new content
' ----------------------------------------------

    'Open the text file in Read Only mode to pull the current content
        Open strWeeklyReviewPath For Input As txtWeeklyReview
            FileContent = Input(LOF(txtWeeklyReview), txtWeeklyReview)
            intSec1 = InStr(1, FileContent, strSelectedTask, vbTextCompare) - 6
            intSec2 = LOF(txtWeeklyReview) - (intSec1)
            
            Debug.Print "Length of File:      " & LOF(txtWeeklyReview)
            Debug.Print "Length of Sections: "; intSec1 + intSec2
        Close txtWeeklyReview
        'Reset
      
    'Create the new string
        FileContent = Left(FileContent, intSec1) & strNewContent & Right(FileContent, intSec2)

    'Open the text file in a Write mode to add the new content
        Open strWeeklyReviewPath For Output As txtWeeklyReview
            Print #txtWeeklyReview, FileContent
        Close txtWeeklyReview



End Sub


