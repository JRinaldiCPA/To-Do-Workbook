VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} uf_RefX_and_wsDaily_Update 
   Caption         =   "RefX Update / Selector"
   ClientHeight    =   7528
   ClientLeft      =   -48
   ClientTop       =   36
   ClientWidth     =   9264.001
   OleObjectBlob   =   "uf_RefX_and_wsDaily_Update.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "uf_RefX_and_wsDaily_Update"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public strTextBoxValue As String
Public int_lst_TxtFile_width As Integer
Public int_frm_UpdateOptions_width As Integer
Public int_UserForm_width As Integer
Public str_UpdateType As String
Private Sub UserForm_Initialize()
        
' Purpose: To initialize the userform, including adding in the data from the arrays and pulling data from the current row if a Project is selected.
' Trigger: Keyboard Shortcut: Ctrl + Shift + M
        
' Updated: 8/1/2023

' Change Log:
'       12/20/2019: Added a new button to Add Today, which will directly add the copied info under the current day in Daily
'       12/20/2019: Updated the Initiailze step to pull in the Project ID if I am in Projects
'       11/28/2021: Removed the Agendas option
'       5/16/2022:  Updated how the Individual RefX gets passed, now it uses the fx_List_Files_In_Folder function
'                   Removed the opt_Frameworks and added it into the Updates section
'       7/5/2022:   Combined the RefX_Update and Daily_Update
'       7/15/2022:  Added the code to pull the value from the cell if in Temp
'       8/1/2023:   Added the str_UpdateType to capture the type of update
'                   Moved the code to manipulate the form based on the option selected into it's own sub
   
' ***********************************************************************************************************************************
        
' -----------------------------------
' Tailor if it's my personal computer
' -----------------------------------

#If Personal = 1 Then

    Me.opt_Individual.Visible = False
    Me.opt_Performance.Visible = False
    Me.opt_RefX.Visible = False
    Me.opt_Updates.Visible = False
    
    Me.opt_Personal_Individual.Visible = True
    Me.opt_Personal_RefX.Visible = True
    Me.opt_Personal_Updates.Visible = True

    ' Move them to their "normal" position
    Me.opt_Personal_Individual.Left = 6
    Me.opt_Personal_RefX.Left = 6
    Me.opt_Personal_Updates.Left = 6

#End If
    
' -------------------
' Fill initial values
' -------------------
    
    Call Me.o_13_Create_Daily_Categories_List
    
    int_UserForm_width = Me.Width
    int_frm_UpdateOptions_width = Me.frm_Update_Options.Width
    int_lst_TxtFile_width = Me.lst_TxtFile.Width
    
' ----------------------------------------------------------------------------------------------------
' If the selection has a value copy paste as values only, otherwise paste the value from the clipboard
' ----------------------------------------------------------------------------------------------------
              
    Dim objDataObj As New MSForms.DataObject
    
    If ActiveSheet.Name = "Projects" Then
        Call myFunctions_ToDo.fx_Copy_Project_ID_to_Clipboard
        objDataObj.GetFromClipboard
        
        strTextBoxValue = Trim(objDataObj.GetText(1))
    ElseIf ActiveSheet.Name = "Temp" And ActiveCell.Value <> "" Then
        strTextBoxValue = ActiveCell.Value2
    Else
        strTextBoxValue = myFunctions_ToDo.fx_Copy_to_Clipboard
    End If

    Me.txt_Text_to_Append.Value = strTextBoxValue
    'Me.txt_Text_to_Append.Value = strTextBoxValue

End Sub
Private Sub cmd_Append_Text_Click()

' Purpose: To add the passed text to the Text file selected.

' Updated: 12/20/2019

' Change Log:
'       12/20/2019: Initial Creation (?)

' ***********************************************************************************************************************************

'Check if txt was added
    If Me.txt_Text_to_Append.Value = "" Then
        MsgBox "There was no note to add"
        Unload Me
    End If

'If a section was selected, add the content under the section, otherwise append
    If Me.lst_TxtFile_Section <> "" Then
        Call o_22_Write_To_Txt
    Else
        Call o_21_Append_To_Txt
    End If

    Unload Me

End Sub
Private Sub cmd_OpenTxt_Click()

' Purpose: To open the selected Text file.

' Updated: 10/30/2021

' Change Log:
'       12/20/2019: Initial Creation (?)
'       10/30/2021: Added the check for the .lst_TextFile

' ***********************************************************************************************************************************

    'If IsNull(Me.lst_TxtFile) = True Then
    If Me.lst_TxtFile.Value = "" Then
        
        Unload Me
        uf_Search_RefX.Show
        
    Else
        Call Me.o_23_Open_Txt
            Unload Me
    End If

End Sub
Private Sub cmd_DailyUpdate_Click()
        
Call myPrivateMacros.DisableForEfficiency

'Check if txt was added
    If Me.txt_Text_to_Append.Value = "" Then
        MsgBox "There was no note to add"
        Unload Me
    End If

    Call o_3_Append_To_Daily
        Unload Me

Call myPrivateMacros.DisableForEfficiencyOff
        
End Sub
Private Sub cmd_Cancel_Click()
    
    Unload Me

End Sub
Private Sub opt_Performance_Click()
    
' Purpose: To populate the files in the lst_TxtFile.
' Trigger: Called: uf_RefX_and_wsDaily_Update
' Updated: 8/1/2023

' Change Log:
'       9/16/2019:  Initial Creation
'       5/16/2022:  Added the code to clear the Sub-Section
'       7/5/2022:   Updated to hide the lst_Daily_Category
'       7/15/2022:  Added the code to hide the inapplicable cmd buttons
'       8/1/2023:   Added the code to capture the "Update Type"

' ***********************************************************************************************************************************
    
    str_UpdateType = "Staff Performance"
    
    Call Me.o_14_Tailor_The_UserForm
    
    Call Me.o_11_Create_Text_File_List
        If Me.opt_Performance.Value = True Then Me.lst_TxtFile.Value = "JAMES"

End Sub
Private Sub opt_Individual_Click()
    
' Purpose: To populate the files in the lst_TxtFile.
' Trigger: Called: uf_RefX_and_wsDaily_Update
' Updated: 8/1/2023

' Change Log:
'       9/16/2019:  Initial Creation
'       5/16/2022:  Added the code to clear the Sub-Section
'       7/5/2022:   Updated to hide the lst_Daily_Category and lst_txtFile_Section
'       7/15/2022:  Added the code to hide the inapplicable cmd buttons
'       8/1/2023:   Added the code to capture the "Update Type"

' ***********************************************************************************************************************************
    
    str_UpdateType = "Individual RefX"
    
    Call Me.o_14_Tailor_The_UserForm
    
    Call Me.o_11_Create_Text_File_List
        Me.lst_TxtFile_Section.Clear
    
End Sub
Private Sub opt_RefX_Click()
    
' Purpose: To populate the files in the lst_TxtFile.
' Trigger: Called: uf_RefX_and_wsDaily_Update
' Updated: 8/1/2023

' Change Log:
'       9/16/2019:  Initial Creation
'       5/16/2022:  Added the code to clear the Sub-Section
'       7/5/2022:   Updated to hide the lst_Daily_Category and lst_txtFile_Section
'       7/15/2022:  Added the code to hide the inapplicable cmd buttons
'       8/1/2023:   Added the code to capture the "Update Type"

' ***********************************************************************************************************************************
    
    str_UpdateType = "RefX"
    
    Call Me.o_14_Tailor_The_UserForm
    
    Call Me.o_11_Create_Text_File_List
        Me.lst_TxtFile_Section.Clear
    
End Sub
Private Sub opt_Updates_Click()

' Purpose: To populate the files in the lst_TxtFile.
' Trigger: Called: uf_RefX_and_wsDaily_Update
' Updated: 8/1/2023

' Change Log:
'       9/16/2019:  Initial Creation
'       5/16/2022:  Added the code to clear the Sub-Section
'       7/5/2022:   Updated to hide the lst_Daily_Category and lst_txtFile_Section
'       7/15/2022:  Added the code to hide the inapplicable cmd buttons
'       8/1/2023:   Added the code to capture the "Update Type"

' ***********************************************************************************************************************************
    
    str_UpdateType = "Updates"
    
    Call Me.o_14_Tailor_The_UserForm
    
    Call Me.o_11_Create_Text_File_List
        Me.lst_TxtFile_Section.Clear
    
End Sub
Private Sub opt_Projects_Click()

' Purpose: To populate the files in the lst_TxtFile.
' Trigger: Called: uf_RefX_and_wsDaily_Update
' Updated: 8/1/2023

' Change Log:
'       5/16/2022:  Initial creation
'       7/5/2022:   Updated to hide the lst_Daily_Category, and adjust the object widths
'       7/15/2022:  Added the code to hide the inapplicable cmd buttons
'       8/1/2023:   Added the code to capture the "Update Type"

' ***********************************************************************************************************************************
    
    str_UpdateType = "Projects"
    
    Call Me.o_14_Tailor_The_UserForm
    
    Call Me.o_11_Create_Text_File_List
        Me.lst_TxtFile_Section.Clear
       
End Sub
Private Sub opt_wsDaily_Click()

' Purpose: To show the data in lst_Daily_Category.
' Trigger: Called: uf_RefX_and_wsDaily_Update
' Updated: 8/1/2023

' Change Log:
'       7/5/2022:   Initial Creation
'       7/15/2022:  Added the code to hide the inapplicable cmd buttons
'                   Added code so that if a selection was already made it doesn't recreate the list
'       8/1/2023:   Added the code to capture the "Update Type"

' ***********************************************************************************************************************************

    str_UpdateType = "ws_Daily"

    Call Me.o_14_Tailor_The_UserForm

    If Me.lst_Daily_Category.Value = "" Then
        Call Me.o_13_Create_Daily_Categories_List
    End If

End Sub
Private Sub opt_Personal_Individual_Click()
    
' Purpose: To populate the files in the lst_TxtFile.
' Trigger: Called: uf_RefX_and_wsDaily_Update
' Updated: 7/5/2022

' Change Log:
'       9/16/2019:  Initial Creation
'       5/16/2022:  Added the code to clear the Sub-Section
'       7/5/2022:   Updated to hide the lst_Daily_Category

' ***********************************************************************************************************************************
    
    Call Me.o_11_Create_Text_File_List
    
    Me.lst_TxtFile.Visible = True
    Me.lst_Daily_Category.Visible = False
    
    Me.lst_TxtFile_Section.Clear
        
End Sub
Private Sub opt_Personal_RefX_Click()
    
' Purpose: To populate the files in the lst_TxtFile.
' Trigger: Called: uf_RefX_and_wsDaily_Update
' Updated: 7/5/2022

' Change Log:
'       9/16/2019:  Initial Creation
'       5/16/2022:  Added the code to clear the Sub-Section
'       7/5/2022:   Updated to hide the lst_Daily_Category

' ***********************************************************************************************************************************
    
    Call Me.o_11_Create_Text_File_List
    
    Me.lst_TxtFile.Visible = True
    Me.lst_Daily_Category.Visible = False
    
    Me.lst_TxtFile_Section.Clear
    
End Sub
Private Sub opt_Personal_Updates_Click()
    
' Purpose: To populate the files in the lst_TxtFile.
' Trigger: Called: uf_RefX_and_wsDaily_Update
' Updated: 7/5/2022

' Change Log:
'       9/16/2019:  Initial Creation
'       5/16/2022:  Added the code to clear the Sub-Section
'       7/5/2022:   Updated to hide the lst_Daily_Category

' ***********************************************************************************************************************************
    
    Call Me.o_11_Create_Text_File_List
    
    Me.lst_TxtFile.Visible = True
    Me.lst_Daily_Category.Visible = False
    
    Me.lst_TxtFile_Section.Clear
    
End Sub
Private Sub lst_TxtFile_Click()

' Purpose: To populate the Sub-Sections based on the selected text file.
' Trigger: Called: uf_RefX_and_wsDaily_Update
' Updated: 5/16/2022

' Change Log:
'       9/16/2019:  Initial Creation
'       5/16/2022:  Removed the code to clear out the Sub-Section, it was redundant

' ***********************************************************************************************************************************

    Call o_12_Create_SubSections_List

End Sub
Private Sub lst_TxtFile_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    'If IsNull(Me.lst_TxtFile) = True Then
    If Me.lst_TxtFile.Value = "" Then
        
        Unload Me
        uf_Search_RefX.Show
        
    Else
        Call Me.o_23_Open_Txt(bolOpenwNotepadPP:=True)
            Unload Me
    End If

End Sub
Private Sub lst_Daily_Category_Click()

' Purpose: To hide / show the applicable objects.
' Trigger: Called: uf_RefX_and_wsDaily_Update
' Updated: 7/15/2022

' Change Log:
'       7/15/2022:   Initial Creation

' ***********************************************************************************************************************************

    Me.opt_wsDaily.Value = True

End Sub
Sub o_11_Create_Text_File_List()

' Purpose: To add items to the Text File List from the TextFile Array.
' Trigger: Called: uf_RefX_and_wsDaily_Update
' Updated: 9/16/2019

' ***********************************************************************************************************************************

    Me.lst_TxtFile.Clear
        Me.lst_TxtFile.List = fx_GetTextFileArray

End Sub
Sub o_12_Create_SubSections_List()

' Purpose: To add items to the Section List from the SubSection Array.
' Trigger: Called: uf_RefX_and_wsDaily_Update
' Updated: 9/16/2019

' ***********************************************************************************************************************************

    Me.lst_TxtFile_Section.Clear
        Me.lst_TxtFile_Section.List = fx_GetSubSectionArray

End Sub
Sub o_13_Create_Daily_Categories_List()

' Purpose: To add items to the Daily Categories list from the Category array.
' Trigger: Called: uf_RefX_and_wsDaily_Update
' Updated: 9/16/2019

' ***********************************************************************************************************************************

    Me.lst_Daily_Category.Clear
        Me.lst_Daily_Category.List = fx_Getws_DailyCategoryArray

End Sub
Sub o_14_Tailor_The_UserForm()

' Purpose: To tailor the UserForm based on which option was selected.
' Trigger: Called: uf_RefX_and_wsDaily_Update
' Updated: 8/1/2023

' Change Log:
'       8/1/2013:   Initial Creation

' ***********************************************************************************************************************************

' ---------------------------------
' Adjust the UserForm Object Widths
' ---------------------------------

    If str_UpdateType = "Projects" Then
 
        Me.Width = int_UserForm_width + 30
        Me.frm_Update_Options.Width = int_frm_UpdateOptions_width + 30
        Me.lst_TxtFile.Width = int_lst_TxtFile_width + 200
    
    Else
    
        Me.Width = int_UserForm_width
        Me.frm_Update_Options.Width = int_frm_UpdateOptions_width
        Me.lst_TxtFile.Width = int_lst_TxtFile_width
    
    End If
    
' -------------------------------------
' Adjust what's visible on the UserForm
' -------------------------------------
    
    If str_UpdateType = "Staff Performance" Then
        
        Me.lst_TxtFile.Visible = True
        Me.lst_Daily_Category.Visible = False
        Me.lst_TxtFile_Section.Visible = True
    
    ElseIf str_UpdateType = "Individual RefX" Then
        
        Me.lst_TxtFile.Visible = True
        Me.lst_Daily_Category.Visible = False
        Me.lst_TxtFile_Section.Visible = False
        
    ElseIf str_UpdateType = "RefX" Then
        
        Me.lst_TxtFile.Visible = True
        Me.lst_Daily_Category.Visible = False
        Me.lst_TxtFile_Section.Visible = False
    
    ElseIf str_UpdateType = "Updates" Then
    
        Me.lst_TxtFile.Visible = True
        Me.lst_Daily_Category.Visible = False
        Me.lst_TxtFile_Section.Visible = True
    
    ElseIf str_UpdateType = "Projects" Then

        Me.lst_TxtFile.Visible = True
        Me.lst_TxtFile_Section.Visible = False
        Me.lst_Daily_Category.Visible = False
        
    ElseIf str_UpdateType = "ws_Daily" Then
        
        Me.lst_TxtFile.Visible = False
        Me.lst_TxtFile_Section.Visible = False
        Me.lst_Daily_Category.Visible = True
    
    End If

' --------------------------------------------------
' Hide the unrelated Command Buttons on the UserForm
' --------------------------------------------------

    If str_UpdateType = "ws_Daily" Then
        
        Me.cmd_Append_Text.Visible = False
        Me.cmd_OpenTxt.Visible = False
        Me.cmd_DailyUpdate.Visible = True
        
    Else
        
        Me.cmd_Append_Text.Visible = True
        Me.cmd_OpenTxt.Visible = True
        Me.cmd_DailyUpdate.Visible = False
    
    End If

End Sub

Sub o_21_Append_To_Txt()

' Purpose: To allow me to quickly append to my text files.
' Trigger: Called: uf_RefX_and_wsDaily_Update
' Updated: 9/16/2019

' ***********************************************************************************************************************************

' ----------------------
' Declare your variables
' ----------------------

    Dim intTxtFile As Long
        intTxtFile = FreeFile
    
    Dim strFileName As String
        strFileName = Me.lst_TxtFile.Value
    
    Dim FilePath As String
        FilePath = fx_OutputFile(strFileName)
        
' -----------------------
' Append to the .txt file
' -----------------------
        
    Open FilePath For Append As intTxtFile
        Print #intTxtFile, vbCrLf & vbCrLf & Date & ": " & Me.txt_Text_to_Append.Value;
    Close intTxtFile
  
End Sub
Sub o_22_Write_To_Txt()

' Purpose: To allow me to quickly add text in a specific section of my text files.
' Trigger: Called: uf_RefX_and_wsDaily_Update
' Updated: 9/16/2019

' ***********************************************************************************************************************************

' ----------------------
' Declare your variables
' ----------------------

    Dim intTxtFile As Long
        intTxtFile = FreeFile
    
    Dim strFileName As String
        strFileName = Me.lst_TxtFile.Value
        
    Dim FilePath As String
        FilePath = fx_OutputFile(strFileName)
    
    Dim FileContent As String
    
    Dim strSelectedTask As String
        strSelectedTask = Me.lst_TxtFile_Section.Value
        
    Dim strNewContent As String
        strNewContent = Me.txt_Text_to_Append.Value
        
    Dim intSec1 As Long
    
    Dim intSec2 As Long
    
' -------------------------------------
' Add the new content under the section
' -------------------------------------

    'Open the text file in Read Only mode to pull the current content
        Open FilePath For Input As intTxtFile
            FileContent = Input(LOF(intTxtFile), intTxtFile)
            intSec1 = InStr(1, FileContent, strSelectedTask, vbTextCompare) + Len(strSelectedTask) - 1
            intSec2 = LOF(intTxtFile) - intSec1
        Close intTxtFile
      
    'Create the new string
        FileContent = Left(FileContent, intSec1) & vbCrLf & Date & ": " & strNewContent & Right(FileContent, intSec2)

    'Open the text file in a Write mode to add the new content
        Open FilePath For Output As intTxtFile
            Print #intTxtFile, FileContent
        Close intTxtFile

End Sub
Sub o_23_Open_Txt(Optional bolOpenwNotepadPP As Boolean)

' Purpose: To allow me to quickly open my text files.
' Trigger: Called: uf_RefX_and_wsDaily_Update
' Updated: 6/23/2022

' Change Log:
'       9/16/2019:  Initial Creation
'       6/2/2021:   Updated to replace Shell with FollowHyperlink
'       8/17/2021:  Switched to opening file with NotepadPP in Shell, for Properties files
'       6/23/2022:  Updated the code when opening a .properties file

' ***********************************************************************************************************************************

' -----------------
' Declare Variables
' -----------------

    Dim intTxtFile As Long
        intTxtFile = FreeFile
    
    Dim strFileName As String
        strFileName = Me.lst_TxtFile.Value
    
    Dim strFilePath As String
        strFilePath = fx_OutputFile(strFileName)
        
' ------------------
' Open the .txt file
' ------------------
    
    
    If Right(strFilePath, 11) = ".properties" Then
        Call Shell("explorer.exe" & " " & strFilePath)
    ElseIf bolOpenwNotepadPP = True Then
        ThisWorkbook.FollowHyperlink (strFilePath)
    Else
        ThisWorkbook.FollowHyperlink (strFilePath)
    End If

End Sub
Sub o_3_Append_To_Daily()

' Purpose: To allow me to quickly add text to my To Do - Daily sheet.
' Trigger: Called: uf_RefX_and_wsDaily_Update
' Updated: 7/5/2022

' Change Log:
'       3/13/2020: Changed int_Category_Location to skip the day listed under Daily Goals
'       4/7/2022:  Updated the intBlankRows for opt_GTD to be 8 rows instead of 6
'       7/5/2022:  Updated to remove the option buttons, now that I don't do the Daily, and all the related code

' ***********************************************************************************************************************************

' -----------------
' Declare Variables
' -----------------

    Dim ws_Daily As Worksheet
        Set ws_Daily = ThisWorkbook.Sheets("Daily")

    Dim strCategory As String
        strCategory = Me.lst_Daily_Category.Value
    
    Dim strText As String
        strText = Me.txt_Text_to_Append.Value
            If strCategory = "Improved" Then
                strText = "IMP: " & strText
                strCategory = "Improved / Learned"
            ElseIf strCategory = "Learned" Then
                strText = "TIL: " & strText
                strCategory = "Improved / Learned"
            End If

    Dim intBlankRows As Long
        intBlankRows = 8 '4/7/2022
    
    Dim int_Category_Location As Long
        int_Category_Location = ws_Daily.Range("B:B").Find(WHAT:=strCategory, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Row
        
    Dim rngCategory As Range
        Set rngCategory = ws_Daily.Range("B" & int_Category_Location & ":" & "B" & (int_Category_Location + intBlankRows))

    Dim curRow As Long
        curRow = rngCategory.Find("").Row
        'Once you find the cateogory go down to the next blank to add, if more then 7 down bad news bears
    
    If curRow - int_Category_Location > intBlankRows Then MsgBox "WARNING: YOU ARE OVERWRITING TEXT"
    
' -------------------------
' Append to the Daily sheet
' -------------------------
    
    ws_Daily.Range("B" & curRow).Value = strText
    
End Sub
Public Function fx_OutputFile(strFileName)

' Purpose: This function outputs the link based on the selected file.
' Trigger: Called: uf_RefX_and_wsDaily_Update
' Updated: 5/12/2023

' Change Log:
'       5/16/2022:  Added the code for Agendas back in
'                   Added the strTextFileLoc to select the .txt if the .properties doesn't exist
'       5/12/2023:  Updated to reflect the switch from Agendas => Projects

' ***********************************************************************************************************************************

' -----------------
' Declare Variables
' -----------------

Dim strTextFileLoc As String

' --------------------------------
' Create the location for the file
' --------------------------------
    
    'Professional
    
    If Me.opt_Performance.Value = True Then
        strTextFileLoc = "C:\U Drive\Reference\RefX\Performance\RefX - Performance - " & strFileName & ".txt"
        
    ElseIf Me.opt_Individual.Value = True Then
        strTextFileLoc = "C:\U Drive\Reference\RefX\Individuals\RefX - " & strFileName & ".txt"
        
    ElseIf Me.opt_RefX.Value = True Then
        strTextFileLoc = "C:\U Drive\Reference\RefX\" & strFileName & ".txt"
        
    ElseIf Me.opt_Updates.Value = True Then
        strTextFileLoc = "C:\U Drive\Reference\RefX\Updates\" & strFileName & ".txt"
                
    ElseIf Me.opt_Projects.Value = True Then
        strTextFileLoc = "C:\U Drive\Reference\RefX\Projects\" & strFileName & ".txt"
        
    'Personal
    
    ElseIf Me.opt_Personal_Individual.Value = True Then
        strTextFileLoc = "D:\D Documents\Reference\RefX\Individuals\" & strFileName & ".txt"
    
    ElseIf Me.opt_Personal_RefX.Value = True Then
        strTextFileLoc = "D:\D Documents\Reference\RefX\" & strFileName & ".txt"
    
    ElseIf Me.opt_Personal_Updates.Value = True Then
        strTextFileLoc = "D:\D Documents\Reference\RefX\Updates\" & strFileName & ".txt"

    End If
  
' ---------------------------
' Output the link to the file
' ---------------------------
  
    If fx_File_Exists(strTextFileLoc) = False Then
        strTextFileLoc = Replace(strTextFileLoc, ".txt", ".properties")
    End If
  
  'properties
  
    fx_OutputFile = strTextFileLoc

End Function
Public Function fx_GetTextFileArray() As Variant
    
' Purpose: An array for the text file names.
' Trigger: Called: uf_RefX_and_wsDaily_Update
' Updated: 5/20/2023

' Change Log:
'       11/13/2020: Changed int_Category_Location to skip the day listed under Daily Goals
'       11/28/2021: Removed the Agenda related code
'       2/8/2022:   Added in Megan Travers
'       5/16/2022:  Udpated the individual code to pass the values from the folder, instead of an explicit list
'       5/20/2023:  Removed reference to the "[For Home]" RefX file

' ***********************************************************************************************************************************
    
' -----------------
' Declare Variables
' -----------------
    
    Dim TxtFile_Array() As String
    
    Dim dictTextFiles As New Dictionary
    
    Dim val As Variant
    
    Dim i As Long
    
' ----------------
' Create the array
' ----------------

    'Professional Arrays
    
    ' Professional - Performance
    If Me.opt_Performance.Value = True Then
        
        ReDim TxtFile_Array(1 To 3)
        
        TxtFile_Array(1) = "James"
        TxtFile_Array(2) = "Axcel"
        TxtFile_Array(3) = "Naomi"
    
    ' Professional - Individual
    ElseIf Me.opt_Individual.Value = True Then
        
        ' Create the dictionary, and then add just the RefX name to the array
        Set dictTextFiles = fx_List_Files_In_Folder("C:\U Drive\Reference\RefX\Individuals")
        
        ReDim TxtFile_Array(0 To dictTextFiles.count - 1)
        
        On Error Resume Next
        For i = 0 To UBound(TxtFile_Array)
            TxtFile_Array(i) = Mid(String:=dictTextFiles.Items(i), Start:=InStr(dictTextFiles.Items(i), "@ "), _
            Length:=InStrRev(dictTextFiles.Items(i), ".prop") - InStr(dictTextFiles.Items(i), "@ "))
        Next i
        On Error GoTo 0
        
    ' Professional - RefX
    ElseIf Me.opt_RefX.Value = True Then
        
        ReDim TxtFile_Array(1 To 4)
        
        TxtFile_Array(1) = "RefX - Coffee Roasters"
        TxtFile_Array(2) = "RefX - Coffee Shops"
        TxtFile_Array(3) = "RefX - Resteraunts"

        TxtFile_Array(4) = "RefX - Efficiency Training"
        
    ' Professional - Updates
    ElseIf Me.opt_Updates.Value = True Then
        
        ' Create the dictionary, and then add just the RefX name to the array
        Set dictTextFiles = fx_List_Files_In_Folder("C:\U Drive\Reference\RefX\Updates")
        
        ReDim TxtFile_Array(0 To dictTextFiles.count - 1)
        
        For i = 0 To UBound(TxtFile_Array)
            TxtFile_Array(i) = Mid(String:=dictTextFiles.Items(i), Start:=InStrRev(dictTextFiles.Items(i), "\") + 1, _
            Length:=InStrRev(dictTextFiles.Items(i), ".txt") - InStrRev(dictTextFiles.Items(i), "\") - 1)
        Next i
        
    ' Professional - Projects
        
    ElseIf Me.opt_Projects.Value = True Then
        
        ' Create the dictionary, and then add just the RefX name to the array
        Set dictTextFiles = fx_List_Files_In_Folder("C:\U Drive\Reference\RefX\Projects")
        
        ReDim TxtFile_Array(0 To dictTextFiles.count - 1)
        
        For i = 0 To UBound(TxtFile_Array)
            TxtFile_Array(i) = Mid(String:=dictTextFiles.Items(i), Start:=InStr(dictTextFiles.Items(i), "RefX "), _
            Length:=InStrRev(dictTextFiles.Items(i), ".prop") - InStr(dictTextFiles.Items(i), "RefX "))
        Next i
        
    'Personal Arrays
    
    ElseIf Me.opt_Personal_Individual.Value = True Then
        
        ReDim TxtFile_Array(1 To 5)
        
        TxtFile_Array(1) = "@ NICOLE"
        TxtFile_Array(2) = "@ MOM"
        TxtFile_Array(3) = "@ DAD"
        TxtFile_Array(4) = "@ MIKE"
        TxtFile_Array(5) = "@ KELSEY"
    
    ElseIf Me.opt_Personal_RefX.Value = True Then
        
        ReDim TxtFile_Array(1 To 3)
        
        TxtFile_Array(1) = "RefX - [MOSQUITO TASKS]"
        TxtFile_Array(2) = "RefX - [@ WEBSTER]"
        TxtFile_Array(3) = "RefX - Quotes to Add"
    
    ElseIf Me.opt_Personal_Updates.Value = True Then
        
        ReDim TxtFile_Array(1 To 5)
        
        TxtFile_Array(1) = "RefX - Vision + Mission"
        TxtFile_Array(2) = "RefX - Productivity Methodology"
        TxtFile_Array(3) = "RefX - [SOP - Productivity]"
        TxtFile_Array(4) = "RefX - [SOP - Personal]"
        TxtFile_Array(5) = "RefX - [SOP - Professional]"
    
    End If
    
' ----------------
' Output the array
' ----------------
    
    fx_GetTextFileArray = TxtFile_Array

End Function
Public Function fx_GetSubSectionArray() As Variant
    
' Purpose: An array for the text file sections.
' Trigger: Called: uf_RefX_and_wsDaily_Update
' Updated: 7/11/2022

' Change Log:
'       9/16/2019:  Initial Creation
'       8/31/2020:  Updated multiple sections to reflect my switch to CR&A
'       5/16/2022:  Updated to remove the opt_Frameworks and move those Sub-Sections under Update
'       7/11/2022:  Updated to reflect the new categories for my goals

' ***********************************************************************************************************************************
    
' -----------
' Declare your variables
' -----------
    
    Dim SubSection_Array() As Variant
    
' -----------
' Create the array
' -----------
    
    'Performance Arrays
    
    If Me.opt_Performance.Value = True Then
                   
        ReDim SubSection_Array(1 To 4)
        
        SubSection_Array(1) = "STRATEGY & INFRASTRUCTURE "
        SubSection_Array(2) = "PROCESS AUTOMATION"
        SubSection_Array(3) = "SPECIAL PROJECTS"
        SubSection_Array(4) = "PERSONAL DEVELOPMENT"
        
    End If
        
    'Updates Arrays
    If Me.opt_Updates.Value = True Then
        
        If Me.lst_TxtFile.Value = "RefX - [ISSUE LOG]" Then
            
            ReDim SubSection_Array(1 To 4)
            
            SubSection_Array(1) = "PERSONAL"
            SubSection_Array(2) = "GTD METHODOLOGY"
            SubSection_Array(3) = "PROCESS / TOOLS"
            SubSection_Array(4) = "OTHER"

        ElseIf Me.lst_TxtFile.Value = "RefX - [BIBLES]" Then
            
            ReDim SubSection_Array(1 To 4)
            
            SubSection_Array(1) = "EXCEL"
            SubSection_Array(2) = "VBA"
            SubSection_Array(3) = "TABLEAU"
            SubSection_Array(4) = "OTHER"

        ElseIf Left(Me.lst_TxtFile.Value, 4) = "CR&A" Then
            ReDim SubSection_Array(1 To 3)
    
            SubSection_Array(1) = "PEOPLE"
            SubSection_Array(2) = "PROCESS"
            SubSection_Array(3) = "TECHNOLOGY"

        End If
        
    End If

' -----------
' Output the array
' -----------

    If Len(Join(SubSection_Array)) = 0 Then
        ReDim SubSection_Array(1 To 1)
        SubSection_Array(1) = ""
    End If
        
    fx_GetSubSectionArray = SubSection_Array

End Function
Public Function fx_Getws_DailyCategoryArray() As Variant
    
' Purpose: An array for the text file names.
' Trigger: Called: uf_RefX_and_wsDaily_Update
' Updated: 9/16/2019

' ***********************************************************************************************************************************
    
' -----------------
' Declare Variables
' -----------------
    
    Dim Category_Array() As String
    
' ----------------
' Create the array
' ----------------
            
        ReDim Category_Array(1 To 5)
    
        Category_Array(1) = "Improved"
        Category_Array(2) = "Learned"
        Category_Array(3) = "Start / Continue"
        Category_Array(4) = "Stop / Change"
        Category_Array(5) = "Positive Experiences"

' ----------------
' Output the array
' ----------------
    
    fx_Getws_DailyCategoryArray = Category_Array

End Function
