VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} uf_Search_File 
   Caption         =   "  --- Dynamic File Search ---"
   ClientHeight    =   6704
   ClientLeft      =   60
   ClientTop       =   408
   ClientWidth     =   17688
   OleObjectBlob   =   "uf_Search_File.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "uf_Search_File"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' Declare Dictionaries / Arrays
Dim dict_FilePath As Scripting.Dictionary

Dim arry_FileType(1 To 4) As String

' Declare Strings
Dim str_FileName As String
Dim str_FileExtension As String
Dim str_PathParent As String
Dim str_FileOpeningType As String

' Declare Other Variables
Dim bol_SimpleSearch As Boolean

Option Explicit
Private Sub UserForm_Initialize()
    
' Purpose: To select a RefX text file from the list and open it.
' Trigger: Selected from the uf_Project_Selector UserForm
' Updated: 9/23/2023
' Reviewd: 9/14/2023

' Change Log:
'       9/14/2023:  Original Creation
'       9/23/2023:  Updated to allow for Simple OR Complex searches

' ***********************************************************************************************************************************
    
    ' Setup the UserForm for a Simple Search
    Me.Height = Me.Height - 18
    Me.lst_Files.Top = 30
    
    Me.txt_1_Modified.Visible = False
    Me.txt_2_Extension.Visible = False
    Me.txt_3_FillName.Visible = False
    Me.txt_4_FullPath.Visible = False
    
    ' Initialize the initial values
    Me.cmb_FileType.List = GetFileTypeArray 'Add the values for the File Type ComboBox
    bol_SimpleSearch = True
    
    ' Create the list of files
    Call Me.o_1_Create_File_List
    
    Me.cmb_DynamicSearch.SetFocus

End Sub
Private Sub cmd_ComplexSearch_Click()

' Purpose: To flip to a complex search and add the related deatil.
' Updated: 9/26/2023
' Reviewd: 9/23/2023

' Change Log:
'       9/23/2023:  Original Creation
'       9/26/2023:  Updated so that if I click Complex it uses the text entered in the Dynamic Search to recreate the list of files

' ***********************************************************************************************************************************

    ' Update to be a Complex Search
    bol_SimpleSearch = False
    
    ' Setup the UserForm for a Complex Search
    Me.Height = Me.Height + 18
    
    Me.cmd_ComplexSearch.Visible = False
    
    Me.txt_1_Modified.Visible = True
    Me.txt_2_Extension.Visible = True
    Me.txt_3_FillName.Visible = True
    Me.txt_4_FullPath.Visible = True

    Me.cmd_OpenFolder.Visible = True
    Me.cmd_OpenFolder.Left = 228

    ' Update to include 3 columns
    With Me.lst_Files
        .ColumnCount = 4
        .ColumnWidths = "60,40,410,360"
        .Top = 48
        .Clear
    End With
    
    ' Recreate the results using what is in the Dynamic Search combobox
    If Me.cmb_DynamicSearch <> "" Then Call Me.cmb_DynamicSearch_Change
    
    ' Set the focus so I can start typing
    Me.cmb_DynamicSearch.SetFocus
    
End Sub
Private Sub cmb_FileType_Change()

' Purpose: To determine the applicable extension to filter on, based on the selected File Type.
' Trigger: Called: uf_Search_File
' Updated: 9/26/2023
' Reviewd: 9/14/2023

' Change Log:
'       9/14/2023:  Initial Creation
'       9/26/2023:  Updated so that if I change the File Type it uses the text entered in the Dynamic Search to recreate the list of files

' ***********************************************************************************************************************************

    Select Case cmb_FileType.Value
    
        Case "Excel"
            str_FileExtension = ".xls"
    
        Case "Word"
            str_FileExtension = ".doc"
    
        Case "Powerpoint"
            str_FileExtension = ".ppt"
    
        Case "PDF"
            str_FileExtension = ".pdf"
    
    End Select

    ' Recreate the results using what is in the Dynamic Search combobox
    If Me.cmb_DynamicSearch <> "" Then Call Me.cmb_DynamicSearch_Change

End Sub
Public Sub cmb_DynamicSearch_Change()

' Purpose: To output the files for the given value in the combo box.
' Trigger: Called: uf_Search_File
' Updated: 9/29/2023
' Reviewd: 9/14/2023

' Change Log:
'       9/14/2023:  Initial Creation, based on uf_Search_RefX code
'       9/21/2023:  Updated to be more dyanmic with the FileExtension determination
'       9/22/2023:  Updated to include a 2nd column for file path
'       9/23/2023:  Removed variables no longer used
'       9/26/2023:  Updated to limit the number of items in the ListBox to increase speed
'       9/29/2023:  Updated to determine the modified date when populating the Dynamic Search instead of the dictionary

' ***********************************************************************************************************************************

' -----------------
' Declare Variables
' -----------------

    Dim var_TargetFile As Variant
    
    Dim intCount As Long
                      
' ---------------------------------------------------
' Copy the values from the dictionary to the list box
' ---------------------------------------------------
    
On Error Resume Next
    
    Me.lst_Files.Clear
    
    For Each var_TargetFile In dict_FilePath
        
        If str_FileExtension = "" Or Left(Mid(var_TargetFile, InStrRev(var_TargetFile, ".")), 4) = str_FileExtension Then ' If the file extension matches include
            If InStr(1, var_TargetFile, cmb_DynamicSearch.Value, vbTextCompare) Then 'If the name is similar then add to the list
                
                If bol_SimpleSearch = True Then
                    Me.lst_Files.AddItem var_TargetFile
                    intCount = intCount + 1 'Increase Counter
                ElseIf bol_SimpleSearch = False Then
                    Me.lst_Files.AddItem
                    Me.lst_Files.List(intCount, 0) = Format(FileDateTime(dict_FilePath(var_TargetFile)), "mm/dd/yyyy")
                    Me.lst_Files.List(intCount, 1) = Left(Mid(var_TargetFile, InStrRev(var_TargetFile, ".")), 5)
                    Me.lst_Files.List(intCount, 2) = var_TargetFile
                    Me.lst_Files.List(intCount, 3) = Mid(Left(dict_FilePath(var_TargetFile), InStrRev(dict_FilePath(var_TargetFile), "\")), Len(str_PathParent))
                    
                    intCount = intCount + 1 'Increase Counter
                End If ' bol_SimpleSearch = True
                
            End If ' InStr
        End If ' str_FileExtension
        
        If intCount = 20 Then Exit For ' Limit the number to improve speed
        
    Next var_TargetFile
    
On Error GoTo 0
    
End Sub
Private Sub lst_Files_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

' Purpose: To open the applicable selected file when double clicking.
' Trigger: Called: uf_Search_File
' Updated: 9/21/2023
' Reviewd: 9/14/2023

' Change Log:
'       9/14/2023:  Initial Creation, based on uf_Search_RefX code
'       9/21/2023:  Updated to be more explicit with how to handle different file types
'                   Updated to activate the applicable Excel if it was opened
'                   Moved the opening code to a new sub Me.o_2_Open_Selected_File

' ***********************************************************************************************************************************

    str_FileOpeningType = "Double Click"

    Call Me.o_2_Open_Selected_File

End Sub
Private Sub lst_Files_Enter()

' Purpose: To open the applicable selected file when entering the list, and only one value is present.
' Trigger: Called: uf_Search_File
' Updated: 9/21/2023
' Reviewd: 9/14/2023

' Change Log:
'       9/14/2023:  Initial Creation, based on uf_Search_RefX code
'       9/21/2023:  Moved the opening code to a new sub Me.o_2_Open_Selected_File

' ***********************************************************************************************************************************

If Me.lst_Files.ListCount = 1 Then

    str_FileOpeningType = "On Enter List"

    Call Me.o_2_Open_Selected_File
    
End If

End Sub
Private Sub lst_Files_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

' Purpose: To open the applicable selected file when hitting enter on the applicable value.
' Trigger: Called: uf_Search_File
' Updated: 9/21/2023
' Reviewd: 9/14/2023

' Change Log:
'       9/14/2023:  Initial Creation, based on uf_Search_RefX code
'       9/21/2023:  Moved the opening code to a new sub Me.o_2_Open_Selected_File

' ***********************************************************************************************************************************

If KeyCode = vbKeyReturn And Me.lst_Files.Value <> "" Then

    str_FileOpeningType = "Press Enter"

    Call Me.o_2_Open_Selected_File
    
End If

End Sub
Private Sub cmd_OpenFolder_Click()

' Purpose: To open the parent folder for the selected file.
' Updated: 10/6/2023
' Reviewd: 10/6/2023

' Change Log:
'       10/6/2023:  Initial Creation

' ***********************************************************************************************************************************

    If Me.lst_Files.Value <> "" Then
        Call Me.o_3_Open_Selected_File_Parent_Folder
    Else
        MsgBox "Please select a file first."
    End If

End Sub
Private Sub cmd_Cancel_Click()
    
    Unload Me
    
End Sub
Sub o_1_Create_File_List()

' Purpose: To create the Dictonary that will be used to house the file name and full path.
' Trigger: UserForm Initialize
' Updated: 11/4/2023
' Reviewd: 9/14/2023

' Change Log:
'       9/14/2023:  Initial Creation
'                   Did some Inception level sh*t and recurred another level down
'       9/29/2023:  Updated to determine the modified date when populating the Dynamic Search instead of the dictionary
'                   Round 2, recurring down to a 3rd level of folders
'       11/4/2023:  Moved the code to Macros and create the Dictionary on Workbook.Opemn event

' ***********************************************************************************************************************************

' -----------------
' Declare Variables
' -----------------

    ' Declare Worksheets
    Dim wsLists As Worksheet
    Set wsLists = ThisWorkbook.Sheets("Lists")

    ' Declare Ranges
    
    Dim intLastRow_wsLists As Long
        intLastRow_wsLists = wsLists.Cells(Rows.count, "F").End(xlUp).Row
    
    Dim rngData As Range
    Set rngData = wsLists.Range("F2:G" & intLastRow_wsLists)
    
    Dim i As Long
    
    ' Declare Strings

    'Dim str_PathParent As String
        #If Personal <> 1 Then
            str_PathParent = "C:\U Drive\"
        #Else
            str_PathParent = "D:\D Documents\"
        #End If
    
    ' Declare Array
    
    Dim arry_Files() As Variant
    'Dim arry_Files(1 To intLastRow_wsLists, 1 To 2) As Variant
    
    ' Declare Dictionaries

    Set dict_FilePath = New Scripting.Dictionary
    
' -----------------------------------------------
' Load the RefX Files into the Array from wsLists
' -----------------------------------------------

    arry_Files = Application.Transpose(rngData)

' ------------------------------------------------------
' Load the RefX Files into the Dictionary from the Array
' ------------------------------------------------------

On Error Resume Next

    For i = 1 To intLastRow_wsLists
        dict_FilePath.Add key:=arry_Files(1, i), Item:=arry_Files(2, i)
    Next i
        
On Error GoTo 0
                       
End Sub
Sub o_2_Open_Selected_File()

' Purpose: To open the selected file.
' Trigger: Double click a value / Hit Enter on a value / Enter the list with only 1 value
' Updated: 10/9/2023
' Reviewd: 9/21/2023

' Change Log:
'       9/21/2023:  Initial Creation
'       9/25/2023:  Added code to activate the Excel workbook if I opened an Excel, so I can still use .FollowHyperlink (Faster)
'       9/27/2023:  Updated to handle the first column not being the file name for a Complex Search
'       10/9/2023:  Added a wait to try to resolve the crashing when opening an Excel 'Run-time error '-2147417848(80010108)''

' ***********************************************************************************************************************************

On Error Resume Next

' -----------------
' Declare Variables
' -----------------

    If bol_SimpleSearch = True Then

        If str_FileOpeningType = "Double Click" Then
            str_FileName = Me.lst_Files.Value
        ElseIf str_FileOpeningType = "Press Enter" Then
            str_FileName = Me.lst_Files.Value
        ElseIf str_FileOpeningType = "On Enter List" Then
            str_FileName = Me.lst_Files.List(0)
        End If
    
    Else
    
        If str_FileOpeningType = "Double Click" Then
            str_FileName = Me.lst_Files.List(lst_Files.ListIndex, 2)
        ElseIf str_FileOpeningType = "Press Enter" Then
            str_FileName = Me.lst_Files.List(lst_Files.ListIndex, 2)
        ElseIf str_FileOpeningType = "On Enter List" Then
            str_FileName = Me.lst_Files.List(0, 2)
        End If
    
    End If
                        
    Dim str_FullPath As String
        str_FullPath = dict_FilePath(str_FileName) 'Pull the full path from the dictionary
            
    Dim str_SelectedFileExtension As String
        str_SelectedFileExtension = Left(Right(str_FileName, Len(str_FileName) - InStrRev(str_FileName, ".") + 1), 4)

' ----------------------
' Open the selected file
' ----------------------
    
    If str_SelectedFileExtension = ".xls" Or str_SelectedFileExtension = ".doc" Or str_SelectedFileExtension = ".ppt" Or str_SelectedFileExtension = ".pdf" Then
        ThisWorkbook.FollowHyperlink (str_FullPath)
    Else
        Call Shell("explorer.exe" & " " & str_FullPath, vbNormalFocus)
    End If
    
    Unload Me
    
    Application.Wait (Now + TimeValue("0:00:01")) ' Added a wait on 10/9/23 to try to resolve the crashing
    
    If str_SelectedFileExtension = ".xls" Then
        Application.SendKeys ("%{TAB}") ' Switch to the Excel
    End If
                       
End Sub
Sub o_3_Open_Selected_File_Parent_Folder()

' Purpose: To open the selected file's parent folder.
' Trigger: Click me.cmd_OpenFolder
' Updated: 10/6/2023
' Reviewd: 10/6/2023

' Change Log:
'       10/6/2023:  Initial Creation

' ***********************************************************************************************************************************

On Error Resume Next

' -----------------
' Declare Variables
' -----------------

   'Dim str_FileName as String
        str_FileName = Me.lst_Files.List(lst_Files.ListIndex, 2)
    
    Dim str_FullPath As String
        str_FullPath = dict_FilePath(str_FileName) 'Pull the full path from the dictionary
            
    Dim str_ParentFolderPath As String
        str_ParentFolderPath = Left(str_FullPath, Len(str_FullPath) - Len(str_FileName))
            
' --------------------------------------
' Open the selected file's parent folder
' --------------------------------------
    
    ThisWorkbook.FollowHyperlink (str_ParentFolderPath)
    
    Unload Me
                       
End Sub
Public Function GetFileTypeArray() As Variant
    
' Purpose: To output the File Type array for my File Selector
' Updated: 9/14/2023
' Reviewd: 9/14/2023

' Change Log:
'       9/14/2023:  Original Creation based on GetAreaArray

' ***********************************************************************************************************************************
    
    arry_FileType(1) = "Excel"
    arry_FileType(2) = "Word"
    arry_FileType(3) = "Powerpoint"
    arry_FileType(4) = "PDF"

    GetFileTypeArray = arry_FileType

End Function
