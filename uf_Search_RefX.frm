VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} uf_Search_RefX 
   Caption         =   "  --- Dynamic RefX Search ---"
   ClientHeight    =   7712
   ClientLeft      =   36
   ClientTop       =   312
   ClientWidth     =   10176
   OleObjectBlob   =   "uf_Search_RefX.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "uf_Search_RefX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Declare Dictionaries
Dim dict_RefX_Files As Scripting.Dictionary

' Declare Strings
Dim str_PathParent As String
Dim str_FileOpeningType As String

Option Explicit
Private Sub UserForm_Initialize()
    
' Purpose: To select a RefX text file from the list and open it.
' Trigger: Selected from the uf_Project_Selector UserForm
' Updated: 5/9/2023

' Change Log:
'       5/9/2023:   Updated to remove part of the file path to shorten the output.

' ***********************************************************************************************************************************
    
    Me.StartUpPosition = 2
    
    Me.cmb_DynamicSearch.SetFocus
        
    With Me.lst_RefX 'Used to keep the simple name and full path
        .ColumnCount = 2
        .ColumnWidths = Me.lst_RefX.Width - 5 & ";" & "1"
    End With

Call Me.o_1_Create_File_List

End Sub
Private Sub cmb_DynamicSearch_Change()

' Purpose: To output the files for the given value in the combo box.
' Trigger: Called: uf_Search_RefX
' Updated: 9/30/2023

' Change Log:
'       9/30/2020:  Refreshed to use a Dictionary instead of an Array.
'       5/9/2023:   Updated to remove part of the file path to shorten the output.
'       9/23/2023:  Removed variables no longer used
'       9/30/2023:  Updated to limit the number of items in the ListBox to increase speed

' ***********************************************************************************************************************************

' -----------------
' Declare Variables
' -----------------

    Dim val As Variant
    
    Dim intCount As Long
                      
' ---------------------------------------------------
' Copy the values from the dictionary to the list box
' ---------------------------------------------------
    
    Me.lst_RefX.Clear
    
    For Each val In dict_RefX_Files
        If InStr(1, val, cmb_DynamicSearch.Value, vbTextCompare) Then 'If the name is similar then add to the list
            Me.lst_RefX.AddItem Mid(val, 8)
            intCount = intCount + 1 'Increase Counter
        End If
    
        If intCount = 25 Then Exit For 'Limit the number to improve speed
        
    Next val
    
End Sub
Private Sub lst_RefX_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

' Purpose: To open the applicable RefX file when double clicking.
' Trigger: Called: uf_Search_RefX
' Updated: 9/25/2023
' Reviewd: 9/25/2023

' Change Log:
'       2/14/2021:  Updated to open with Notepad++
'       5/9/2023:   Updated to add in the missing part of the file path, as a result of shortening the path in the RefX list
'       9/25/2023:  Moved the opening code to a new sub Me.o_2_Open_Selected_File

' ***********************************************************************************************************************************

    str_FileOpeningType = "Double Click"
    
    Call Me.o_2_Open_Selected_File

End Sub
Private Sub lst_RefX_Enter()

' Purpose: To open the applicable RefX file when entering the list, and only one value is present.
' Trigger: Called: uf_Search_RefX
' Updated: 9/25/2023
' Reviewd: 9/25/2023

' Change Log:
'       8/31/2023:  Original Creation, based on lst_RefX_KeyDown
'       9/25/2023:  Moved the opening code to a new sub Me.o_2_Open_Selected_File

' ***********************************************************************************************************************************

If Me.lst_RefX.ListCount = 1 Then
    
    str_FileOpeningType = "On Enter List"

    Call Me.o_2_Open_Selected_File
    
End If

End Sub
Private Sub lst_RefX_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

' Purpose: To open the applicable RefX file when hitting enter on the applicable value.
' Trigger: Called: uf_Search_RefX
' Updated: 9/25/2023
' Reviewd: 9/25/2023

' Change Log:
'       2/14/2021:  Initial Creation, splited out Enter vs Double Click, Enter should be quicker
'       5/9/2023:   Updated to add in the missing part of the file path, as a result of shortening the path in the RefX list
'       9/25/2023:  Moved the opening code to a new sub Me.o_2_Open_Selected_File

' ***********************************************************************************************************************************

If KeyCode = vbKeyReturn And Me.lst_RefX.Value <> "" Then
    
    str_FileOpeningType = "Press Enter"

    Call Me.o_2_Open_Selected_File
    
End If

End Sub
Private Sub cmd_Cancel_Click()
    
    Unload Me
    
End Sub
Sub o_1_Create_File_List()

' Purpose: To create the initial array that will be used to ID the files.
' Updated: 5/14/2023

' Change Log:
'       9/30/2020:  Initial Creation
'       10/26/2020: Added in the split for personal vs work computers
'       5/14/2023:  Updated to exclude the '(ARCHIVE)' folders
'                   Added ' And Left(val, 4) = "RefX" ' to ignore Templates

' ***********************************************************************************************************************************

' ----------------------
' Declare your variables
' ----------------------

    ' Declare Strings

    'Dim str_PathParent As String
        #If Personal <> 1 Then
            str_PathParent = "C:\U Drive\"
        #Else
            str_PathParent = "D:\D Documents\"
        #End If
        
    ' Declare Objects
    
    Dim objFSO As Object
        Set objFSO = CreateObject("Scripting.FileSystemObject")
    
    Dim objDirParent As Object
    Set objDirParent = objFSO.GetFolder(str_PathParent & "Reference\RefX\")
    
    Dim objSubFolder As Object
    
    Dim objFile As Object

    ' Declare Dictionaries

    Set dict_RefX_Files = New Scripting.Dictionary
    
' ---------------------------------------
' Load the RefX Files into the dictionary
' ---------------------------------------

On Error Resume Next

    'Run it for the main folder
        For Each objFile In objDirParent.Files
            If Left(objFile.Name, 4) = "RefX" Then ' Exclude templates
                dict_RefX_Files.Add key:=objFile.Name, Item:=objFile.path
            End If
        Next objFile
                        
    'Recur through each subfolder, excluding the archive folders
    For Each objSubFolder In objDirParent.SubFolders
        If objSubFolder.Name <> "(ARCHIVE)" Then
            
            For Each objFile In objSubFolder.Files
                If Left(objFile.Name, 4) = "RefX" Then ' Exclude templates
                    dict_RefX_Files.Add key:=objFile.Name, Item:=objFile.path
                End If
            Next objFile
            
        End If
    Next objSubFolder '
                       
On Error GoTo 0
                       
End Sub
Sub o_2_Open_Selected_File()

' Purpose: To open the selected file.
' Updated: 9/25/2023
' Reviewd: 9/25/2023

' Change Log:
'       9/25/2023:  Initial Creation

' ***********************************************************************************************************************************

On Error Resume Next

' -----------------
' Declare Variables
' -----------------

    Dim str_RefXFileName As String
    
    If str_FileOpeningType = "Double Click" Then
        str_RefXFileName = Me.lst_RefX.Value
    ElseIf str_FileOpeningType = "Press Enter" Then
        str_RefXFileName = Me.lst_RefX.Value
    ElseIf str_FileOpeningType = "On Enter List" Then
        str_RefXFileName = Me.lst_RefX.List(0)
    End If
                        
    Dim str_FullPath As String
        str_FullPath = dict_RefX_Files("RefX - " & str_RefXFileName) 'Pull the full path from the dictionary
            
' ----------------------
' Open the selected file
' ----------------------
    
    Call Shell("explorer.exe" & " " & str_FullPath, vbNormalFocus)
    
    Unload Me
    
End Sub
