VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} uf_Color_Selector 
   Caption         =   "Color Selector"
   ClientHeight    =   1884
   ClientLeft      =   0
   ClientTop       =   180
   ClientWidth     =   9456.001
   OleObjectBlob   =   "uf_Color_Selector.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "uf_Color_Selector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' Declare Strings
Dim str_Fill_Type As String

' Declare "Colors"
Dim clrGrey1 As Long

Dim clrBlue1 As Long
Dim clrBlue2 As Long

Dim clrRed1 As Long
Dim clrRed2 As Long
    
Dim clrPurple1 As Long
Dim clrPurple2 As Long

Dim clrGreen1 As Long
Dim clrGreen2 As Long
Dim clrGreen3 As Long
        
Dim clrOrange1 As Long

Dim clrSelected As Long

Option Explicit
Private Sub UserForm_Initialize()

' ***********************************************************************************************************************************
'
' Purpose: To apply the cell fill / tab color with the click of a button.
'
' Trigger: Ribbon Icon - Personal Macros > TBD
'
' Updated: 2/8/2022

' Change Log:
'       6/1/2021:  Initial Creation
'       6/10/2021: Added code to set the default selection
'       6/10/2021: Added the default button colors and code to remove the caption for the selected button
'       6/18/2021: Split out selecting and applying the color to make the code more dynamic
'       2/8/2022:  Added the code to update the ControlTipText, based on selection
'                  Updated to limit the colors being cycled through, and merged the 1 and 2 colors
' ***********************************************************************************************************************************

' ----------------
' Assign Variables
' ----------------

    ' Assign "Colors"
    clrGrey1 = RGB(230, 230, 230)

    clrBlue1 = RGB(211, 223, 238)
    clrBlue2 = RGB(184, 204, 228)
    
    clrRed1 = RGB(239, 211, 210)
    clrRed2 = RGB(230, 184, 183)
        
    clrPurple1 = RGB(222, 215, 231)
    clrPurple2 = RGB(204, 192, 218)
    
    clrGreen1 = RGB(230, 238, 213)
    clrGreen2 = RGB(216, 228, 188)
    
    clrOrange1 = RGB(253, 233, 217)

' -------------------
' Apply Button Colors
' -------------------

    cmd_Grey.BackColor = clrGrey1
    cmd_Blue.BackColor = clrBlue2
    cmd_Red.BackColor = clrRed2
    cmd_Purple.BackColor = clrPurple2
    cmd_Green.BackColor = clrGreen2
    cmd_Orange.BackColor = clrOrange1

' ------------
' Set Defaults
' ------------

    If ActiveCell.Value2 <> "" Then
        Me.opt_CellFill.Value = True
    Else
        Me.opt_TabColor.Value = True
    End If

End Sub
Private Sub opt_CellFill_Click()

    ' Assign Strings
    str_Fill_Type = "Cell"
    
    Call Me.o_02_Add_ControlTipText

End Sub
Private Sub opt_TabColor_Click()

    str_Fill_Type = "Tab"
    ' Assign Strings
    
    Call Me.o_02_Add_ControlTipText
    
End Sub
Private Sub cmd_Grey_Click()

' ------------------------------------
' Save the color as the selected color
' ------------------------------------

    clrSelected = clrGrey1

' ---------------
' Apply the color
' ---------------

    Call o_1_Apply_Color

    cmd_Grey.BackColor = clrSelected
        cmd_Grey.Caption = ""

End Sub
Private Sub cmd_Blue_Click()

' ------------------------------------------
' Cycle through the color for the Cell / Tab
' ------------------------------------------

    ' Select Fill Cell Color
    If str_Fill_Type = "Cell" Then
        clrSelected = clrBlue1
    End If
    
    ' Select Tab Color
    If str_Fill_Type = "Tab" Then
        
        With ActiveWorkbook.ActiveSheet.Tab
            If .Color = clrBlue2 Then
                clrSelected = clrBlue1
            Else
                clrSelected = clrBlue2
            End If
        End With
        
    End If

' ---------------
' Apply the color
' ---------------

    Call o_1_Apply_Color

    cmd_Blue.BackColor = clrSelected
        cmd_Blue.Caption = ""

End Sub
Private Sub cmd_Red_Click()

' ------------------------------------------
' Cycle through the color for the Cell / Tab
' ------------------------------------------

    ' Apply Fill Cell Color
    If str_Fill_Type = "Cell" Then
        clrSelected = clrRed1
    End If
    
    ' Apply Tab Color
    If str_Fill_Type = "Tab" Then
        
        With ActiveWorkbook.ActiveSheet.Tab
            If .Color = clrRed2 Then
                clrSelected = clrRed1
            Else
                clrSelected = clrRed2
            End If
        End With
        
    End If

' ---------------
' Apply the color
' ---------------

    Call o_1_Apply_Color

    cmd_Red.BackColor = clrSelected
        cmd_Red.Caption = ""
    
End Sub
Private Sub cmd_Purple_Click()

' ------------------------------------------
' Cycle through the color for the Cell / Tab
' ------------------------------------------

    ' Apply Fill Cell Color
    If str_Fill_Type = "Cell" Then
        clrSelected = clrPurple1
    End If
    
    ' Apply Tab Color
    If str_Fill_Type = "Tab" Then
        
        With ActiveWorkbook.ActiveSheet.Tab
            If .Color = clrPurple2 Then
                clrSelected = clrPurple1
            Else
                clrSelected = clrPurple2
            End If
        End With
        
    End If

' ---------------
' Apply the color
' ---------------

    Call o_1_Apply_Color

    cmd_Purple.BackColor = clrSelected
        cmd_Purple.Caption = ""

End Sub
Private Sub cmd_Green_Click()

' ------------------------------------------
' Cycle through the color for the Cell / Tab
' ------------------------------------------

    ' Apply Fill Cell Color
    If str_Fill_Type = "Cell" Then
        clrSelected = clrGreen1
    End If

    ' Apply Tab Color
    If str_Fill_Type = "Tab" Then
        
        With ActiveWorkbook.ActiveSheet.Tab
            If .Color = clrGreen2 Then
                clrSelected = clrGreen1
            Else
                clrSelected = clrGreen2
            End If
        End With
        
    End If

' ---------------
' Apply the color
' ---------------

    Call o_1_Apply_Color

    cmd_Green.BackColor = clrSelected
        cmd_Green.Caption = ""

End Sub
Private Sub cmd_Orange_Click()

' -----------
' Save the color as the selected color
' -----------

    clrSelected = clrOrange1

' ---------------
' Apply the color
' ---------------

    Call o_1_Apply_Color

    cmd_Orange.BackColor = clrSelected
        cmd_Orange.Caption = ""

End Sub
Private Sub cmd_Cancel_Click()

    Unload Me

End Sub
Sub o_02_Add_ControlTipText()

' Purpose: To apply the selected color to the cell / tab.
' Trigger: Called by OptionButtons
' Updated: 12/6/2023

' Change Log:
'       6/18/2021:  Intial Creation
'       6/18/2021:  Switched to use SelectedSheets
'       12/6/2023:  Updated the Control Tips

' ***********************************************************************************************************************************

    If str_Fill_Type = "Cell" Then
        
        Me.cmd_Grey.ControlTipText = "Standard Header (Original or Static Data)"
        Me.cmd_Blue.ControlTipText = "Standard Header (Manually Reviewed/Updated Data)"
        Me.cmd_Red.ControlTipText = "Virtual Field (Calculated Data)"
        Me.cmd_Orange.ControlTipText = "Virtual Field (Lookup Data)"
        Me.cmd_Purple.ControlTipText = "Virtual Field (Created or Calculated by Code)"
        Me.cmd_Green.ControlTipText = "Virtual Field (Change Flag or Temp Field)"
        
    ElseIf str_Fill_Type = "Tab" Then
        
        Me.cmd_Grey.ControlTipText = "Not Defined"
        Me.cmd_Blue.ControlTipText = "Notes, Information, Reference"
        Me.cmd_Red.ControlTipText = "Main Sheet / Output / Support Worksheets"
        Me.cmd_Purple.ControlTipText = "Main Data Source / Original Raw Data"
        Me.cmd_Green.ControlTipText = "One off reporting, changes, helper sheets, temp pivot tables"
        Me.cmd_Orange.ControlTipText = "Not Defined"
        
    End If

End Sub
Sub o_1_Apply_Color()

' Purpose: To apply the selected color to the cell / tab.
' Trigger: Called by Color cmd button
' Updated: 6/18/2021

' Change Log:
'       6/18/2021: Intial Creation
'       6/18/2021: Switched to use SelectedSheets

' ***********************************************************************************************************************************

Dim ws As Variant

' -----------
' Apply the color to the Cell / Tab
' -----------

    If str_Fill_Type = "Cell" Then
        Selection.Interior.Color = clrSelected
    ElseIf str_Fill_Type = "Tab" Then
        For Each ws In ActiveWorkbook.Windows(1).SelectedSheets
            ws.Tab.Color = clrSelected
        Next ws
    End If

End Sub
