Attribute VB_Name = "Reminders"
Option Explicit
Sub r_Mini_Review()

' -----------
' Set my UserForm properties
' -----------

    uf_BLANK.Show vbModeless
        uf_BLANK.Width = uf_BLANK.Width + 300
            uf_BLANK.lbl_BLANK.Width = uf_BLANK.lbl_BLANK.Width + 300
        uf_BLANK.Height = uf_BLANK.Height + 70
    
    uf_BLANK.Caption = "Mini Review"

' -----------
' Run your code
' -----------

uf_BLANK.lbl_BLANK.Caption = "   Purpose: To outline my approach for completing a mini review when I am feeling unfocused." & Chr(10) _
    & "      The purpose of the mini review is just to spend 5-10 minutes on a meta-moment to refocus.  " & _
    "If you feel unfocused then you should reassess / review at the next level up (ie what's the project, whats the purpose, etc.)" & Chr(10) & Chr(10) _
    & "      A few questions to help with this process:" & Chr(10) _
    & "      1) What is it that's really hanging me up right now?" & Chr(10) _
    & "      2) Am I really working toward a goal, or have I just become stuck on distracting pseudo-work?" & Chr(10) _
    & "      3) Is it really the interruptions that are bugging me, or has my trusted system just gone temporarily farkatke?" & Chr(10) _
    & "      4) Is everything here where it belongs just now?" & Chr(10) _
    & "      5) Is there something bugging me that I can just articulate as a problem and shunt into the right shelf in my system?" & Chr(10) _
    & "      6) Are any of these next actions completed, expired, or obviated?" & Chr(10) _
    & "      7) Has my inbox secretly turned into a safe harbor for stuff I just don't want to think about?" & Chr(10) _
    & "      8) What am I really committed to right now, and what's it going to take to move closer to completion today?" & Chr(10) _
    & "      9) Say to yourself: ""Have the kind of day that would make you proud""" & Chr(10)

End Sub

