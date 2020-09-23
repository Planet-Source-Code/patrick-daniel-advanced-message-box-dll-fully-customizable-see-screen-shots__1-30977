VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   975
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3240
   LinkTopic       =   "Form1"
   ScaleHeight     =   975
   ScaleWidth      =   3240
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "CLICK TO DEMONSTRATE"
      Height          =   855
      Left            =   75
      TabIndex        =   0
      Top             =   45
      Width           =   3030
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-- You Won't find many comments in here because I just decided to post it for the hell of it
'-- If you can't understand it...well, here's a good time to learn VB

Private Sub Command1_Click()
   
   Dim strMessage As String
   
   strMessage = "<B><FONT=Tahoma>This is an example of my Fully Customizable, Multi-Line Message Box!</FONT></B>" & _
            vbCrLf & vbCrLf & "<B><FONT=Tahoma><COLOR=" & vbRed & ">" & _
            "Using my FormatRTF Function, It provides the following Features</B></FONT></COLOR>" & _
            vbCrLf & vbCrLf & "<ALIGN=Center><SIZE=14>Sizes</SIZE>, <FONT=Tahoma>Font Names</FONT>, " & _
                     "<B>Bold</B>, <U>Underline</U>, <U>Italics</U>, <STRIKE>Strikethru</STRIKE>, " & _
                     "<COLOR=" & vbBlue & ">Colors</COLOR>" & vbCrLf & _
                     "<BULLET>Bullets</BULLET> And Alignment</ALIGN>" & vbCrLf & vbCrLf & _
            "<ALIGN=Left><B>It also supports Help Buttons, Default Buttons, Sounds</B></ALIGN>" & vbCrLf & _
            "<B>And is Backwards compatible with the standard MessageBox!" & vbCrLf & vbCrLf & _
            "<COLOR=" & vbBlue & ">Just Referece the DLL and issue a Find/Replace " & _
            "Msgbox With MsgBoxEx, and there you go!</COLOR></B>" & vbCrLf & vbCrLf
               
   '-- If you want the messagebox to format the code..such as above..you must set
   '-- Autosize to False and Supply Width and Height parameters.
   MsgBoxEx strMessage, vbYesNoCancel + vbMsgBoxHelpButton + vbDefaultButton1 + _
            vbInformation, "Bottomless, Customizable Message Box", False, 6000, 3100
   
   strMessage = "This is just like the plain old messagebox" & vbCrLf & _
               "It autogrows in width and height and sounds a beep and has all the features of a standard Msgbox" & _
               vbCrLf & vbCrLf & "-----------------------------------------------------------" & vbCrLf & vbCrLf & _
            "----THIS DISPLAYS IT:" & vbCrLf & _
            "MsgBoxEx strMessage, vbOKOnly + vbInformation, No Formatting, Regular Message Box"
   
   '-- If you want a standard..autosizing..normal messagebox
   '-- Leave autosize defaulted to True and DO NOT pass width and height parameters
   MsgBoxEx strMessage, vbOKOnly + vbInformation, "No Formatting, Regular Message Box"
   
End Sub
