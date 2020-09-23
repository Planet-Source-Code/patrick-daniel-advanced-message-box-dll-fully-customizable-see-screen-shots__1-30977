VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Begin VB.Form frmMsgBox 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " "
   ClientHeight    =   1545
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5220
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1545
   ScaleWidth      =   5220
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdButton 
      Height          =   390
      Index           =   3
      Left            =   3960
      TabIndex        =   5
      Top             =   1065
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.CommandButton cmdButton 
      Height          =   390
      Index           =   2
      Left            =   2700
      TabIndex        =   4
      Top             =   1065
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.CommandButton cmdButton 
      Height          =   390
      Index           =   1
      Left            =   1410
      TabIndex        =   3
      Top             =   1065
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.CommandButton cmdButton 
      Height          =   390
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   1065
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   75
      Left            =   -90
      TabIndex        =   2
      Top             =   855
      Width           =   5340
   End
   Begin RichTextLib.RichTextBox txtMessage 
      Height          =   705
      Left            =   900
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   120
      Width           =   4260
      _ExtentX        =   7514
      _ExtentY        =   1244
      _Version        =   393217
      BackColor       =   -2147483633
      BorderStyle     =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      MousePointer    =   1
      Appearance      =   0
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"frmMsgBox.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblMessage 
      Caption         =   "Label1"
      Height          =   195
      Left            =   900
      TabIndex        =   6
      Top             =   120
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Image imgQuestion 
      Height          =   480
      Left            =   225
      Picture         =   "frmMsgBox.frx":007B
      Top             =   150
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgInformation 
      Height          =   480
      Left            =   225
      Picture         =   "frmMsgBox.frx":0385
      Top             =   150
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgExclaimation 
      Height          =   480
      Left            =   225
      Picture         =   "frmMsgBox.frx":07C7
      Top             =   120
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgCritical 
      Height          =   480
      Left            =   225
      Picture         =   "frmMsgBox.frx":0AD1
      Top             =   135
      Visible         =   0   'False
      Width           =   480
   End
End
Attribute VB_Name = "frmMsgBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-- You Won't find many comments in here because I just decided to post it for the hell of it
'-- If you can't understand it...well, here's a good time to learn VB

Private strCaption As String

Private Sub Form_Load()
   lblMessage.AutoSize = True
End Sub

Public Property Get ButtonCaption()
On Error Resume Next

   ButtonCaption = strCaption
   
End Property

Private Sub cmdButton_Click(Index As Integer)
On Error Resume Next

   strCaption = cmdButton(Index).Caption
   Me.Hide
   
End Sub

Private Sub Form_Unload(Cancel As Integer)

   Set frmMsgBox = Nothing
   gblnFormLoaded = False
   
End Sub
