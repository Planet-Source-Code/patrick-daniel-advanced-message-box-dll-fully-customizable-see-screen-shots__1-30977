VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAdvancedMsgBox"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'-- You Won't find many comments in here because I just decided to post it for the hell of it
'-- If you can't understand it...well, here's a good time to learn VB

Dim Button1    As CommandButton
Dim Button2    As CommandButton
Dim Button3    As CommandButton
Dim button4    As CommandButton
Dim lHelp      As Long
Dim bAuto      As Boolean

Private Function LineCount(lHwnd As Long) As Long
   LineCount = SendMessage(lHwnd, EM_GETLINECOUNT, 0&, 0&)
End Function

Private Function LineIndex(LineNum As Integer) As Long
  LineIndex = SendMessage(frmMsgBox.txtMessage.hWnd, EM_LINEINDEX, LineNum, 0)
End Function

Private Function LineLength(LineNum As Integer) As Long
   LineLength = SendMessage(frmMsgBox.txtMessage.hWnd, EM_LINELENGTH, LineIndex(LineNum), 0)
End Function

Public Function MsgBoxEx(Prompt, Optional Buttons As VbMsgBoxStyle = vbOKOnly + vbInformation, Optional Title, _
                     Optional AutoSize As Boolean = True, _
                     Optional iWidth As Integer, Optional iHeight As Integer) As VbMsgBoxResult
   
   MsgBoxEx = ShowMe(Prompt, Buttons, Title, AutoSize, iWidth, iHeight)
   
End Function

Public Function ShowMe(Prompt, Optional Buttons As VbMsgBoxStyle = vbOKOnly + vbInformation, Optional Title, _
                     Optional AutoSize As Boolean = True, _
                     Optional iWidth As Integer, Optional iHeight As Integer) As VbMsgBoxResult
                     
On Error Resume Next

   Dim lTemp As Long
   Dim lButtons As Long
   Dim lDefault As Long
   Dim lModal As VbMsgBoxStyle
   Dim strCaption As String
   Dim lBeep As Long
   Dim vKey As Variant
   Dim vCurrentMousePointer As Variant
   
   bAuto = AutoSize
   
   If gblnFormLoaded Then Exit Function
   
   '-- Explicitly reference the Command Buttons - Is faster and more efficient
   With frmMsgBox
      Set Button1 = .cmdButton(0)
      Set Button2 = .cmdButton(1)
      Set Button3 = .cmdButton(2)
      Set button4 = .cmdButton(3)
   End With
   
   '-- Check for Help Button
   If Buttons >= vbMsgBoxHelpButton Then
      lTemp = Buttons Mod vbMsgBoxHelpButton
      lHelp = vbMsgBoxHelpButton
      BuildHelp
   Else
      lTemp = Buttons
      lHelp = 0
   End If
   
   '-- We check for SystemModal...But us is never used.  I force Application Modal
   '-- As of WINNT and Win2000, System Modal is no longer an option but it still exists
   '-- in the msgboxtype constants so we have to check for it
   If lTemp >= vbSystemModal Then
      lTemp = lTemp Mod vbSystemModal
   End If
   
   '-- Determine if the user has changed the Default Button
   If lTemp >= vbDefaultButton2 Then
      lButtons = lTemp Mod vbDefaultButton2
      lDefault = lTemp - lButtons
      SetDefault lDefault
   Else
      lButtons = lTemp
      lDefault = 0
      SetDefault lDefault
   End If
   
   '-- We must AutoSize(if requested) first before we add the buttons...or the buttons won't center properly
   If AutoSize Then
      Call AutoSizeForm(Prompt)
   Else
      
      If iWidth = 0 Then
         iWidth = 4260
      End If
      
      If iHeight = 0 Then
         iHeight = 800
      End If
      
      With frmMsgBox
         .txtMessage = Prompt
         FormatText
         FormatSize iWidth, iHeight
      End With
   End If
   
   '********************************************************
   '-- Check for the Picture and buttons
   If lButtons >= 0 And lButtons <= 15 Then
      BuildButtons lButtons
      BuildPicture 65
      lBeep = MB_OK
      GoTo Bottom
   End If
   
   If lButtons >= 16 And lButtons <= 31 Then
      BuildPicture 16
      BuildButtons lButtons Mod 16
      lBeep = MB_ICONASTERISK
      GoTo Bottom
   End If
   
   If lButtons >= 32 And lButtons <= 47 Then
      BuildPicture 32
      BuildButtons lButtons Mod 32
      lBeep = MB_OK
      GoTo Bottom
   End If
   
   If lButtons >= 48 And lButtons <= 63 Then
      BuildPicture 48
      BuildButtons lButtons Mod 48
      lBeep = MB_ICONASTERISK
      GoTo Bottom
   End If
   
   If lButtons >= 64 And lButtons <= 255 Then
      BuildPicture 64
      BuildButtons lButtons Mod 64
      lBeep = MB_OK
      GoTo Bottom
   End If
   '********************************************************
   
Bottom:
   
   With frmMsgBox
      
      '-- Align the text to Image if fewer than 3 lines visible
      If AutoSize Then
         If Not .lblMessage.WordWrap Then
            vKey = Split(.lblMessage.Caption, vbCrLf)
            If UBound(vKey) + 1 = 1 Then
               .lblMessage.Top = .lblMessage.Top + 150
            ElseIf UBound(vKey) + 1 = 2 Then
               .lblMessage.Top = .lblMessage.Top + 75
            End If
         End If
      Else
         If LineCount(.txtMessage.hWnd) = 1 Then
            .txtMessage.Top = .txtMessage.Top + 150
            .txtMessage.Height = .txtMessage.Height - 150
         ElseIf LineCount(.txtMessage.hWnd) = 2 Then
            .txtMessage.Top = .txtMessage.Top + 75
            .txtMessage.Height = .txtMessage.Height - 75
         End If
      End If
      
      If Not IsMissing(Title) Then
         If Title = "" Then
            .Caption = " "
         Else
            .Caption = Title
         End If
      Else
         '-- Due to the Form style..something must be in the caption or header
         '-- will vanish
         .Caption = " "
      End If
            
      MessageBeep lBeep
      
      gblnFormLoaded = True
      
      DoEvents
      
      .Show vbModal
      
   End With
   
   strCaption = frmMsgBox.ButtonCaption
   
   ShowMe = GetAnswer(strCaption)
   
   Unload frmMsgBox
   
End Function

Private Sub FormatSize(iWidth As Integer, iHeight As Integer)
   Dim iTemp As Integer
   Dim i As Integer
   
   With frmMsgBox
   
      .txtMessage.Width = iWidth
      .txtMessage.Height = iHeight
         
      iTemp = .txtMessage.Top + .txtMessage.Height
         
      .Width = .txtMessage.Width + 1100
      .Frame1.Width = .Width + 100
      .Frame1.Top = iTemp + 25
         
      For i = 0 To 3
         .cmdButton(i).Top = iTemp + 225
      Next
         
      .Height = iTemp + 1100
      
   End With
   
End Sub

Private Sub AutoSizeForm(ByVal strPrompt As String)
   
   Dim iTemp As Integer '-- Holds the Default Height of form
   Dim i As Integer
   
   With frmMsgBox
      
      .lblMessage.Caption = strPrompt
      
      If .ScaleX(.lblMessage.Width + 1200, vbTwips, vbPixels) >= 640 Then
         .lblMessage.Width = 8500
         .lblMessage.WordWrap = True
         .Width = .ScaleX(640, vbPixels, vbTwips)
      Else
         .Width = .lblMessage.Width + 1200
      End If
      
      .Frame1.Width = .Width + 100
      
      If .lblMessage.Height > .txtMessage.Height Then
         
         iTemp = .lblMessage.Height - .txtMessage.Height
         
         .Height = .Height + iTemp
         
         For i = 0 To 3
            .cmdButton(i).Top = .cmdButton(i).Top + iTemp
         Next
         
         .Frame1.Top = .Frame1.Top + iTemp
      
      End If
      
      .lblMessage.Visible = True
      .txtMessage.Visible = False
      
   End With
   
End Sub

Private Function GetAnswer(strCaption As String) As VbMsgBoxResult
On Error Resume Next

   Select Case strCaption
   
      Case "OK"
         GetAnswer = vbOK
      Case "Cancel"
         GetAnswer = vbCancel
      Case "Yes"
         GetAnswer = vbYes
      Case "No"
         GetAnswer = vbNo
      Case "Retry"
         GetAnswer = vbRetry
      Case "Ignore"
         GetAnswer = vbIgnore
      Case "Abort"
         GetAnswer = vbAbort
      Case "Help"
         GetAnswer = 9
         
   End Select
      
End Function

Private Sub FormatText()
   
   Dim i As Integer
   
   Dim strTags() As Variant
   
   Dim iLength As Integer
   Dim strValue As String
   Dim iStart As Integer
   Dim iEnd As Integer
   Dim strStartTag As String
   Dim strEndTag As String
   Dim iStartTag As Integer
   Dim iEndTag As Integer
   Dim iLenST As Integer
   Dim iCount As Integer
   
   With frmMsgBox.txtMessage
            
      '-- Look For Font Tags
      strStartTag = "<FONT="
      strEndTag = "</FONT>"
      iStart = InStr(1, .Text, strStartTag)
      iCount = 0
      
      Do While iStart > 0

         iLenST = Len(strStartTag)
         iEndTag = Len(strEndTag)
         
         iStart = InStr(1, .Text, strStartTag)
         
         If iStart = 0 Then Exit Do

         iEnd = InStr(iStart, .Text, strEndTag)

         strValue = Mid(.Text, (iStart + iLenST), InStr((iStart + iLenST), .Text, ">") - (iStart + iLenST))

         iStartTag = Len(strStartTag & strValue & ">")

         iLength = iEnd - iStartTag - iStart
         
         .SelStart = iEnd - 1
         .SelLength = iEndTag
         .SelText = ""
         
         .SelStart = iStart - 1
         .SelLength = iStartTag
         .SelText = ""

         .SelStart = iStart - 1
         .SelLength = iLength
         .SelFontName = strValue
         
         iCount = iCount + 1
         If iCount > 100 Then Exit Do
         
      Loop

      '-- Look For Font Size Tags
      strStartTag = "<SIZE="
      strEndTag = "</SIZE>"
      iStart = InStr(1, .Text, strStartTag)
      iCount = 0
      
      Do While iStart > 0

         iLenST = Len(strStartTag)
         iEndTag = Len(strEndTag)
         
         iStart = InStr(1, .Text, strStartTag)
         
         If iStart = 0 Then Exit Do

         iEnd = InStr(iStart, .Text, strEndTag)

         strValue = Mid(.Text, (iStart + iLenST), InStr((iStart + iLenST), .Text, ">") - (iStart + iLenST))

         iStartTag = Len(strStartTag & strValue & ">")

         iLength = iEnd - iStartTag - iStart
         
         .TextRTF = Replace(.TextRTF, strStartTag & strValue & ">", "", , 1)
         .TextRTF = Replace(.TextRTF, strEndTag, "", , 1)

         .SelStart = iStart - 1
         .SelLength = iLength
         .SelFontSize = CInt(strValue)
         
         iCount = iCount + 1
         If iCount > 100 Then Exit Do
         
      Loop
      
      '-- Font Colors
      strStartTag = "<COLOR="
      strEndTag = "</COLOR>"
      iStart = InStr(1, .Text, strStartTag)
      iCount = 0
      
      Do While iStart > 0
      
         iLenST = Len(strStartTag)
         iEndTag = Len(strEndTag)
         
         iStart = InStr(1, .Text, strStartTag)
         
         If iStart = 0 Then Exit Do

         iEnd = InStr(iStart, .Text, strEndTag)

         strValue = Mid(.Text, (iStart + iLenST), InStr((iStart + iLenST), .Text, ">") - (iStart + iLenST))

         iStartTag = Len(strStartTag & strValue & ">")

         iLength = iEnd - iStartTag - iStart
         
         .TextRTF = Replace(.TextRTF, strStartTag & strValue & ">", "", , 1)
         .TextRTF = Replace(.TextRTF, strEndTag, "", , 1)

         .SelStart = iStart - 1
         .SelLength = iLength
         .SelColor = CLng(strValue)
         
         iCount = iCount + 1
         If iCount > 100 Then Exit Do
         
      Loop
      
      '-- Alignment
      strStartTag = "<ALIGN="
      strEndTag = "</ALIGN>"
      iStart = InStr(1, .Text, strStartTag)
      iCount = 0
      
      Do While iStart > 0
      
         iLenST = Len(strStartTag)
         iEndTag = Len(strEndTag)
         
         iStart = InStr(1, .Text, strStartTag)
         
         If iStart = 0 Then Exit Do

         iEnd = InStr(iStart, .Text, strEndTag)

         strValue = Mid(.Text, (iStart + iLenST), InStr((iStart + iLenST), .Text, ">") - (iStart + iLenST))

         iStartTag = Len(strStartTag & strValue & ">")

         iLength = iEnd - iStartTag - iStart
         
         .TextRTF = Replace(.TextRTF, strStartTag & strValue & ">", "", , 1)
         .TextRTF = Replace(.TextRTF, strEndTag, "", , 1)

         .SelStart = iStart - 1
         .SelLength = iLength
         
         Select Case UCase(strValue)
            Case "LEFT"
               .SelAlignment = rtfLeft
            Case "RIGHT"
               .SelAlignment = rtfRight
            Case "CENTER"
               .SelAlignment = rtfCenter
         End Select
         
         iCount = iCount + 1
         If iCount > 100 Then Exit Do
         
      Loop
      
      '-- All Others
      
      ReDim strTags(4)
      strTags(0) = "B"
      strTags(1) = "U"
      strTags(2) = "I"
      strTags(3) = "STRIKE"
      strTags(4) = "BULLET"
      
      For i = LBound(strTags) To UBound(strTags)
      
         strStartTag = "<" & strTags(i) & ">"
         strEndTag = "</" & strTags(i) & ">"
         iStart = InStr(1, .Text, strStartTag)
         iCount = 0
         
         Do While iStart > 0
   
            iLenST = Len(strStartTag)
            iEndTag = Len(strEndTag)
            
            iStart = InStr(1, .Text, strStartTag)
            
            If iStart = 0 Then Exit Do
   
            iEnd = InStr(iStart, .Text, strEndTag)
   
            iStartTag = Len(strStartTag)
   
            iLength = iEnd - iStartTag - iStart
            
            .TextRTF = Replace(.TextRTF, strStartTag, "", , 1)
            .TextRTF = Replace(.TextRTF, strEndTag, "", , 1)
   
            .SelStart = iStart - 1
            .SelLength = iLength
            
            If i = 0 Then
               .SelBold = True
            ElseIf i = 1 Then
               .SelUnderline = True
            ElseIf i = 2 Then
               .SelItalic = True
            ElseIf i = 3 Then
               .SelStrikeThru = True
            ElseIf i = 4 Then
               .SelBullet = True
            End If
            
            iCount = iCount + 1
            If iCount > 100 Then Exit Do
            
         Loop
            
      Next
        
      .SelStart = 0
        
   End With
   
End Sub

Private Sub BuildButtons(Buttons As Long)
On Error Resume Next

   Dim i As Integer, iTemp As Integer
   
   With frmMsgBox
   
      Select Case Buttons
      
         Case vbOKOnly
            
            Button1.Visible = True
            Button1.Caption = "OK"
            i = 1
            
         Case vbOKCancel
            
            Button1.Visible = True
            Button1.Caption = "OK"
            Button2.Visible = True
            Button2.Caption = "Cancel"
            i = 2
            
         Case vbAbortRetryIgnore
            
            Button1.Visible = True
            Button1.Caption = "Abort"
            Button2.Visible = True
            Button2.Caption = "Retry"
            Button3.Visible = True
            Button3.Caption = "Ignore"
            i = 3
            
         Case vbYesNoCancel
         
            Button1.Visible = True
            Button1.Caption = "Yes"
            Button2.Visible = True
            Button2.Caption = "No"
            Button3.Visible = True
            Button3.Caption = "Cancel"
            i = 3
            
         Case vbYesNo
            
            Button1.Visible = True
            Button1.Caption = "Yes"
            Button2.Visible = True
            Button2.Caption = "No"
            i = 2
            
         Case vbRetryCancel
            
            Button1.Visible = True
            Button1.Caption = "Retry"
            Button2.Visible = True
            Button2.Caption = "Cancel"
            i = 2
            
      End Select
   
   End With
   
   Call AlignButtons(i)
   
End Sub

Private Sub BuildHelp()
On Error Resume Next

   button4.Visible = True
   button4.Caption = "Help"
   
End Sub

Private Sub BuildPicture(Picture As Long)
On Error Resume Next

   Select Case Picture
      
      Case vbCritical
         
         frmMsgBox.imgCritical.Visible = True
         
      Case vbQuestion
         
         frmMsgBox.imgQuestion.Visible = True
         
      Case vbExclamation
      
         frmMsgBox.imgExclaimation.Visible = True
      
      Case vbInformation
      
         frmMsgBox.imgInformation.Visible = True
      
      Case Else
         
         With frmMsgBox
            .txtMessage.Width = .Width - 400
            .txtMessage.Left = 200
         End With
      
   End Select
   
End Sub

Private Sub SetDefault(lDefault As Long)
On Error Resume Next

   With frmMsgBox
   
      Select Case lDefault
         
         Case vbDefaultButton1
            
            Button1.TabIndex = 0
            Button2.TabIndex = 1
            Button3.TabIndex = 2
            button4.TabIndex = 3
            
         Case vbDefaultButton2
            
            Button1.TabIndex = 3
            Button2.TabIndex = 0
            Button3.TabIndex = 1
            button4.TabIndex = 2
         
         Case vbDefaultButton3
            
            Button1.TabIndex = 2
            Button2.TabIndex = 3
            Button3.TabIndex = 0
            button4.TabIndex = 1
            
         Case vbDefaultButton4
            
            Button1.TabIndex = 1
            Button2.TabIndex = 2
            Button3.TabIndex = 3
            button4.TabIndex = 0
      
      End Select
   
   End With
   
End Sub

Private Sub AlignButtons(iAlign As Integer)
On Error Resume Next
   
   Dim iTemp As Integer
   Dim iWidth As Integer
   
   Select Case iAlign
      
      Case 1
         
         If lHelp = vbMsgBoxHelpButton Then
            
            If bAuto Then
               iWidth = Button1.Width
               iTemp = ((iWidth + 100) * 2) + 300
               If frmMsgBox.Width < iTemp Then
                  frmMsgBox.Width = iTemp
                  frmMsgBox.Frame1.Width = frmMsgBox.Width + 100
               End If
            End If
            
            Button1.Left = (frmMsgBox.Width / 2) - Button1.Width - 100
            button4.Left = Button1.Left + Button1.Width + 100
            
         Else
            
            If bAuto Then
               iWidth = Button1.Width
               iTemp = iWidth + 400
               If frmMsgBox.Width < iTemp Then
                  frmMsgBox.Width = iTemp
                  frmMsgBox.Frame1.Width = frmMsgBox.Width + 100
               End If
               
               Button1.Left = (frmMsgBox.Width / 2) - (Button1.Width / 2) - 30
                  
            Else
            
               Button1.Left = (frmMsgBox.Width / 2) - (Button1.Width / 2)
               
            End If
            
         End If
      
      Case 2
         
         If lHelp = vbMsgBoxHelpButton Then
            
            If bAuto Then
               iWidth = Button1.Width
               iTemp = ((iWidth + 100) * 3) + 400
               If frmMsgBox.Width < iTemp Then
                  frmMsgBox.Width = iTemp
                  frmMsgBox.Frame1.Width = frmMsgBox.Width + 100
               End If
            End If
            
            Button1.Left = (frmMsgBox.Width / 2) - (Button2.Width / 2) - button4.Width - 125
            Button2.Left = Button1.Left + Button1.Width + 100
            button4.Left = Button2.Left + Button2.Width + 100
            
         Else
            
            If bAuto Then
               iWidth = Button1.Width
               iTemp = ((iWidth + 100) * 2) + 300
               If frmMsgBox.Width < iTemp Then
                  frmMsgBox.Width = iTemp
                  frmMsgBox.Frame1.Width = frmMsgBox.Width + 100
               End If
            End If
            
            Button1.Left = (frmMsgBox.Width / 2) - Button2.Width - 100
            Button2.Left = Button1.Left + Button1.Width + 100
            
         End If
         
      Case 3
      
         If lHelp = vbMsgBoxHelpButton Then
         
            If bAuto Then
               iWidth = Button1.Width
               iTemp = ((iWidth + 100) * 4) + 400
               If frmMsgBox.Width < iTemp Then
                  frmMsgBox.Width = iTemp
                  frmMsgBox.Frame1.Width = frmMsgBox.Width + 100
               End If
            End If
         
            Button1.Left = (frmMsgBox.Width / 2) - Button1.Width - (Button2.Width / 2) - (Button3.Width / 2) - (Button3.Width / 8) - 50
            Button2.Left = Button1.Left + Button1.Width + 100
            Button3.Left = Button2.Left + Button2.Width + 100
            button4.Left = Button3.Left + Button3.Width + 100
            
         Else
            
            If bAuto Then
               iWidth = Button1.Width
               iTemp = ((iWidth + 100) * 3) + 300
               If frmMsgBox.Width < iTemp Then
                  frmMsgBox.Width = iTemp
                  frmMsgBox.Frame1.Width = frmMsgBox.Width + 100
               End If
            End If
         
            Button1.Left = (frmMsgBox.Width / 2) - (Button2.Width / 2) - Button3.Width - 150
            Button2.Left = Button1.Left + Button1.Width + 100
            Button3.Left = Button2.Left + Button2.Width + 100
            
         End If
         
   End Select
   
End Sub

Private Sub Class_Terminate()
   
   Set Button1 = Nothing
   Set Button2 = Nothing
   Set Button3 = Nothing
   Set button4 = Nothing
   Set frmMsgBox = Nothing
   
End Sub
