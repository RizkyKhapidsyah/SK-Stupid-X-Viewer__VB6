VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Stupid XViewer"
   ClientHeight    =   5160
   ClientLeft      =   495
   ClientTop       =   840
   ClientWidth     =   7065
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5160
   ScaleWidth      =   7065
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   4815
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6735
      Begin VB.VScrollBar VScroll1 
         Height          =   2895
         LargeChange     =   100
         Left            =   4920
         SmallChange     =   25
         TabIndex        =   1
         Top             =   1200
         Width           =   255
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   255
         LargeChange     =   100
         Left            =   1080
         SmallChange     =   25
         TabIndex        =   2
         Top             =   4080
         Width           =   3855
      End
      Begin VB.CommandButton Command1 
         CausesValidation=   0   'False
         Enabled         =   0   'False
         Height          =   255
         Left            =   4920
         TabIndex        =   3
         Top             =   4080
         Width           =   255
      End
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   2655
         Left            =   240
         ScaleHeight     =   177
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   289
         TabIndex        =   4
         Top             =   360
         Width           =   4335
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileSp1 
         Caption         =   "-- File --"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuFileView 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileSp2 
         Caption         =   "-- Options --"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuFileClear 
         Caption         =   "C&lear"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuFileCut 
         Caption         =   "C&ut"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuFileCopy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuFilePaste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuFileSp3 
         Caption         =   "-- Help and Exit --"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuFileAbout 
         Caption         =   "&About"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private trytoexit As Boolean 'Enable user to unload the main form

'API Call that enables you to get Window Control Box Menu
Private Declare Function GetSystemMenu Lib "user32" _
(ByVal hWnd As Long, ByVal bRevert As Long) As Long

'API Call that enables you to delete any menu
Private Declare Function DeleteMenu Lib "user32" _
(ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long

'Flags for current menu position
Private Const MF_BYPOSITION = &H400&

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

 Select Case KeyCode
  
  Case vbKeyLeft:
   Me.HScroll1.SetFocus
    
  Case vbKeyRight:
    Me.HScroll1.SetFocus
   
  Case vbKeyUp:
    Me.VScroll1.SetFocus
  
  Case vbKeyDown:
   Me.VScroll1.SetFocus
   
  Case vbKeyPageDown:
   Me.VScroll1.SetFocus
  
  Case vbKeyPageUp:
   Me.VScroll1.SetFocus
   
 End Select
 
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

 Select Case KeyCode
  
  Case vbKeyLeft:
   Me.HScroll1.SetFocus
    
  Case vbKeyRight:
    Me.HScroll1.SetFocus
   
  Case vbKeyUp:
    Me.VScroll1.SetFocus
  
  Case vbKeyDown:
   Me.VScroll1.SetFocus
   
  Case vbKeyPageDown:
   Me.VScroll1.SetFocus
  
  Case vbKeyPageUp:
   Me.VScroll1.SetFocus
    
 End Select
 
End Sub

Private Sub Form_Load()

 Dim sysmenu As Long
 
 trytoexit = False
 sysmenu = GetSystemMenu(Me.hWnd, False)
 
 'Disable the close menu and the separator
 DeleteMenu sysmenu, 6, MF_BYPOSITION
 DeleteMenu sysmenu, 5, MF_BYPOSITION
 
 With Me
 
  .Frame1.Move 0, 0
  .Frame1.Width = Me.ScaleWidth - 100
  .Frame1.Height = Me.ScaleHeight - 100
  
  .Picture1.Move 0, 0
  .Picture1.AutoSize = False
  .Picture1.ScaleMode = vbPixels
  
  .VScroll1.Max = .Picture1.Height
  .VScroll1.Value = 0
  
  .HScroll1.Max = .Picture1.Width
  .HScroll1.Value = 0
  
  .VScroll1.Move .Frame1.Width - .VScroll1.Width, .Frame1.Top, _
                 .VScroll1.Width, .Frame1.Height - .Command1.Height
                 
  .HScroll1.Move .Frame1.Left, .Frame1.Height - .HScroll1.Height, _
                 .Frame1.Width - .Command1.Width, .HScroll1.Height
                 
  .Command1.Move .Frame1.Width - .Command1.Width, .Frame1.Height - .Command1.Height
  
 End With
 
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

 If trytoexit Then Cancel = False
 If Not trytoexit Then Cancel = True
 
End Sub

Private Sub Form_Resize()

 On Error Resume Next
 
 With Me
 
  .Frame1.Move 0, 0
  .Frame1.Width = Me.ScaleWidth - 100
  .Frame1.Height = Me.ScaleHeight - 100
  
  .Picture1.Move 0, 0
  .Picture1.AutoSize = True
  .Picture1.ScaleMode = vbPixels
  
  .VScroll1.Max = .Picture1.Height
  .VScroll1.Value = 0
  
  .HScroll1.Max = .Picture1.Width
  .HScroll1.Value = 0
  
  .VScroll1.Move .Frame1.Width - .VScroll1.Width, .Frame1.Top, _
                 .VScroll1.Width, .Frame1.Height - .Command1.Height
                 
  .HScroll1.Move .Frame1.Left, .Frame1.Height - .HScroll1.Height, _
                 .Frame1.Width - .Command1.Width, .HScroll1.Height
                 
  .Command1.Move .Frame1.Width - .Command1.Width, .Frame1.Height - .Command1.Height
  
 End With

End Sub

Private Sub Frame1_Click()
 
 Me.Picture1.SetFocus
 
End Sub

Private Sub HScroll1_Change()
 
 On Error Resume Next
 
 With Me
   
   .Picture1.Left = -.HScroll1.Value
   
 End With
 
End Sub

Private Sub HScroll1_GotFocus()

 Me.SetFocus
 
End Sub

Private Sub HScroll1_KeyDown(KeyCode As Integer, Shift As Integer)

 Select Case KeyCode
  
  Case vbKeyLeft:
  Me.HScroll1.SetFocus
    
  Case vbKeyRight:
    Me.HScroll1.SetFocus
   
  Case vbKeyUp:
    Me.VScroll1.SetFocus
  
  Case vbKeyDown:
   Me.VScroll1.SetFocus
 
  Case vbKeyPageDown:
   Me.VScroll1.SetFocus
  
  Case vbKeyPageUp:
   Me.VScroll1.SetFocus
      
 End Select
 
End Sub

Private Sub mnuExit_Click()
 
 trytoexit = True
 
 End
 
End Sub


Private Sub mnuFile_Click()
  
  If Me.Picture1.Picture.Handle = 0 Then
   
   Me.mnuFileCopy.Enabled = False
   Me.mnuFileSave.Enabled = False
   Me.mnuFileCut.Enabled = False
   Me.mnuFileClear.Enabled = False
   
  Else
   
   Me.mnuFileCopy.Enabled = True
   Me.mnuFileSave.Enabled = True
   Me.mnuFileCut.Enabled = True
   Me.mnuFileClear.Enabled = True
   
  End If

  If Clipboard.GetData Then
    
    Me.mnuFilePaste.Enabled = True
  
  Else
  
    Me.mnuFilePaste.Enabled = False
  
  End If
   
End Sub

Private Sub mnuFileAbout_Click()

Dim msg As String
msg = " Brought for free by ©MasterX Artwork"
msg = msg & vbCrLf & " Version " & App.Major & "." & App.Minor & vbCrLf
msg = msg & " E-mail: gonejoe@hotmail.com" & vbCrLf
msg = msg & " Web: http://www.geocities.com/m_bachok" & vbCrLf
msg = msg & " Happy Viewing :)"

MsgBox msg, vbInformation, "About Stupid Xviewer"

End Sub

Private Sub mnuFileClear_Click()

 If Me.Picture1.Picture.Handle = 0 Then Exit Sub
 
 Me.Picture1.Picture = Nothing
 Me.Picture1.AutoSize = False
 Me.VScroll1.Value = 0
 Me.HScroll1.Value = 0
 Me.Caption = "Stupid Viewer"
 
End Sub

Private Sub mnuFileCopy_Click()

 If Me.Picture1.Picture.Handle = 0 Then Exit Sub

 If Me.Picture1.Picture.Type = vbPicTypeIcon Then
  MsgBox "Cannot copy Icon", vbCritical, Err.Description
  Exit Sub
 End If
 
 If Me.Picture1.Picture.Type = vbPicTypeBitmap Then
   Clipboard.Clear
   Clipboard.SetData Me.Picture1.Picture, vbCFBitmap
 End If
 
 If Me.Picture1.Picture.Type = vbPicTypeMetafile Then
   Clipboard.Clear
   Clipboard.SetData Me.Picture1.Picture, vbCFMetafile
 End If
 
 If Me.Picture1.Picture.Type = vbPicTypeEMetafile Then
   Clipboard.Clear
   Clipboard.SetData Me.Picture1.Picture, vbCFEMetafile
 End If

End Sub

Private Sub mnuFileCut_Click()

 If Me.Picture1.Picture.Handle = 0 Then Exit Sub
 
 If Me.Picture1.Picture.Type = vbPicTypeIcon Then
  MsgBox "Cannot copy Icon. Save as bitmap first.", vbCritical, Err.Description
  Exit Sub
 End If
 
 If Me.Picture1.Picture.Type = vbPicTypeBitmap Then
   Clipboard.Clear
   Clipboard.SetData Me.Picture1.Picture, vbCFBitmap
 End If
 
 If Me.Picture1.Picture.Type = vbPicTypeMetafile Then
   Clipboard.Clear
   Clipboard.SetData Me.Picture1.Picture, vbCFMetafile
 End If
 
 If Me.Picture1.Picture.Type = vbPicTypeEMetafile Then
   Clipboard.Clear
   Clipboard.SetData Me.Picture1.Picture, vbCFEMetafile
 End If

 Me.Picture1.Picture = Nothing
 Me.Picture1.AutoSize = False
 Me.Caption = "Stupid Viewer"
 Me.VScroll1.Value = 0
 Me.HScroll1.Value = 0
 
End Sub

Private Sub mnuFilePaste_Click()

On Error Resume Next

If Clipboard.GetData Then
 With Me
 'Load the image file from clipboard
 Me.Picture1.Picture = Clipboard.GetData
 Me.Picture1.AutoSize = True
 Me.Caption = "Stupid Viewer - " & _
              CLng(Me.Picture1.ScaleY(Me.Picture1.Picture.Height)) & _
             "x" & CLng(Me.Picture1.ScaleX(Me.Picture1.Picture.Width)) & _
             " pixels"
             
   Me.Picture1.Move 0, 0
  
   Me.VScroll1.Max = .Picture1.Height
  Me.VScroll1.Value = 0
  
   Me.HScroll1.Max = .Picture1.Width
   Me.HScroll1.Value = 0
   
 End With
End If

End Sub

Private Sub mnuFileSave_Click()

 On Error GoTo ed
 
 Dim cmddlgSave As New clsOpenSave
 
 If Me.Picture1.Picture.Handle = 0 Then Exit Sub
 
 With cmddlgSave

 If Me.Picture1.Picture.Type = vbPicTypeIcon Then
 
  .CancelError = True
  .DialogTitle = "Save Icon"
  .FileName = "Untitled"
  .Filter = "Icon|*.ico"
  .hWnd = 0
  .Flags = 4 Or OFN_OVERWRITEPROMPT
  .InitDir = ""
  .ShowSave
  If .FileName <> "" Then SavePicture Me.Picture1.Picture, .FileName
  Exit Sub
   
 End If
  
  If Me.Picture1.Picture.Type = vbPicTypeMetafile Then
 
  .CancelError = True
  .DialogTitle = "Save MetaFile"
  .FileName = "Untitled"
  .Filter = "Windows Metafile|*.wmf"
  .hWnd = 0
  .Flags = 4 Or OFN_OVERWRITEPROMPT
  .InitDir = ""
  .ShowSave
  If .FileName <> "" Then SavePicture Me.Picture1.Picture, .FileName
  Exit Sub
   
 End If
 
 If Me.Picture1.Picture.Type = vbPicTypeEMetafile Then
 
  .CancelError = True
  .DialogTitle = "Save Enhanced MetaFile"
  .FileName = "Untitled"
  .Filter = "Enhanced Metafile|*.emf"
  .hWnd = 0
  .Flags = 4 Or OFN_OVERWRITEPROMPT
  .InitDir = ""
  .ShowSave
  If .FileName <> "" Then SavePicture Me.Picture1.Picture, .FileName
  Exit Sub
   
 End If

 If Me.Picture1.Picture.Type = vbPicTypeBitmap Then
 
  .CancelError = True
  .DialogTitle = "Save Bitmap Graphics"
  .FileName = "Untitled"
  .Filter = "Bitmap|*.bmp"
  .hWnd = 0
  .Flags = 4 Or OFN_OVERWRITEPROMPT
  .InitDir = ""
  .ShowSave
  If .FileName <> "" Then SavePicture Me.Picture1.Picture, .FileName
  Exit Sub
  
 End If
 
 End With

ed:
 If Err.Number = 32755 Then Exit Sub
 MsgBox "Error while saving file", vbCritical, Err.Description
 Exit Sub
 
End Sub



Private Sub mnuFileView_Click()
 
 On Error GoTo ed
 
 Dim cmddlg As New clsOpenSave
 
 With cmddlg
 
  .DialogTitle = "Open Graphics Files"
  .CancelError = False
  .FileName = ""
  .hWnd = 0
  .Filter = "All Supported Graphics|*.dib;*.bmp;*.jpg;*.jpeg;*.jfif;" & _
            "*.gif;*.cur;*.ico;*.icl;*.wmf;*.emf"
  .Flags = 4 Or OFN_FILEMUSTEXIST
  .InitDir = ""
  .ShowOpen
  If .FileName <> "" Then
  
    Me.Picture1.Picture = LoadPicture(.FileName, 0, 0, 0, 0)
    Me.Picture1.AutoSize = True
    Me.Caption = "Stupid Viewer - " & _
              CLng(Me.Picture1.ScaleY(Me.Picture1.Picture.Height)) & _
             "x" & CLng(Me.Picture1.ScaleX(Me.Picture1.Picture.Width)) & _
             " pixels"
 With Me
 
  .VScroll1.Max = .Picture1.Height
  .VScroll1.Value = 0
  
  .HScroll1.Max = .Picture1.Width
  .HScroll1.Value = 0
  
    Me.VScroll1.Value = 0
    Me.HScroll1.Value = 0
    
  End With
  
  End If
  Exit Sub
  
 End With
 
ed:
  MsgBox "Error loading file.", vbCritical, Err.Description
  Exit Sub
End Sub

Private Sub Picture1_KeyDown(KeyCode As Integer, Shift As Integer)

 Select Case KeyCode
  
  Case vbKeyLeft:
   Me.HScroll1.SetFocus
    
  Case vbKeyRight:
    Me.HScroll1.SetFocus
   
  Case vbKeyUp:
    Me.VScroll1.SetFocus
  
  Case vbKeyDown:
   Me.VScroll1.SetFocus
   
  Case vbKeyPageDown:
   Me.VScroll1.SetFocus
  
  Case vbKeyPageUp:
   Me.VScroll1.SetFocus
    
 End Select
 
End Sub

Private Sub VScroll1_Change()
 
 On Error Resume Next
 
 With Me
   
   .Picture1.Top = -.VScroll1.Value
   
 End With
 
End Sub

Private Sub VScroll1_GotFocus()

 Me.SetFocus
 
End Sub

Private Sub VScroll1_KeyDown(KeyCode As Integer, Shift As Integer)

 Select Case KeyCode
  
  Case vbKeyLeft:
  Me.HScroll1.SetFocus
    
  Case vbKeyRight:
    Me.HScroll1.SetFocus
   
  Case vbKeyUp:
    Me.VScroll1.SetFocus
  
  Case vbKeyDown:
   Me.VScroll1.SetFocus
   
  Case vbKeyPageDown:
   Me.VScroll1.SetFocus
  
  Case vbKeyPageUp:
   Me.VScroll1.SetFocus
    
   
 End Select
 
End Sub

