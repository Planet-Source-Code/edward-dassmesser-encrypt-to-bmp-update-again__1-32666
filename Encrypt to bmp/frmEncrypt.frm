VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   Caption         =   "Picture Encryption"
   ClientHeight    =   6120
   ClientLeft      =   1035
   ClientTop       =   1320
   ClientWidth     =   13740
   Icon            =   "frmEncrypt.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6120
   ScaleWidth      =   13740
   Begin VB.CommandButton cmdLoadBMP 
      Caption         =   "Load .BMP"
      Height          =   195
      Left            =   12360
      TabIndex        =   7
      Top             =   5880
      Width           =   1095
   End
   Begin VB.CommandButton cmdSaveBMP 
      Caption         =   "Save .BMP"
      Height          =   195
      Left            =   11040
      TabIndex        =   6
      Top             =   5880
      Width           =   1095
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save to Text"
      Height          =   195
      Left            =   9480
      TabIndex        =   5
      Top             =   5880
      Width           =   1095
   End
   Begin MSComDlg.CommonDialog loadfile 
      Left            =   4080
      Top             =   5880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load File"
      Height          =   195
      Left            =   600
      TabIndex        =   4
      Top             =   5880
      Width           =   975
   End
   Begin VB.CommandButton cmdDecrypt 
      Caption         =   "Decrypt Text"
      Height          =   195
      Left            =   8280
      TabIndex        =   3
      Top             =   5880
      Width           =   1095
   End
   Begin VB.CommandButton cmdEncrypt 
      Caption         =   "Encrypt Text"
      Height          =   195
      Left            =   2160
      TabIndex        =   1
      Top             =   5880
      Width           =   1215
   End
   Begin VB.PictureBox picEncrypted 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   5775
      Left            =   7080
      ScaleHeight     =   383
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   439
      TabIndex        =   2
      Top             =   0
      Width           =   6615
   End
   Begin VB.TextBox txtEncrypt 
      Height          =   5775
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   0
      Width           =   7095
   End
   Begin VB.Menu mnuStuff 
      Caption         =   "&Stuff you might want to do"
      Begin VB.Menu mnuVote 
         Caption         =   "&Vote at Planet Source Code"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuEmail 
         Caption         =   "E-mail Author (durnurd@hotmail.com)"
         Shortcut        =   ^E
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Decrypt()
    Dim txtLength As Long, X As Integer, Y As Integer
    Dim picWidth As Long, picHeight As Long, currPixel As Long, Red As Integer
    Dim Green As Integer, Blue As Integer
    picWidth = picEncrypted.ScaleWidth: picHeight = picWidth
    txtEncrypt.Text = ""
    For X = 1 To picWidth
        For Y = 1 To picHeight
            currPixel = GetPixel(picEncrypted.hdc, X, Y)
            Red = 0
            Blue = 0
            Green = 0
            If currPixel = -1 Then GoTo MidDecrypt
            Blue = Int(currPixel / 65536)
            If Blue = 255 Then Exit For
            Green = Int((currPixel - (Blue * 65536)) / 256)
            If Green = 255 Then Exit For
            Red = Int(currPixel - Blue * 65536 - Green * 256)
            If Red = 255 Then Exit For
            txtEncrypt.Text = txtEncrypt.Text & Chr$(Red) & Chr$(Green) & Chr$(Blue)
MidDecrypt: Next Y
    Next X
End Sub
Private Sub Encrypt()
    Dim txtLength As Long, X As Integer, Y As Integer
    Dim picWidth As Long, picHeight As Long, currPixel As Long, Red As Integer
    Dim Green As Integer, Blue As Integer
    txtEncrypt.Text = Replace(txtEncrypt.Text, vbCrLf, vbLf)
    txtLength = Len(txtEncrypt.Text)
    picWidth = Int(Sqr(txtLength / 2)): picHeight = picWidth
    currPixel = 0
    picEncrypted.Width = picWidth * Screen.TwipsPerPixelX + 100
    picEncrypted.Height = picHeight * Screen.TwipsPerPixelY + 100
    For X = 1 To picWidth
        For Y = 1 To picHeight
            currPixel = currPixel + 1
            Red = 0
            Green = 0
            Blue = 0
            If currPixel > txtLength Then Exit For
            Red = Asc(Mid(txtEncrypt.Text, currPixel, 1))
            currPixel = currPixel + 1
            If currPixel > txtLength Then GoTo MidColor
            Green = Asc(Mid(txtEncrypt.Text, currPixel, 1))
            currPixel = currPixel + 1
            If currPixel > txtLength Then GoTo MidColor
            Blue = Asc(Mid(txtEncrypt.Text, currPixel, 1))
MidColor:   SetPixelV picEncrypted.hdc, X, Y, RGB(Red, Green, Blue)
        Next Y
    Next X
End Sub
Private Sub cmdEncrypt_Click()
    picEncrypted.Cls
    picEncrypted.Picture = LoadPicture("")
    Encrypt
    SavePicture picEncrypted.Image, "C:\Encrypted File.bmp"
End Sub

Private Sub cmdDecrypt_Click()
    On Error Resume Next
    picEncrypted.Cls
    picEncrypted.Picture = LoadPicture("C:\Encrypted File.bmp")
    Decrypt
End Sub

Private Sub cmdLoad_Click()
    Dim TxtLoaded As String, TempTxt As String
    On Error Resume Next
    loadfile.ShowOpen
    Open loadfile.FileName For Input As #1
        Do While Not EOF(1)
            Line Input #1, TempTxt
            TxtLoaded = TxtLoaded & vbCrLf & TempTxt
        Loop
    Close #1
    TxtLoaded = Right(TxtLoaded, Len(TxtLoaded) - 2)
    picEncrypted.Cls
    picEncrypted.Picture = LoadPicture("")
    txtEncrypt.Text = TxtLoaded
    Encrypt
    SavePicture picEncrypted.Image, "C:\Encrypted File.bmp"
End Sub

Private Sub cmdLoadBMP_Click()
    picEncrypted.Cls
    loadfile.Filter = "*.bmp|*.bmp"
    loadfile.ShowOpen
    If loadfile.FileName = "" Then Exit Sub
    picEncrypted.Picture = LoadPicture(loadfile.FileName)
    loadfile.Filter = ""
    Decrypt
End Sub

Private Sub cmdSave_Click()
    txtSaved As String
    loadfile.ShowSave
    Set P = picEncrypted
    On Error Resume Next
    picEncrypted.Cls
    picEncrypted.Picture = LoadPicture("C:\Encrypted File.bmp")
    Open loadfile.FileName For Output As #1
    Decrypt
    txtSaved = txtEncrypt
    Print #1, txtSaved
    Close #1
End Sub

Private Sub cmdSaveBMP_Click()
    loadfile.Filter = "*.bmp|*.bmp"
    loadfile.ShowSave
    If loadfile.FileName = "" Then Exit Sub
    SavePicture picEncrypted.Image, loadfile.FileName
    loadfile.Filter = ""
End Sub

Private Sub mnuEmail_Click()
    StartURL "mailto:durnurd@hotmail.com"
End Sub
Private Sub mnuVote_Click()
    StartURL "http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=32625&lngWId=1"
End Sub
Private Sub StartURL(strURL As String)
    On Error Resume Next
    Shell "Explorer """ & strURL & """"
    If Err.Number <> 0 Then
        Err.Clear
        Shell "Start """ & strURL & """"
    End If
    If Err.Number <> 0 Then
        If MsgBox("Can't figure out how to navigate on this OS.  Copy the URL to the clipboard?", vbExclamation + vbYesNo) = vbYes Then
            Clipboard.SetText strURL
        End If
    End If
End Sub
