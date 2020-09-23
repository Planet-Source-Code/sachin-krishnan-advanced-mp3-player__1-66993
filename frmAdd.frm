VERSION 5.00
Begin VB.Form frmAdd 
   BackColor       =   &H0080FF80&
   BorderStyle     =   0  'None
   Caption         =   "Add File/Directory"
   ClientHeight    =   6570
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6540
   LinkTopic       =   "Form2"
   ScaleHeight     =   6570
   ScaleWidth      =   6540
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CmdXit 
      BackColor       =   &H0000C000&
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5760
      Width           =   1335
   End
   Begin VB.CommandButton cmdAddDir 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Add"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Adds all the files in the current directory to the playlist."
      Top             =   3240
      Width           =   855
   End
   Begin VB.ListBox LstOrg 
      Height          =   450
      Left            =   240
      TabIndex        =   6
      Top             =   5880
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton AddFile 
      BackColor       =   &H00FF80FF&
      Caption         =   "Add these Files"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5760
      Width           =   1935
   End
   Begin VB.ListBox LstFiles 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      Height          =   1395
      Left            =   960
      TabIndex        =   3
      Top             =   4320
      Width           =   3375
   End
   Begin VB.FileListBox File1 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      Height          =   1395
      Left            =   960
      MultiSelect     =   2  'Extended
      Pattern         =   "*.mp3;*.mp2;*.mp1;*.wma;*.wav;*.wmv;*.avi"
      TabIndex        =   2
      Top             =   2880
      Width           =   3375
   End
   Begin VB.DirListBox Dir1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      Height          =   1665
      Left            =   960
      TabIndex        =   1
      Top             =   1200
      Width           =   3375
   End
   Begin VB.DriveListBox Drive1 
      BackColor       =   &H008080FF&
      Height          =   315
      Left            =   1200
      TabIndex        =   0
      Top             =   840
      Width           =   2895
   End
   Begin VB.Label Titleb 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "ADD FILES"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   465
      Left            =   0
      TabIndex        =   4
      Top             =   240
      Width           =   5370
   End
End
Attribute VB_Name = "frmAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long

Dim OldX As Integer
Dim OldY As Integer
Dim MoveMe As Boolean
Sub ShapeForm()
Dim l As Long
l = CreateEllipticRgn(5, 5, 350, 475)
SetWindowRgn Me.hwnd, l, True

l = CreateEllipticRgn(10, 10, 50, 50)
SetWindowRgn cmdAddDir.hwnd, l, True

l = CreateEllipticRgn(2, 2, 85, 49)
SetWindowRgn CmdXit.hwnd, l, True

l = CreateEllipticRgn(2, 2, 125, 49)
SetWindowRgn AddFile.hwnd, l, True

l = CreateRoundRectRgn(0, 0, Drive1.Width / Screen.TwipsPerPixelX, Drive1.Height / Screen.TwipsPerPixelY, 30, 30)
SetWindowRgn Drive1.hwnd, l, True

l = CreateRoundRectRgn(2, 1, Dir1.Width / Screen.TwipsPerPixelX, Dir1.Height / Screen.TwipsPerPixelY, 30, 30)
SetWindowRgn Dir1.hwnd, l, True

l = CreateRoundRectRgn(2, 1, File1.Width / Screen.TwipsPerPixelX, File1.Height / Screen.TwipsPerPixelY, 30, 30)
SetWindowRgn File1.hwnd, l, True

l = CreateRoundRectRgn(2, 1, LstFiles.Width / Screen.TwipsPerPixelX, LstFiles.Height / Screen.TwipsPerPixelY, 30, 30)
SetWindowRgn LstFiles.hwnd, l, True
End Sub

Private Sub AddFile_Click()
Dim i As Integer
For i = 0 To LstOrg.ListCount - 1
frmMusica.Lst.AddItem LstOrg.List(i)
frmMusica.LstPlay.AddItem LstFiles.List(i)
Next
Unload Me
End Sub

Private Sub AddFile_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
AddFile.BackColor = vbRed
End Sub

Private Sub cmdAddDir_Click()
Dim i As Integer
For i = 0 To File1.ListCount - 1
    File1.Selected(i) = True
    File1_DblClick
Next

End Sub

Private Sub cmdAddDir_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdAddDir.BackColor = vbRed
End Sub

Private Sub CmdXit_Click()
Unload Me
End Sub

Private Sub CmdXit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
CmdXit.BackColor = vbRed
End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
End Sub



Private Sub File1_DblClick()
Dim FN As String
If Right(File1.Path, 1) = "\" Then
FN = File1.Path & File1.FileName
Else
FN = File1.Path & "\" & File1.FileName
End If

If File1.Pattern = "*.m3u" Then
    LoadPlaylist FN, LstOrg, LstFiles
Else
    LstOrg.AddItem FN
    Dim Pos As Integer
    Dim STRF As String, i As Integer
    STRF = File1.List(File1.ListIndex)
    For i = 1 To Len(STRF)
    If Mid(STRF, i, 1) = "\" Then Pos = i
    Next
    LstFiles.AddItem Mid(STRF, Pos + 1)
End If
End Sub



Private Sub Form_Load()
ShapeForm
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdAddDir.BackColor = &HFFC0C0
CmdXit.BackColor = &HC000&
AddFile.BackColor = &HFF80FF
Titleb.ForeColor = vbBlue
End Sub

Private Sub LstFiles_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = vbKeyDelete Then LstOrg.RemoveItem (LstFiles.ListIndex): LstFiles.RemoveItem (LstFiles.ListIndex)
End Sub

Private Sub Titleb_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MoveMe = True
    OldX = X
    OldY = Y
End Sub

Private Sub Titleb_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Titleb.ForeColor = &H80FF&
    If MoveMe = True Then
        Me.Left = Me.Left + (X - OldX)
        Me.Top = Me.Top + (Y - OldY)
    End If
End Sub

Private Sub Titleb_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.Left = Me.Left + (X - OldX)
    Me.Top = Me.Top + (Y - OldY)
    MoveMe = False
End Sub

Private Function LoadPlaylist(FileName As String, OrgLB As ListBox, VisLB As ListBox)
On Error Resume Next
Dim s As String
Dim Pos As Integer
Dim STRF As String, i As Integer

If FileName = "" Then Exit Function

Open FileName For Input As 1

While Not EOF(1)
Line Input #1, s
OrgLB.AddItem RTrim(s)

STRF = s
For i = 1 To Len(STRF)
If Mid(STRF, i, 1) = "\" Then Pos = i
Next
VisLB.AddItem Mid(STRF, Pos + 1)
Wend
Close 1
End Function
