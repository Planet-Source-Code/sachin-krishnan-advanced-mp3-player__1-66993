VERSION 5.00
Begin VB.Form frmMusica 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "MP3 Player"
   ClientHeight    =   3675
   ClientLeft      =   0
   ClientTop       =   45
   ClientWidth     =   3870
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3675
   ScaleWidth      =   3870
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdLoadPlay 
      BackColor       =   &H00FFFF80&
      Caption         =   "L"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Load Playlist"
      Top             =   3240
      Width           =   375
   End
   Begin VB.CommandButton cmdLstSav 
      BackColor       =   &H00FFFF80&
      Caption         =   "S"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Save Playlist"
      Top             =   3240
      Width           =   375
   End
   Begin VB.CommandButton CmdLstClr 
      BackColor       =   &H00FFFF80&
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Clear Playlist"
      Top             =   3240
      Width           =   375
   End
   Begin VB.Timer PlayNxt 
      Interval        =   1
      Left            =   2520
      Top             =   2160
   End
   Begin VB.Timer TmrMarquee 
      Interval        =   100
      Left            =   1920
      Top             =   2160
   End
   Begin VB.CommandButton PosSlider 
      BackColor       =   &H00FF80FF&
      Height          =   135
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Position Slider"
      Top             =   1320
      Width           =   375
   End
   Begin VB.CommandButton cmdCont 
      BackColor       =   &H00000000&
      Height          =   615
      Index           =   6
      Left            =   240
      Picture         =   "MP3Player.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   0
      Width           =   615
   End
   Begin VB.CommandButton cmdCont 
      BackColor       =   &H00000000&
      Height          =   615
      Index           =   5
      Left            =   3000
      Picture         =   "MP3Player.frx":117A
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   0
      Width           =   615
   End
   Begin VB.CommandButton cmdCont 
      BackColor       =   &H00000000&
      Height          =   615
      Index           =   4
      Left            =   2880
      Picture         =   "MP3Player.frx":22F4
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   1560
      Width           =   615
   End
   Begin VB.CommandButton cmdCont 
      BackColor       =   &H00000000&
      Height          =   615
      Index           =   3
      Left            =   480
      Picture         =   "MP3Player.frx":346E
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1560
      Width           =   615
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   1320
      Top             =   2040
   End
   Begin VB.TextBox lbltxt 
      Height          =   285
      Left            =   3480
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   240
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox Lst 
      Height          =   450
      Left            =   2880
      TabIndex        =   7
      Top             =   120
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdCont 
      BackColor       =   &H00000000&
      Height          =   615
      Index           =   2
      Left            =   2280
      Picture         =   "MP3Player.frx":45E8
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1560
      Width           =   615
   End
   Begin VB.CommandButton cmdCont 
      BackColor       =   &H00000000&
      Height          =   615
      Index           =   1
      Left            =   1680
      Picture         =   "MP3Player.frx":5762
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1560
      Width           =   615
   End
   Begin VB.ListBox LstPlay 
      Appearance      =   0  'Flat
      Height          =   1005
      Left            =   120
      TabIndex        =   4
      Top             =   2160
      Width           =   3615
   End
   Begin VB.CommandButton CmdLstSub 
      BackColor       =   &H00FFFF80&
      Caption         =   "D"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Delete Selected item."
      Top             =   3240
      Width           =   375
   End
   Begin VB.CommandButton cmdLstAdd 
      BackColor       =   &H00FFFF80&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Add Files into the playlist."
      Top             =   3240
      Width           =   375
   End
   Begin VB.CommandButton cmdCont 
      BackColor       =   &H00000000&
      Height          =   615
      Index           =   0
      Left            =   1080
      Picture         =   "MP3Player.frx":68DC
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1560
      Width           =   615
   End
   Begin VB.PictureBox PosTrack 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   120
      ScaleHeight     =   105
      ScaleWidth      =   3585
      TabIndex        =   15
      ToolTipText     =   "Position Bar"
      Top             =   1320
      Width           =   3615
   End
   Begin VB.Label lblMarquee 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   1845
      TabIndex        =   16
      Top             =   720
      Width           =   165
   End
   Begin VB.Label ElaspTime 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "00:00"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   960
      Width           =   3615
   End
   Begin VB.Label Titleb 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "MUSICA"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3855
   End
End
Attribute VB_Name = "frmMusica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                       MY PLAYER
'                       DATE:- 02/10/06
'                       BY SACHIN K
'                       MY DAY : - 01/10/06
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Private Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long

Dim TFile As String
Dim FileN As String

Dim OldX As Integer
Dim OldY As Integer
Dim DragMode As Boolean
Dim MoveMe As Boolean
Dim i As Integer
Dim Paused As Boolean
Dim tSec As Long, tDur As Long
Dim n As Integer

Dim SlidDown As Boolean

Private Sub cmdCont_Click(Index As Integer)
On Error Resume Next
Select Case Index

    Case 0:  '' Play or resume
    If Paused = False Then
        lbltxt.Text = Lst.List(LstPlay.ListIndex)
        FileN = lbltxt.Text
        mciSendString "close " & TFile, 0&, 0&, 0&
        TFile$ = Chr$(34) + Trim(FileN$) + Chr$(34)
        mciSendString "open " & TFile, 0&, 0&, 0&
        mciSendString "play " & TFile, "", 0&, 0&
        Timer1.Enabled = True
    Else
        Paused = False
        mciSendString "play " & TFile, "", 0&, 0&
        Timer1.Enabled = True
    End If

    Case 1:  '' Pause or resume
    If Paused = False Then
        Paused = True
        TFile$ = Chr$(34) + Trim(FileN$) + Chr$(34)
        mciSendString "Stop " & TFile, 0, 0, 0
        tSec = 0
        tDur = 1
    Else
        Paused = False
        mciSendString "play " & TFile, "", 0&, 0&
    End If
    
    Case 2:  '' Stop
    TFile$ = Chr$(34) + Trim(FileN$) + Chr$(34)
    mciSendString "close " & TFile, 0&, 0&, 0&
    
    Case 3: '' Previous
    cmdCont_Click (2)
        If Not LstPlay.ListIndex < 1 Then
            LstPlay.Selected(LstPlay.ListIndex - 1) = True
            cmdCont_Click (0)
        End If
        
    Case 4: '' Next
    tSec = tDur
   
    Case 5: ''Exit
    cmdCont_Click (2)
    SavePlayList App.Path & "\musica.m3u", Lst
    End
    
    Case 6: ''About
    Form1.Show
    
End Select
End Sub

Private Sub cmdCont_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'On Error Resume Next
For i = 0 To 6
cmdCont(i).Picture = LoadPicture(App.Path & "\Buttons\" & i & ".bmp")
Next
cmdCont(Index).Picture = LoadPicture(App.Path & "\Buttons\" & Index & "_o.bmp")
End Sub

Private Sub cmdLoadPlay_Click()
On Error Resume Next
frmAdd.File1.Pattern = "*.m3u"
frmAdd.File1.Path = App.Path & "\Playlists"
frmAdd.Show

End Sub

Private Sub cmdLstAdd_Click()
frmAdd.File1.Pattern = "*.mp3;*.mp2;*.mp1;*.wma;*.wav;*.wmv;*.avi"
frmAdd.Show
End Sub

Private Sub CmdLstClr_Click()
LstPlay.Clear
Lst.Clear
End Sub

Private Sub cmdLstSav_Click()
Dim S As String
S = InputBox("Enter the name for your playlist.", "MUSICA")
SavePlayList App.Path & "\Playlists\" & S & ".m3u", Lst
End Sub

Private Sub CmdLstSub_Click()
Dim Pos As Integer
Pos = LstPlay.ListIndex
If Pos > -1 Then
LstPlay.RemoveItem (Pos)
Lst.RemoveItem (Pos)
End If
End Sub



Private Sub Form_Load()
ShapeForm
LoadControls
LoadPlaylist App.Path & "\musica.m3u", Lst, LstPlay
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Titleb.ForeColor = vbRed
For i = 0 To 6
cmdCont(i).Picture = LoadPicture(App.Path & "\Buttons\" & i & ".bmp")
Next
PosSlider.BackColor = &HFF80FF
End Sub

Private Sub LstPlay_Click()
lbltxt.Text = Lst.List(LstPlay.ListIndex)
End Sub

Private Sub LstPlay_DblClick()
lbltxt.Text = Lst.List(LstPlay.ListIndex)
cmdCont_Click (0)
End Sub

Private Sub LstPlay_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = vbKeyDelete Then Lst.RemoveItem (LstPlay.ListIndex): LstPlay.RemoveItem (LstPlay.ListIndex)
End Sub

Private Sub PlayNxt_Timer()
If LstPlay.ListCount > -1 And tSec > 0 And tSec = tDur Then
cmdCont_Click (2)
    If LstPlay.ListIndex <> LstPlay.ListCount - 1 Then
    LstPlay.Selected(LstPlay.ListIndex + 1) = True
    Else
    LstPlay.Selected(0) = True
    cmdCont_Click (0)
    End If
cmdCont_Click (0)
End If
End Sub

Private Sub PosSlider_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
SlidDown = True
OldX = X
OldY = Y
cmdCont_Click (1)
End Sub

Private Sub PosSlider_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
PosSlider.BackColor = vbGreen
If SlidDown = True Then PosSlider.Left = PosSlider.Left + (X - OldX)

If PosSlider.Left < 120 Then PosSlider.Left = 120
If PosSlider.Left > 3360 Then PosSlider.Left = 3360
End Sub

Private Sub PosSlider_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
PosSlider.Left = PosSlider.Left + (X - OldX)
SlidDown = False
If PosSlider.Left < 120 Then PosSlider.Left = 120
If PosSlider.Left > 3360 Then PosSlider.Left = 3360
ChangePositionTo (Int(PosSlider.Left / n))
cmdCont_Click (0)
End Sub

Private Sub PosTrack_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
cmdCont_Click (1)
PosSlider.Left = X
If PosSlider.Left < 120 Then PosSlider.Left = 120
If PosSlider.Left > 3360 Then PosSlider.Left = 3360
ChangePositionTo (Int(PosSlider.Left / n))
cmdCont_Click (0)
End Sub

Private Sub PosTrack_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
PosSlider.BackColor = vbGreen

End Sub

Private Sub Timer1_Timer()
On Error Resume Next
ElaspTime.Caption = Position & " / " & Duration

If tDur <> 0 Then
n = 3240 / tDur
PosSlider.Left = 120 + (tSec * Int(n))
End If
End Sub

Private Sub Titleb_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MoveMe = True
    OldX = X
    OldY = Y
End Sub

Private Sub Titleb_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Titleb.ForeColor = vbBlue
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

Sub LoadControls()
On Error Resume Next
For i = 0 To 6
cmdCont(i).Picture = LoadPicture(App.Path & "\Buttons\" & i & ".bmp")
Next
End Sub

Sub ShapeForm()
'' Form
Dim l As Long
'l = CreateEllipticRgn(50, 50, 500, 500)
l = CreateRoundRectRgn(0, 0, Me.Width / Screen.TwipsPerPixelX, Me.Height / Screen.TwipsPerPixelY, 100, 100)
SetWindowRgn Me.hwnd, l, True

'' List
l = CreateRoundRectRgn(0, 0, LstPlay.Width / Screen.TwipsPerPixelX, LstPlay.Height / Screen.TwipsPerPixelY, 30, 30)
SetWindowRgn LstPlay.hwnd, l, True

'' Controls
For i = 0 To 6
l = CreateEllipticRgn(10, 10, 35, 35)
SetWindowRgn cmdCont(i).hwnd, l, True
Next

'' Pos Slider
l = CreateRoundRectRgn(0, 0, PosSlider.Width / Screen.TwipsPerPixelX, PosSlider.Height / Screen.TwipsPerPixelY, 30, 30)
SetWindowRgn PosSlider.hwnd, l, True
l = CreateRoundRectRgn(0, 0, PosTrack.Width / Screen.TwipsPerPixelX, PosTrack.Height / Screen.TwipsPerPixelY, 30, 30)
SetWindowRgn PosTrack.hwnd, l, True

End Sub


Private Function Duration()
On Error Resume Next
Dim TotalTime As String * 128
Dim T As String
Dim lTotalTime As Long
TFile$ = Chr$(34) + Trim(FileN$) + Chr$(34)
    mciSendString "set " & TFile & " time format ms", TotalTime, 128, 0&
    mciSendString "status " & TFile & " length", TotalTime, 128, 0&

    mciSendString "set " & TFile & " time format frames", 0&, 0&, 0&
    
    lTotalTime = Val(TotalTime)
   T = GetThisTime(lTotalTime)
    Duration = T
    tDur = lTotalTime / 1000
End Function

Private Function Position()
Dim Sec As Long
Dim mins As Long
On Error Resume Next
Static S As String * 30
TFile$ = Chr$(34) + Trim(FileN$) + Chr$(34)
    mciSendString "set " & TFile & " time format milliseconds", 0, 0, 0
    mciSendString "status " & TFile & " position", S, Len(S), 0
    Sec = Round(Mid$(S, 1, Len(S)) / 1000)
    tSec = Val(Sec)
    If Sec < 60 Then Position = "0:" & Format(Sec, "00")
    If Sec > 59 Then
        mins = Int(Sec / 60)
        Sec = Sec - (mins * 60)
        Position = Format(mins, "00") & ":" & Format(Sec, "00")
    End If
End Function

Private Function GetThisTime(ByVal timein As Long) As String
    On Error Resume Next
    Dim conH As Integer
    Dim conM As Integer
    Dim conS As Integer
    Dim remTime As Long
    Dim strRetTime As String
    TheFile$ = Chr$(34) + Trim(File$) + Chr$(34)
    remTime = timein / 1000
    conH = Int(remTime / 3600)
    remTime = remTime Mod 3600
    conM = Int(remTime / 60)
    remTime = remTime Mod 60
    conS = remTime
    
    If conH > 0 Then
        strRetTime = Trim(Str(conH)) & ":"
    Else
        strRetTime = ""
    End If
    
    If conM >= 10 Then
        strRetTime = strRetTime & Trim(Str(conM))
    ElseIf conM > 0 Then
        strRetTime = strRetTime & "0" & Trim(Str(conM))
    Else
        strRetTime = strRetTime & "00"
    End If
    
    strRetTime = strRetTime & ":"
    
    If conS >= 10 Then
        strRetTime = strRetTime & Trim(Str(conS))
    ElseIf conS > 0 Then
        strRetTime = strRetTime & "0" & Trim(Str(conS))
    Else
        strRetTime = strRetTime & "00"
    End If
    
    GetThisTime = strRetTime
End Function

Private Sub TmrMarquee_Timer()
lblMarquee.Caption = FileN
If Not (lblMarquee.Left + lblMarquee.Width) < 0 Then
    lblMarquee.Left = lblMarquee.Left - 50
Else
    lblMarquee.Left = Me.Width + 100
End If
End Sub

Private Function Playing() As Boolean
On Error Resume Next
Static S As String * 30
mciSendString "status " & TFile & " mode", S, Len(S), 0
Playing = (Mid$(S, 1, 7) = "playing")
End Function

Private Function ChangePositionTo(Second)
On Error Resume Next
TFile$ = Chr$(34) + Trim(FileN$) + Chr$(34)
Second = Second * 1000
mciSendString "set time format milliseconds", 0, 0, 0
If Playing = True Then
mciSendString "play " & TFile & " from " & Second, 0, 0, 0
ElseIf Playing = False Then
mciSendString "seek " & TFile & " to " & Second, 0, 0, 0
End If
End Function

Private Function SavePlayList(FileName As String, OrgListBox)
Dim i As Integer
Dim a As String
On Error Resume Next
Open FileName For Output As #1
For i = 0 To OrgListBox.ListCount - 1
a$ = OrgListBox.List(i)
Print #1, a$
Next
Close 1
End Function

Private Function LoadPlaylist(FileName As String, OrgLB As ListBox, VisLB As ListBox)
On Error Resume Next
Dim S As String
Dim Pos As Integer
Dim STRF As String, i As Integer

If FileName = "" Then Exit Function

Open FileName For Input As 1

While Not EOF(1)
Line Input #1, S
OrgLB.AddItem RTrim(S)

STRF = S
For i = 1 To Len(STRF)
If Mid(STRF, i, 1) = "\" Then Pos = i
Next
VisLB.AddItem Mid(STRF, Pos + 1)
Wend
Close 1
End Function
