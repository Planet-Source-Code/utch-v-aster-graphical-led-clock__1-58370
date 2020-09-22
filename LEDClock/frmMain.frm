VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   10185
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10035
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10185
   ScaleWidth      =   10035
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picColon 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   180
      Index           =   1
      Left            =   1110
      Picture         =   "frmMain.frx":0000
      ScaleHeight     =   180
      ScaleWidth      =   105
      TabIndex        =   7
      Top             =   60
      Width           =   105
   End
   Begin VB.PictureBox picColon 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   180
      Index           =   0
      Left            =   540
      Picture         =   "frmMain.frx":0123
      ScaleHeight     =   180
      ScaleWidth      =   105
      TabIndex        =   6
      Top             =   60
      Width           =   105
   End
   Begin VB.PictureBox picSec 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   1
      Left            =   1440
      Picture         =   "frmMain.frx":0246
      ScaleHeight     =   195
      ScaleWidth      =   255
      TabIndex        =   5
      Top             =   60
      Width           =   255
   End
   Begin VB.PictureBox picSec 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   0
      Left            =   1200
      Picture         =   "frmMain.frx":0480
      ScaleHeight     =   195
      ScaleWidth      =   255
      TabIndex        =   4
      Top             =   60
      Width           =   255
   End
   Begin VB.PictureBox picMin 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   1
      Left            =   870
      Picture         =   "frmMain.frx":06BA
      ScaleHeight     =   195
      ScaleWidth      =   255
      TabIndex        =   3
      Top             =   60
      Width           =   255
   End
   Begin VB.PictureBox picMin 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   0
      Left            =   630
      Picture         =   "frmMain.frx":08F4
      ScaleHeight     =   195
      ScaleWidth      =   255
      TabIndex        =   2
      Top             =   60
      Width           =   255
   End
   Begin VB.PictureBox picHour 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   1
      Left            =   300
      Picture         =   "frmMain.frx":0B2E
      ScaleHeight     =   195
      ScaleWidth      =   255
      TabIndex        =   1
      Top             =   60
      Width           =   255
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   3570
      Top             =   1980
   End
   Begin VB.PictureBox picHour 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   0
      Left            =   60
      Picture         =   "frmMain.frx":0D68
      ScaleHeight     =   195
      ScaleWidth      =   255
      TabIndex        =   0
      Top             =   60
      Width           =   255
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00808080&
      X1              =   1470
      X2              =   1980
      Y1              =   6120
      Y2              =   6420
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00808080&
      X1              =   1350
      X2              =   1860
      Y1              =   6150
      Y2              =   6450
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   1110
      X2              =   1620
      Y1              =   6210
      Y2              =   6510
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   1230
      X2              =   1740
      Y1              =   6180
      Y2              =   6480
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const mFormat As String = "hhmmss"
Const mSkin As String = "Green"

Private Sub Form_DblClick()
  End
End Sub

Private Sub Form_Load()
  Call DisplayTime
  Timer1.Enabled = True
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Call FormDrag(Me)
End Sub

Private Sub Form_Resize()
  Line1.Y1 = 0
  Line1.Y2 = 0
  Line1.X1 = 0
  Line1.X2 = ScaleWidth
  
  Line2.Y1 = 0
  Line2.Y2 = ScaleHeight
  Line2.X1 = 0
  Line2.X2 = 0

  Line3.Y1 = ScaleHeight - 10
  Line3.Y2 = ScaleHeight - 10
  Line3.X1 = 0
  Line3.X2 = ScaleWidth - 10
  
  Line4.Y1 = 0
  Line4.Y2 = ScaleHeight - 10
  Line4.X1 = ScaleWidth - 10
  Line4.X2 = ScaleWidth - 10
End Sub

Private Sub picHour_DblClick(Index As Integer)
  End
End Sub

Private Sub picHour_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  Call FormDrag(Me)
End Sub

Private Sub picMin_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  Call FormDrag(Me)
End Sub

Private Sub picSec_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  Call FormDrag(Me)
End Sub

Private Sub picColon_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  Call FormDrag(Me)
End Sub

Private Sub picMin_DblClick(Index As Integer)
  End
End Sub

Private Sub picSec_DblClick(Index As Integer)
  End
End Sub

Private Sub picColon_DblClick(Index As Integer)
  End
End Sub

Private Sub Timer1_Timer()
  Call DisplayTime
End Sub

Sub DisplayTime()
  Dim T As String
  T = Format(Time, mFormat)
  
  picHour(0).Left = 30
  picHour(0).Picture = LoadPicture(App.Path & "\" & mSkin & "\" & Mid$(T, 1, 1) & ".gif")
  
  picHour(1).Left = picHour(0).Left + picHour(0).Width
  picHour(1).Picture = LoadPicture(App.Path & "\" & mSkin & "\" & Mid$(T, 2, 1) & ".gif")
  
  picColon(0).Left = picHour(1).Left + picHour(1).Width
  picColon(0).Picture = LoadPicture(App.Path & "\" & mSkin & "\colon.gif")
  
  picMin(0).Left = picColon(0).Left + picColon(0).Width
  picMin(0).Picture = LoadPicture(App.Path & "\" & mSkin & "\" & Mid$(T, 3, 1) & ".gif")
  
  picMin(1).Left = picMin(0).Left + picMin(0).Width
  picMin(1).Picture = LoadPicture(App.Path & "\" & mSkin & "\" & Mid$(T, 4, 1) & ".gif")
  
  picColon(1).Left = picMin(1).Left + picMin(1).Width
  picColon(1).Picture = LoadPicture(App.Path & "\" & mSkin & "\colon.gif")
  
  picSec(0).Left = picColon(1).Left + picColon(1).Width
  picSec(0).Picture = LoadPicture(App.Path & "\" & mSkin & "\" & Mid$(T, 5, 1) & ".gif")
  
  picSec(1).Left = picSec(0).Left + picSec(0).Width
  picSec(1).Picture = LoadPicture(App.Path & "\" & mSkin & "\" & Mid$(T, 6, 1) & ".gif")
  
  Dim W As Long
  Dim H As Long
  W = picSec(1).Left + picSec(1).Width + 60
  H = picHour(0).Height + 120
  If Width <> W Then Width = W
  If Height <> H Then Height = H
End Sub
