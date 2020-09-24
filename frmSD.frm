VERSION 5.00
Begin VB.Form frmSD 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Silly Drawing"
   ClientHeight    =   6195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9000
   Icon            =   "frmSD.frx":0000
   LinkTopic       =   "Silly Drawing"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6195
   ScaleWidth      =   9000
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton optFree 
      Caption         =   "FreeHand"
      Height          =   300
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton cmdAbout 
      Caption         =   "?"
      Height          =   375
      Left            =   6960
      TabIndex        =   25
      Top             =   5400
      Width           =   375
   End
   Begin VB.Frame fraMode 
      Caption         =   " Mode"
      Height          =   800
      Left            =   120
      TabIndex        =   22
      Top             =   1650
      Width           =   1215
      Begin VB.OptionButton optCom 
         Caption         =   "&Command"
         Height          =   195
         Left            =   120
         TabIndex        =   24
         Top             =   480
         Width           =   1020
      End
      Begin VB.OptionButton optPick 
         Caption         =   "&Pick Point"
         Height          =   195
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Value           =   -1  'True
         Width           =   1020
      End
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   375
      Left            =   5400
      TabIndex        =   19
      Top             =   5400
      Width           =   1000
   End
   Begin VB.Frame fraBack 
      Caption         =   " Background "
      Height          =   800
      Left            =   120
      TabIndex        =   18
      Top             =   4480
      Width           =   1215
      Begin VB.OptionButton optW 
         Caption         =   "White"
         Height          =   195
         Left            =   120
         TabIndex        =   21
         Top             =   480
         Width           =   855
      End
      Begin VB.OptionButton optB 
         Caption         =   "Black"
         Height          =   195
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Value           =   -1  'True
         Width           =   855
      End
   End
   Begin VB.Frame fraColor 
      Caption         =   " Color "
      Height          =   2000
      Left            =   120
      TabIndex        =   9
      Top             =   2450
      Width           =   1215
      Begin VB.OptionButton optYellow 
         Caption         =   "&Yellow"
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   1680
         Width           =   900
      End
      Begin VB.OptionButton optGreen 
         Caption         =   "Gree&n"
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   1440
         Width           =   900
      End
      Begin VB.OptionButton optRed 
         Caption         =   "&Red"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   1200
         Width           =   900
      End
      Begin VB.OptionButton optBlue 
         Caption         =   "&Blue"
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   960
         Width           =   900
      End
      Begin VB.OptionButton optWhite 
         Caption         =   "&White"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   720
         Value           =   -1  'True
         Width           =   900
      End
      Begin VB.OptionButton optGray 
         Caption         =   "&Gray"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   480
         Width           =   900
      End
      Begin VB.OptionButton optBlack 
         Caption         =   "Blac&k"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   900
      End
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   375
      Left            =   4320
      TabIndex        =   8
      Top             =   5400
      Width           =   1000
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   7920
      TabIndex        =   6
      Top             =   5400
      Width           =   1000
   End
   Begin VB.TextBox txtData 
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Text            =   "500"
      Top             =   5400
      Width           =   4095
   End
   Begin VB.PictureBox picBoard 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   5175
      Left            =   1560
      ScaleHeight     =   5145
      ScaleWidth      =   7335
      TabIndex        =   4
      Top             =   120
      Width           =   7365
   End
   Begin VB.OptionButton optRect 
      Caption         =   "Rectangular"
      Height          =   300
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1020
      Width           =   1215
   End
   Begin VB.OptionButton optTri 
      Caption         =   "Triangle"
      Height          =   300
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   720
      Width           =   1215
   End
   Begin VB.OptionButton optCirc 
      Caption         =   "Circle"
      Height          =   300
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   420
      Width           =   1215
   End
   Begin VB.OptionButton optLine 
      Caption         =   "Line"
      Height          =   300
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VB.Label lblST 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Pick 2 points to draw a line"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   5880
      Width           =   7575
   End
   Begin VB.Label lblPos 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0,0"
      Height          =   255
      Left            =   7720
      TabIndex        =   7
      Top             =   5880
      Width           =   1200
   End
End
Attribute VB_Name = "frmSD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Shp As Integer, pick As Integer, side As Integer, Colour, Mode As Integer
Public X1 As Integer, X2 As Integer, Y1 As Integer, Y2 As Integer, i As Integer, j As Integer

Private Sub cmdAbout_Click()
    MsgBox "Created by: Waleed Nassef, Egypt 2002", vbOKOnly, "Author"
End Sub

Private Sub cmdSave_Click()
frmSave.Show vbModal
End Sub

Private Sub cmdClear_Click()
picBoard.Cls
End Sub

Private Sub cmdExit_Click()
End
End Sub


Private Sub Form_Load()
For i = 0 To 7400 Step 100
Me.Line (1560 + i, 15)-(1560 + i, 90)
Next i

For i = 0 To 5000 Step 100
Me.Line (1450, 120 + i)-(1530, 120 + i)
Next i

End Sub


Private Sub optB_Click()
picBoard.BackColor = vbBlack
End Sub

Private Sub optCirc_Click()
lblST.Caption = "Specify the radius of the circle and pick the center point"
pick = 0
End Sub

Private Sub optFree_Click()
lblST.Caption = "Hold the left mouse key to draw"
pick = 0
End Sub

Private Sub optLine_Click()
lblST.Caption = "Pick 2 points to draw a line"
pick = 0
End Sub

Private Sub optRect_Click()
lblST.Caption = "Pick 2 corner points to draw a rectangular"
pick = 0
End Sub

Private Sub optTri_Click()
lblST.Caption = "Specify the length of the side and pick the top angle"
pick = 0
End Sub

Private Sub optW_Click()
picBoard.BackColor = vbWhite
End Sub

Private Sub picBoard_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
picBoard.CurrentX = X
picBoard.CurrentY = Y
End Sub

Private Sub picBoard_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
GetModes

lblPos.Caption = Str(X) + "," + Str(Y)
If optFree.Value = True Then
If Button = 1 Then
picBoard.Line (picBoard.CurrentX, picBoard.CurrentY)-(X, Y), Colour
End If
End If
End Sub

Private Sub picBoard_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

GetModes

Select Case Mode
Case 1
    Select Case Shp
    Case 1
    If pick = 0 Then
    X1 = X: Y1 = Y: pick = 1
    Else: X2 = X: Y2 = Y: pick = 0
    picBoard.Line (X1, Y1)-(X2, Y2), Colour
    End If
    Case 2
    picBoard.Circle (X, Y), Val(txtData.Text), Colour
    Case 3
    side = Val(txtData.Text)
    picBoard.Line (X, Y)-(X - side / 2, Int(Y + 3 ^ 0.5 / 2 * side)), Colour
    picBoard.Line (X, Y)-(X + side / 2, Int(Y + 3 ^ 0.5 / 2 * side)), Colour
    picBoard.Line (X - side / 2, Int(Y + 3 ^ 0.5 / 2 * side))-(X + side / 2, Int(Y + 3 ^ 0.5 / 2 * side)), Colour
    Case 4
    If pick = 0 Then
    X1 = X: Y1 = Y: pick = 1
    Else: X2 = X: Y2 = Y: pick = 0
    picBoard.Line (X1, Y1)-(X2, Y2), Colour, B
    End If
    Case 5
    
    End Select
    
Case 2
MsgBox "Comming soon !", vbOKOnly, "Beta Version"
End Select

End Sub

Private Sub txtdata_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then Exit Sub
If KeyAscii < 48 Or KeyAscii > 57 Then
KeyAscii = 0
End If
End Sub


Public Sub GetModes()
If optPick.Value = True Then Mode = 1
If optCom.Value = True Then Mode = 2

If optLine.Value = True Then Shp = 1
If optCirc.Value = True Then Shp = 2
If optTri.Value = True Then Shp = 3
If optRect.Value = True Then Shp = 4
If optFree.Value = True Then Shp = 5

If optBlack.Value = True Then Colour = RGB(0, 0, 0)
If optGray.Value = True Then Colour = RGB(125, 125, 125)
If optWhite.Value = True Then Colour = RGB(255, 255, 255)
If optBlue.Value = True Then Colour = RGB(0, 0, 255)
If optRed.Value = True Then Colour = RGB(255, 0, 0)
If optGreen.Value = True Then Colour = RGB(0, 255, 0)
If optYellow.Value = True Then Colour = RGB(255, 255, 0)

End Sub
