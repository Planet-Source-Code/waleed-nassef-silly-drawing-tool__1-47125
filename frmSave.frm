VERSION 5.00
Begin VB.Form frmSave 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Save Picture (BMP) Format"
   ClientHeight    =   2355
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5535
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2355
   ScaleWidth      =   5535
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4320
      TabIndex        =   4
      Top             =   1800
      Width           =   1095
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Default         =   -1  'True
      Height          =   375
      Left            =   4320
      TabIndex        =   3
      Top             =   1320
      Width           =   1095
   End
   Begin VB.TextBox txtFileName 
      Height          =   375
      Left            =   2160
      TabIndex        =   0
      Text            =   "silly"
      Top             =   600
      Width           =   3255
   End
   Begin VB.DriveListBox Drive 
      Height          =   315
      Left            =   2160
      TabIndex        =   2
      Top             =   1800
      Width           =   1695
   End
   Begin VB.DirListBox Dir 
      Height          =   2115
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "File Name (without extension)"
      Height          =   195
      Left            =   2160
      TabIndex        =   5
      Top             =   240
      Width           =   2070
   End
End
Attribute VB_Name = "frmSave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim FilePath As String

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    If Len(Dir.Path) = 3 Then
        FilePath = Dir.Path + txtFileName.Text + ".bmp"
    Else
        FilePath = Dir.Path + "\" + txtFileName.Text + ".bmp"
    End If

    SavePicture frmSD.picBoard.Image, FilePath
    Unload Me
End Sub

Private Sub Drive_Change()
On Error GoTo out
    Dir.Path = Drive.Drive
out:
    Exit Sub
End Sub
