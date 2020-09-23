VERSION 5.00
Begin VB.Form frmMenu2 
   BackColor       =   &H80000009&
   Caption         =   "Jazz"
   ClientHeight    =   6765
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12720
   LinkTopic       =   "Form1"
   Picture         =   "Form3.frx":0000
   ScaleHeight     =   6765
   ScaleWidth      =   12720
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      BackColor       =   &H80000009&
      Caption         =   "Stop the MUSIC !"
      Height          =   1215
      Left            =   5880
      Picture         =   "Form3.frx":48526
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1440
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000009&
      Caption         =   "Play the MUSIC !"
      Height          =   615
      Left            =   2760
      Picture         =   "Form3.frx":49001
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   600
      Width           =   8355
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   2460
      Left            =   2520
      Picture         =   "Form3.frx":4ACF1
      ScaleHeight     =   2460
      ScaleWidth      =   3345
      TabIndex        =   0
      Top             =   4320
      Width           =   3345
   End
End
Attribute VB_Name = "frmMenu2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    PlaySound App.Path & "\Bond.wav"
End Sub

Private Sub Command2_Click()
    PlaySound App.Path & "\"
End Sub
