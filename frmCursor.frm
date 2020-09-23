VERSION 5.00
Begin VB.Form frmCursor 
   BackColor       =   &H80000009&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cursor"
   ClientHeight    =   6795
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12720
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmCursor.frx":0000
   ScaleHeight     =   6795
   ScaleWidth      =   12720
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   480
      Top             =   6120
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000E&
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2520
      TabIndex        =   0
      Top             =   6120
      Width           =   10215
   End
End
Attribute VB_Name = "frmCursor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Timer1_Timer()
Dim x As Long, y As Long
    
    If CursorPos(x, y) Then
        Label1.Caption = "Cursor position (pixels) = " & x & " , " & y
    End If
End Sub
