VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H80000009&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About All In One"
   ClientHeight    =   6795
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12750
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmAbout.frx":0000
   ScaleHeight     =   6795
   ScaleWidth      =   12750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Pooh"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   3480
      TabIndex        =   5
      Text            =   "ALL IN ONE"
      Top             =   360
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000009&
      Caption         =   "Start"
      BeginProperty Font 
         Name            =   "Pooh"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Timer Timer2 
      Interval        =   50
      Left            =   2760
      Top             =   5520
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   2760
      Top             =   5040
   End
   Begin VB.Label Label4 
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   "ALL IN ONE"
      BeginProperty Font 
         Name            =   "Pooh"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5400
      TabIndex        =   3
      Top             =   1680
      Width           =   7335
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000012&
      BackStyle       =   0  'Transparent
      Caption         =   "ALL IN ONE"
      BeginProperty Font 
         Name            =   "Pooh"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   615
      Left            =   2640
      TabIndex        =   2
      Top             =   2280
      Width           =   7695
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   "ALL IN ONE"
      BeginProperty Font 
         Name            =   "Pooh"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   735
      Left            =   2640
      TabIndex        =   1
      Top             =   3240
      Width           =   7695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "ALL IN ONE"
      BeginProperty Font 
         Name            =   "Pooh"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   735
      Left            =   4920
      TabIndex        =   0
      Top             =   2640
      Width           =   7455
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const Logo = "ALL IN ONE , ALL IN ONE , ALL IN ONE , ALL IN ONE  "
Public Z

Private Sub Command1_Click()
Text1.Text = "ALL IN ONE"



Command1.Visible = False
Text1.Text = "ALL IN ONE , ALL IN ONE ,  "
Text1.Text = Text1.Text + Chr(13) + Chr(10) + Chr(7)
DoLabel4
Label1.Caption = Logo
Label2.Caption = Logo
Label3.Caption = Logo
Timer2.Enabled = True
Timer1.Enabled = True
End Sub



Private Sub Form_Load()
Text1.Text = "ALL IN ONE , ALL IN ONE , ALL IN ONE , ALL IN ONE "
Randomize
Label4.Top = -500
Label2.Left = -2000
Label3.Left = -2000
End Sub

Private Sub Timer1_Timer()
Dim a As String
Label2.Left = Label2.Left + 200
Label3.Left = Label3.Left + 200
If Label2.Left = 2000 Then
a = InStr(1, Text1.Text, (Chr(13) + Chr(10) + Chr(7)), 0)
If a = 0 Then a = Len(Text1.Text)
Label1.Caption = Mid(Text1, 1, a)
Text1.Text = Mid(Text1.Text, a + 2, Len(Text1.Text))
End If
If Label2.Left > 8000 Then
Timer1.Enabled = False
Label2.Left = -3000
Label3.Left = -3000
a = InStr(1, Text1.Text, (Chr(13) + Chr(10) + Chr(7)), 0)
If a = 0 Then

Label1.Caption = ""
Label2.Caption = ""
Label3.Caption = ""
Label4.Caption = ""
Command1.Visible = True
Else
Timer2.Enabled = True
DoLabel4
Timer1.Enabled = True
End If
End If
End Sub

Private Sub Timer2_Timer()
Dim x, q
Label4.Top = Label4.Top + 120
Label4.Caption = ""
For x = 1 To Z
q = Int(Rnd * 13)
Label4.Caption = Label4.Caption + Mid(Logo, q + 1, 1)
Next x
If Label4.Top > 5000 Then
Label4.Top = -500
Timer2.Enabled = False
End If
End Sub

Private Sub DoLabel4()
Dim a As String
a = InStr(1, Text1.Text, (Chr(13) + Chr(10)), 0)
Z = Len(Mid(Text1, 1, a))
End Sub

