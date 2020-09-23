VERSION 5.00
Begin VB.Form frmMenu 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "All In One"
   ClientHeight    =   6825
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12780
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   6825
   ScaleMode       =   0  'User
   ScaleWidth      =   12093.31
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "Pooh"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   210
      Left            =   480
      TabIndex        =   7
      Top             =   6230
      Width           =   435
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "Picture 5"
      BeginProperty Font 
         Name            =   "Pooh"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   4830
      Width           =   855
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "Picture 4"
      BeginProperty Font 
         Name            =   "Pooh"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   4200
      Width           =   870
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "Picture 3"
      BeginProperty Font 
         Name            =   "Pooh"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   3550
      Width           =   855
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "Picture 2"
      BeginProperty Font 
         Name            =   "Pooh"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   2955
      Width           =   855
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "Picture 1"
      BeginProperty Font 
         Name            =   "Pooh"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   2385
      Width           =   840
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "ABOUT"
      BeginProperty Font 
         Name            =   "Pooh"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   210
      Left            =   360
      TabIndex        =   1
      Top             =   1320
      Width           =   600
   End
   Begin VB.Image imgRollOvr 
      Height          =   1455
      Left            =   5880
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Image imgButton 
      Height          =   300
      Left            =   120
      Top             =   0
      Width           =   1935
   End
   Begin VB.Label lblAbout 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Pooh"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   4245
      Left            =   2760
      TabIndex        =   0
      Top             =   1320
      Visible         =   0   'False
      Width           =   3690
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim LastButton As Integer




Private Sub Form_Load()
    LastButton = -1
    ButtonsOn = True

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim intButton As Integer
    
    
    intButton = CursorOnButton(x, y)
    
        
        If intButton > -1 Then
            If LastButton <> intButton Then
                lblAbout = sMenuButtonText(intButton)
                lblAbout.Visible = True
                imgButton.Move rectangle(intButton).Left, rectangle(intButton).Top
                imgButton.Picture = LoadPicture(sGraphics(intButton))
                imgButton.Visible = True
                imgButton.Enabled = True
                imgRollOvr.Picture = LoadPicture(sRollOvr(intButton))
                imgRollOvr.Visible = True
                LastButton = intButton
            End If
        Else
            lblAbout.Visible = False
            imgButton.Visible = False
            imgButton.Enabled = False
            imgRollOvr.Visible = False
            LastButton = -1
        End If
    
End Sub




Private Sub imgButton_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim i As Integer
        
        'This controls what happens when a menu button is clisked
        If PtInRegion(hRectRgn(0), x + imgButton.Left, y + imgButton.Top) Then
            frmAbout.Show vbModal
            
        ElseIf PtInRegion(hRectRgn(1), x + imgButton.Left, y + imgButton.Top) Then
            frmMenu1.Show
            
        ElseIf PtInRegion(hRectRgn(2), x + imgButton.Left, y + imgButton.Top) Then
            frmMenu2.Show
            
        ElseIf PtInRegion(hRectRgn(3), x + imgButton.Left, y + imgButton.Top) Then
            frmCursor.Show vbModal
            
        ElseIf PtInRegion(hRectRgn(4), x + imgButton.Left, y + imgButton.Top) Then
            MsgBox "Picture 4!"
            
        ElseIf PtInRegion(hRectRgn(5), x + imgButton.Left, y + imgButton.Top) Then
            MsgBox "Picture 5!"
            
        ElseIf PtInRegion(hRectRgn(6), x + imgButton.Left, y + imgButton.Top) Then
            If MsgBox("This is the way out ! If you want to quit clik OK !", vbOKCancel Or vbInformation) = vbOK Then
            
            MsgBox "      Thank you for using the ""ALL IN ONE"" program." & vbCrLf & vbCrLf _
            & "         All rights reserved,  Srdjan Josipovic - MCSD " & vbCrLf _
          & "                     srdjan.j@sezampro.yu"
            Unload Me
            End If
        End If
       
End Sub




Private Sub Label1_Click()
frmAbout.Show vbModal
End Sub

Private Sub Label2_Click()
frmMenu1.Show
End Sub

Private Sub Label3_Click()
frmMenu2.Show
End Sub

Private Sub Label4_Click()
frmCursor.Show vbModal
End Sub

Private Sub Label5_Click()
MsgBox "Picture 4!"
End Sub

Private Sub Label6_Click()
MsgBox "Picture 5!"
End Sub

Private Sub Label7_Click()
 
 If MsgBox("This is the way out ! If you want to quit clik OK !", vbOKCancel Or vbInformation) = vbOK Then
 
 MsgBox "      Thank you for using the ""ALL IN ONE"" program." & vbCrLf & vbCrLf _
            & "         All rights reserved,  Srdjan Josipovic - MCSD " & vbCrLf _
          & "                     srdjan.j@sezampro.yu"
            Unload Me
 End If
End Sub
